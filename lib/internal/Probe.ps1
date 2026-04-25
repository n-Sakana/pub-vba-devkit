param(
    [ValidateSet('Interactive', 'Basic', 'Extended', 'GenerateStorage')]
    [string]$Mode = 'Interactive',

    [ValidateSet('Local', 'OneDrive_Synced')]
    [string]$Scenario = 'Local',

    [string]$OutputDirectory,

    [string]$ProbeReportName = 'probe.txt',

    [string]$ReturnTextFile
)

$ErrorActionPreference = 'Stop'

# ============================================================================
# Environment Probe - Tests EDR/compat patterns via Excel COM
# Creates temporary xlsm files, injects VBA code, tests save/run/result
# ============================================================================

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$results = [System.Collections.ArrayList]::new()
$tempDir = Join-Path ([IO.Path]::GetTempPath()) "vba-probe-$([guid]::NewGuid().ToString('N').Substring(0,8))"
New-Item $tempDir -ItemType Directory -Force | Out-Null
$script:ItemTimeoutSeconds = 90

function Test-SkipRequested {
    # Check keyboard input (requires TreatControlCAsInput = $true)
    # Ctrl+C = skip current item, Ctrl+Esc = abort entire probe
    while ([Console]::KeyAvailable) {
        $key = [Console]::ReadKey($true)
        if ($key.Key -eq [ConsoleKey]::Escape -and ($key.Modifiers -band [ConsoleModifiers]::Control)) {
            Write-Host "`n  >> Ctrl+Esc: aborting probe..." -ForegroundColor Red
            throw [System.Management.Automation.PipelineStoppedException]::new('User abort (Ctrl+Esc)')
        }
        if ($key.Key -eq [ConsoleKey]::C -and ($key.Modifiers -band [ConsoleModifiers]::Control)) {
            Write-Host "`n  >> Ctrl+C: skipping current item..." -ForegroundColor Yellow
            return $true
        }
    }
    return $false
}

function Add-ProbeResult {
    param([string]$Level, [string]$Category, [string]$Pattern, [string]$Target,
          [string]$Result, [string]$Phase = '', [int]$ErrNum = 0, [string]$ErrMsg = '', [string]$Detail = '')
    [void]$script:results.Add([ordered]@{
        Level = $Level; Category = $Category; Pattern = $Pattern; Target = $Target
        Result = $Result; Phase = $Phase; ErrNum = $ErrNum; ErrMsg = $ErrMsg; Detail = $Detail
    })
    $color = switch ($Result) { 'OK' { 'Green' } 'FAIL' { 'Red' } 'SKIP' { 'DarkGray' } default { 'Yellow' } }
    Write-Host "  [$Result] $Pattern - $Target $(if($Phase){"($Phase) "})$Detail" -ForegroundColor $color
}

function Mask-ProbePath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return '(empty)'
    }

    if ($Path -match '^(?<scheme>https?://)(?<rest>.+)$') {
        $parts = @($matches['rest'] -split '/' | Where-Object { $_ })
        $masked = ($parts | ForEach-Object { '***' }) -join '/'
        return "$($matches['scheme'])$masked(depth=$($parts.Count))"
    }

    if ($Path.StartsWith('\\')) {
        $parts = @($Path.Substring(2) -split '\\' | Where-Object { $_ })
        $masked = ($parts | ForEach-Object { '***' }) -join '\'
        return "UNC:\\$masked(depth=$($parts.Count))"
    }

    $drive = ''
    $rest = $Path
    if ($Path.Length -ge 3 -and $Path[1] -eq ':') {
        $drive = $Path.Substring(0, 2) + '\'
        $rest = $Path.Substring(3)
    }

    $segments = @($rest -split '\\' | Where-Object { $_ })
    $maskedLocal = ($segments | ForEach-Object { '***' }) -join '\'
    return "Local:$drive$maskedLocal(depth=$($segments.Count))"
}

function Get-ProbeOneDriveEnvRows {
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($name in @('OneDriveCommercial', 'OneDrive', 'OneDriveConsumer')) {
        $value = [Environment]::GetEnvironmentVariable($name)
        $exists = $false
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            try {
                $exists = Test-Path -LiteralPath $value
            } catch {
                $exists = $false
            }
        }

        $rows.Add([ordered]@{
            Name   = $name
            Value  = $value
            Exists = $exists
        }) | Out-Null
    }

    return @($rows)
}

function Get-ProbePreferredSyncRoot {
    $rows = @(Get-ProbeOneDriveEnvRows)
    foreach ($preferred in @('OneDriveCommercial', 'OneDrive', 'OneDriveConsumer')) {
        $row = @($rows | Where-Object { $_.Name -eq $preferred -and -not [string]::IsNullOrWhiteSpace($_.Value) } | Select-Object -First 1)[0]
        if ($row) {
            return $row
        }
    }

    return $null
}

function Test-HostAction {
    param(
        [string]$Category,
        [string]$Pattern,
        [string]$Target,
        [scriptblock]$Action,
        [hashtable]$Variables = @{}
    )

    # Run action in a separate runspace so we can enforce a preemptive timeout.
    # Types loaded via Add-Type (e.g. [ProbeHostNative]) are AppDomain-wide
    # and available in any runspace within the same process.
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    foreach ($name in @('Get-HostWindowTarget', 'Get-WindowSummary', 'Mask-ProbePath', 'Get-ProbeOneDriveEnvRows', 'Get-ProbePreferredSyncRoot')) {
        $func = Get-Command $name -ErrorAction SilentlyContinue
        if ($func -and $func.CommandType -eq 'Function') {
            $iss.Commands.Add(
                [System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($name, $func.ScriptBlock.ToString())
            )
        }
    }

    $rs = [runspacefactory]::CreateRunspace($iss)
    $rs.Open()
    foreach ($key in $Variables.Keys) {
        $rs.SessionStateProxy.SetVariable($key, $Variables[$key])
    }

    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($Action.ToString())

    $handle = $ps.BeginInvoke()
    $deadline = [DateTime]::UtcNow.AddSeconds($script:ItemTimeoutSeconds)
    $completed = $false
    $skipped = $false

    # Poll with 500ms intervals: check completion, timeout, and Ctrl+C
    while ([DateTime]::UtcNow -lt $deadline) {
        if ($handle.AsyncWaitHandle.WaitOne(500)) {
            $completed = $true
            break
        }
        if (Test-SkipRequested) {
            $skipped = $true
            break
        }
    }

    if ($completed) {
        try {
            $output = $ps.EndInvoke($handle)
            if ($ps.HadErrors) {
                $errMsg = ($ps.Streams.Error[0]).Exception.Message
                Add-ProbeResult 'Extended' $Category $Pattern $Target 'FAIL' 'Host' 0 $errMsg
            } else {
                $detail = if ($output.Count -gt 0) { [string]$output[-1] } else { '' }
                Add-ProbeResult 'Extended' $Category $Pattern $Target 'OK' 'Host' 0 '' $detail
            }
        } catch {
            Add-ProbeResult 'Extended' $Category $Pattern $Target 'FAIL' 'Host' 0 $_.Exception.Message
        }
        $ps.Dispose()
        $rs.Dispose()
    } elseif ($skipped) {
        # Ctrl+C: non-blocking stop, don't wait for cleanup
        $ps.BeginStop($null, $null) | Out-Null
        Add-ProbeResult 'Extended' $Category $Pattern $Target 'SKIP' 'Host' 0 '' 'Skipped by user (Ctrl+C)'
    } else {
        # Timeout: non-blocking stop, don't wait for cleanup
        $ps.BeginStop($null, $null) | Out-Null
        Add-ProbeResult 'Extended' $Category $Pattern $Target 'TIMEOUT' 'Host' 0 '' "Exceeded ${script:ItemTimeoutSeconds}s"
    }
}

function Initialize-HostNativeBridge {
    if ('ProbeHostNative' -as [type]) {
        Add-ProbeResult 'Extended' 'Host / PInvoke' 'PowerShell / C# Add-Type' 'DllImport user32/kernel32' 'OK' 'Host' 0 '' 'Already loaded'
        return $true
    }

    try {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class ProbeHostNative
{
    [DllImport("kernel32.dll", ExactSpelling = true)]
    public static extern ulong GetTickCount64();

    [DllImport("kernel32.dll", ExactSpelling = true)]
    public static extern uint GetCurrentThreadId();

    [DllImport("kernel32.dll", ExactSpelling = true)]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll", ExactSpelling = true)]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", ExactSpelling = true)]
    public static extern IntPtr GetDesktopWindow();

    [DllImport("user32.dll", ExactSpelling = true)]
    public static extern IntPtr GetShellWindow();

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr FindWindowW(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr FindWindowExW(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int GetWindowTextW(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int GetClassNameW(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool MessageBeep(uint uType);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr SendMessageW(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr SendMessageTimeoutW(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool PostMessageW(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
}
'@ -Language CSharp -ErrorAction Stop

        Add-ProbeResult 'Extended' 'Host / PInvoke' 'PowerShell / C# Add-Type' 'DllImport user32/kernel32' 'OK' 'Host' 0 '' 'Compiled'
        return $true
    } catch {
        Add-ProbeResult 'Extended' 'Host / PInvoke' 'PowerShell / C# Add-Type' 'DllImport user32/kernel32' 'FAIL' 'Host' 0 $_.Exception.Message
        return $false
    }
}

function Get-HostWindowTarget {
    $handle = [ProbeHostNative]::GetForegroundWindow()
    if ($handle -ne [IntPtr]::Zero) {
        return [ordered]@{
            Handle = $handle
            Source = 'GetForegroundWindow'
        }
    }

    $handle = [ProbeHostNative]::FindWindowW('Shell_TrayWnd', $null)
    if ($handle -ne [IntPtr]::Zero) {
        return [ordered]@{
            Handle = $handle
            Source = 'FindWindowW(Shell_TrayWnd)'
        }
    }

    $handle = [ProbeHostNative]::GetConsoleWindow()
    if ($handle -ne [IntPtr]::Zero) {
        return [ordered]@{
            Handle = $handle
            Source = 'GetConsoleWindow'
        }
    }

    throw 'No target window handle found.'
}

function Get-WindowSummary {
    param([IntPtr]$Handle)

    $title = New-Object System.Text.StringBuilder 260
    $class = New-Object System.Text.StringBuilder 260
    $titleChars = [ProbeHostNative]::GetWindowTextW($Handle, $title, $title.Capacity)
    $titleError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
    $classChars = [ProbeHostNative]::GetClassNameW($Handle, $class, $class.Capacity)
    $classError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
    $processId = [uint32]0
    $threadId = [ProbeHostNative]::GetWindowThreadProcessId($Handle, [ref]$processId)
    $threadError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
    $visible = [ProbeHostNative]::IsWindowVisible($Handle)
    $visibleError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
    $isWindow = [ProbeHostNative]::IsWindow($Handle)
    $windowError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()

    return "Handle=0x{0}; TitleChars={1}; TitleErr={2}; Title={3}; ClassChars={4}; ClassErr={5}; Class={6}; ThreadId={7}; ProcessId={8}; ThreadErr={9}; Visible={10}; VisibleErr={11}; IsWindow={12}; IsWindowErr={13}" -f `
        $Handle.ToInt64(), $titleChars, $titleError, $title.ToString(), $classChars, $classError, $class.ToString(), $threadId, $processId, $threadError, $visible, $visibleError, $isWindow, $windowError
}

function Test-HostWindowAutomation {
    if (-not (Initialize-HostNativeBridge)) {
        return
    }

    Test-HostAction 'Host / PInvoke' 'PowerShell / PInvoke call' 'GetTickCount64' {
        "Ticks=$([ProbeHostNative]::GetTickCount64())"
    }

    Test-HostAction 'Host / PInvoke' 'PowerShell / PInvoke call' 'GetCurrentThreadId' {
        "ThreadId=$([ProbeHostNative]::GetCurrentThreadId())"
    }

    Test-HostAction 'Host / PInvoke' 'PowerShell / PInvoke call' 'MessageBeep' {
        $ok = [ProbeHostNative]::MessageBeep([uint32]::MaxValue)
        $lastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        "Result=$ok; LastError=$lastError"
    }

    Test-HostAction 'Host / Window' 'Window handle lookup' 'GetForegroundWindow / FindWindowW / GetConsoleWindow' {
        $target = Get-HostWindowTarget
        "Handle=0x{0}; Source={1}" -f $target.Handle.ToInt64(), $target.Source
    }

    Test-HostAction 'Host / Window' 'Desktop window handle' 'GetDesktopWindow' {
        $handle = [ProbeHostNative]::GetDesktopWindow()
        Get-WindowSummary -Handle $handle
    }

    Test-HostAction 'Host / Window' 'Shell window handle' 'GetShellWindow' {
        $handle = [ProbeHostNative]::GetShellWindow()
        Get-WindowSummary -Handle $handle
    }

    Test-HostAction 'Host / Window' 'FindWindowW' 'Shell_TrayWnd' {
        $handle = [ProbeHostNative]::FindWindowW('Shell_TrayWnd', $null)
        $lastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        "Handle=0x{0}; LastError={1}" -f $handle.ToInt64(), $lastError
    }

    Test-HostAction 'Host / Window' 'FindWindowExW' 'Shell_TrayWnd child' {
        $parent = [ProbeHostNative]::FindWindowW('Shell_TrayWnd', $null)
        $handle = [ProbeHostNative]::FindWindowExW($parent, [IntPtr]::Zero, $null, $null)
        $lastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        "Parent=0x{0}; Child=0x{1}; LastError={2}" -f $parent.ToInt64(), $handle.ToInt64(), $lastError
    }

    Test-HostAction 'Host / Window' 'Window class name' 'GetClassNameW' {
        $target = Get-HostWindowTarget
        $sb = New-Object System.Text.StringBuilder 260
        $chars = [ProbeHostNative]::GetClassNameW($target.Handle, $sb, $sb.Capacity)
        $lastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        "Handle=0x{0}; Source={1}; Chars={2}; LastError={3}; Class={4}" -f $target.Handle.ToInt64(), $target.Source, $chars, $lastError, $sb.ToString()
    }

    Test-HostAction 'Host / Window' 'Window caption' 'GetWindowTextW' {
        $target = Get-HostWindowTarget
        $sb = New-Object System.Text.StringBuilder 260
        $chars = [ProbeHostNative]::GetWindowTextW($target.Handle, $sb, $sb.Capacity)
        $lastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        "Handle=0x{0}; Source={1}; Chars={2}; LastError={3}; Title={4}" -f $target.Handle.ToInt64(), $target.Source, $chars, $lastError, $sb.ToString()
    }

    Test-HostAction 'Host / Window' 'Window metadata' 'IsWindow / IsWindowVisible / GetWindowThreadProcessId' {
        $target = Get-HostWindowTarget
        "Source=$($target.Source); $(Get-WindowSummary -Handle $target.Handle)"
    }

    Test-HostAction 'Host / Window' 'Foreground activation' 'SetForegroundWindow(current)' {
        $target = Get-HostWindowTarget
        $result = [ProbeHostNative]::SetForegroundWindow($target.Handle)
        $lastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        "Handle=0x{0}; Source={1}; Result={2}; LastError={3}" -f $target.Handle.ToInt64(), $target.Source, $result, $lastError
    }

    Test-HostAction 'Host / Window' 'Window show state' 'ShowWindow(SW_SHOWNORMAL)' {
        $target = Get-HostWindowTarget
        $result = [ProbeHostNative]::ShowWindow($target.Handle, 1)
        $lastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        "Handle=0x{0}; Source={1}; Result={2}; LastError={3}" -f $target.Handle.ToInt64(), $target.Source, $result, $lastError
    }

    Test-HostAction 'Host / Message' 'SendMessage' 'WM_NULL' {
        $target = Get-HostWindowTarget
        $result = [ProbeHostNative]::SendMessageW($target.Handle, 0, [IntPtr]::Zero, [IntPtr]::Zero)
        $lastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        "Handle=0x{0}; Source={1}; Result=0x{2}; LastError={3}" -f $target.Handle.ToInt64(), $target.Source, $result.ToInt64().ToString('X'), $lastError
    }

    Test-HostAction 'Host / Message' 'SendMessageTimeout' 'WM_NULL' {
        $target = Get-HostWindowTarget
        $resultValue = [IntPtr]::Zero
        $result = [ProbeHostNative]::SendMessageTimeoutW($target.Handle, 0, [IntPtr]::Zero, [IntPtr]::Zero, 2, 500, [ref]$resultValue)
        $lastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        "Handle=0x{0}; Source={1}; Return=0x{2}; Result=0x{3}; LastError={4}" -f $target.Handle.ToInt64(), $target.Source, $result.ToInt64().ToString('X'), $resultValue.ToInt64().ToString('X'), $lastError
    }

    Test-HostAction 'Host / Message' 'PostMessage' 'WM_NULL' {
        $target = Get-HostWindowTarget
        $result = [ProbeHostNative]::PostMessageW($target.Handle, 0, [IntPtr]::Zero, [IntPtr]::Zero)
        $lastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        "Handle=0x{0}; Source={1}; Result={2}; LastError={3}" -f $target.Handle.ToInt64(), $target.Source, $result, $lastError
    }

    Test-HostAction 'Host / UIAutomation' '.NET UIAutomation' 'AutomationElement.RootElement' {
        Add-Type -AssemblyName UIAutomationClient -ErrorAction Stop
        $root = [System.Windows.Automation.AutomationElement]::RootElement
        if ($null -eq $root) {
            throw 'RootElement was null.'
        }
        $name = $root.Current.Name
        $controlType = $root.Current.ControlType.ProgrammaticName
        "Name=$name; ControlType=$controlType"
    }

    Test-HostAction 'Host / UIAutomation' '.NET UIAutomation' 'AutomationElement.FromHandle(foreground)' {
        Add-Type -AssemblyName UIAutomationClient -ErrorAction Stop
        $target = Get-HostWindowTarget
        $element = [System.Windows.Automation.AutomationElement]::FromHandle($target.Handle)
        if ($null -eq $element) {
            throw 'AutomationElement.FromHandle returned null.'
        }

        $children = $element.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)
        "Handle=0x{0}; Source={1}; Name={2}; ControlType={3}; Children={4}" -f `
            $target.Handle.ToInt64(), $target.Source, $element.Current.Name, $element.Current.ControlType.ProgrammaticName, $children.Count
    }

    Test-HostAction 'Host / WScript' 'WScript.Shell' 'CreateObject' {
        $wsh = New-Object -ComObject WScript.Shell
        try {
            "Type=$($wsh.GetType().FullName)"
        } finally {
            if ($wsh) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wsh) }
        }
    }

    Test-HostAction 'Host / WScript' 'WScript.Shell AppActivate' 'Foreground window process' {
        $target = Get-HostWindowTarget
        [uint32]$processId = 0
        [void][ProbeHostNative]::GetWindowThreadProcessId($target.Handle, [ref]$processId)
        $wsh = New-Object -ComObject WScript.Shell
        try {
            $result = $wsh.AppActivate([int]$processId)
            "Handle=0x{0}; Source={1}; ProcessId={2}; Result={3}" -f $target.Handle.ToInt64(), $target.Source, $processId, $result
        } finally {
            if ($wsh) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wsh) }
        }
    }

    Test-HostAction 'Host / WScript' 'WScript.Shell SendKeys' 'empty string' {
        $target = Get-HostWindowTarget
        [uint32]$processId = 0
        [void][ProbeHostNative]::GetWindowThreadProcessId($target.Handle, [ref]$processId)
        $wsh = New-Object -ComObject WScript.Shell
        try {
            [void]$wsh.AppActivate([int]$processId)
            Start-Sleep -Milliseconds 100
            $wsh.SendKeys('')
            "ProcessId=$processId; Sent=(empty)"
        } finally {
            if ($wsh) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wsh) }
        }
    }

    Test-HostAction 'Host / Clipboard' 'Windows Forms Clipboard' 'ContainsText / ContainsFileDropList' {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $containsText = [System.Windows.Forms.Clipboard]::ContainsText()
        $containsFiles = [System.Windows.Forms.Clipboard]::ContainsFileDropList()
        "ContainsText=$containsText; ContainsFileDropList=$containsFiles"
    }

    Test-HostAction 'Host / IPC' 'NamedPipe server' 'NamedPipeServerStream' {
        $pipeName = "vba-probe-$([guid]::NewGuid().ToString('N').Substring(0,8))"
        $server = [System.IO.Pipes.NamedPipeServerStream]::new($pipeName, [System.IO.Pipes.PipeDirection]::InOut, 1, [System.IO.Pipes.PipeTransmissionMode]::Byte, [System.IO.Pipes.PipeOptions]::Asynchronous)
        try {
            "Pipe=$pipeName"
        } finally {
            $server.Dispose()
        }
    }

    Test-HostAction 'Host / IPC' 'NamedPipe roundtrip' 'NamedPipeServerStream + NamedPipeClientStream' {
        $pipeName = "vba-probe-$([guid]::NewGuid().ToString('N').Substring(0,8))"
        $server = [System.IO.Pipes.NamedPipeServerStream]::new($pipeName, [System.IO.Pipes.PipeDirection]::InOut, 1, [System.IO.Pipes.PipeTransmissionMode]::Byte, [System.IO.Pipes.PipeOptions]::Asynchronous)
        $client = [System.IO.Pipes.NamedPipeClientStream]::new('.', $pipeName, [System.IO.Pipes.PipeDirection]::InOut)
        try {
            $serverTask = $server.WaitForConnectionAsync()
            $client.Connect(1000)
            $serverTask.Wait(1000) | Out-Null
            $writer = [System.IO.StreamWriter]::new($client)
            $writer.AutoFlush = $true
            $reader = [System.IO.StreamReader]::new($server)
            $writer.WriteLine('probe')
            $line = $reader.ReadLine()
            "Pipe=$pipeName; Message=$line"
        } finally {
            if ($writer) { $writer.Dispose() }
            if ($reader) { $reader.Dispose() }
            $client.Dispose()
            $server.Dispose()
        }
    }

    Test-HostAction 'Host / IPC' 'TCP listener' 'System.Net.Sockets.TcpListener' {
        $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
        try {
            $listener.Start()
            $endpoint = [System.Net.IPEndPoint]$listener.LocalEndpoint
            "Endpoint=127.0.0.1:$($endpoint.Port)"
        } finally {
            $listener.Stop()
        }
    }

    Test-HostAction 'Host / IPC' 'TCP loopback roundtrip' 'TcpListener + TcpClient' {
        $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
        try {
            $listener.Start()
            $endpoint = [System.Net.IPEndPoint]$listener.LocalEndpoint
            $acceptTask = $listener.AcceptTcpClientAsync()
            $client = [System.Net.Sockets.TcpClient]::new()
            $client.Connect('127.0.0.1', $endpoint.Port)
            $serverClient = $acceptTask.Result
            try {
                $clientWriter = [System.IO.StreamWriter]::new($client.GetStream())
                $clientWriter.AutoFlush = $true
                $serverReader = [System.IO.StreamReader]::new($serverClient.GetStream())
                $clientWriter.WriteLine('probe')
                $line = $serverReader.ReadLine()
                "Endpoint=127.0.0.1:$($endpoint.Port); Message=$line"
            } finally {
                if ($clientWriter) { $clientWriter.Dispose() }
                if ($serverReader) { $serverReader.Dispose() }
                $serverClient.Dispose()
                $client.Dispose()
            }
        } finally {
            $listener.Stop()
        }
    }

    Test-HostAction 'Host / IPC' 'HttpListener' 'http://127.0.0.1:0/' {
        $port = Get-Random -Minimum 20000 -Maximum 45000
        $listener = [System.Net.HttpListener]::new()
        try {
            $listener.Prefixes.Add("http://127.0.0.1:$port/")
            $listener.Start()
            "Prefix=http://127.0.0.1:$port/"
        } finally {
            if ($listener.IsListening) {
                $listener.Stop()
            }
            $listener.Close()
        }
    }

    Test-HostAction 'Host / IPC' 'HttpListener loopback' 'HttpListener + WebRequest' {
        $port = Get-Random -Minimum 20000 -Maximum 45000
        $listener = [System.Net.HttpListener]::new()
        try {
            $prefix = "http://127.0.0.1:$port/"
            $listener.Prefixes.Add($prefix)
            $listener.Start()
            $serverTask = $listener.GetContextAsync()
            $request = [System.Net.WebRequest]::Create($prefix)
            $response = $request.GetResponse()
            try {
                $context = $serverTask.Result
                try {
                    $writer = [System.IO.StreamWriter]::new($context.Response.OutputStream)
                    $writer.Write('ok')
                    $writer.Flush()
                } finally {
                    if ($writer) { $writer.Dispose() }
                    $context.Response.Close()
                }
                "Prefix=$prefix; Status=$([int]$response.StatusCode)"
            } finally {
                $response.Dispose()
            }
        } finally {
            if ($listener.IsListening) {
                $listener.Stop()
            }
            $listener.Close()
        }
    }
}

function Test-HostStoragePaths {
    foreach ($row in @(Get-ProbeOneDriveEnvRows)) {
        if ([string]::IsNullOrWhiteSpace($row.Value)) {
            Add-ProbeResult 'Extended' 'Host / Storage' 'OneDrive env' $row.Name 'SKIP' 'Host' 0 '' 'Not set'
            continue
        }

        Add-ProbeResult 'Extended' 'Host / Storage' 'OneDrive env' $row.Name 'OK' 'Host' 0 '' ("Value={0}; Exists={1}" -f (Mask-ProbePath -Path $row.Value), $row.Exists)
    }

    $root = Get-ProbePreferredSyncRoot
    if (-not $root) {
        Add-ProbeResult 'Extended' 'Host / Storage' 'Local sync root enumeration' 'Get-ChildItem / Directory.GetFiles' 'SKIP' 'Host' 0 '' 'No OneDrive environment variable was set.'
        return
    }

    if (-not $root.Exists) {
        Add-ProbeResult 'Extended' 'Host / Storage' 'Local sync root enumeration' $root.Name 'FAIL' 'Host' 0 '' ("Configured path not found: {0}" -f (Mask-ProbePath -Path $root.Value))
        return
    }

    Add-ProbeResult 'Extended' 'Host / Storage' 'Preferred sync root' $root.Name 'OK' 'Host' 0 '' ("Path={0}" -f (Mask-ProbePath -Path $root.Value))

    Test-HostAction 'Host / Storage' 'Get-ChildItem' $root.Name -Variables @{ rootPath = $root.Value } -Action {
        $entries = @(Get-ChildItem -LiteralPath $rootPath -Force -ErrorAction Stop | Select-Object -First 5)
        $preview = if ($entries.Count -gt 0) {
            ($entries | ForEach-Object { "{0}:{1}" -f $_.PSIsContainer, (Mask-ProbePath -Path $_.FullName) }) -join '; '
        } else {
            '(empty)'
        }
        "Entries=$($entries.Count); Preview=$preview"
    }

    Test-HostAction 'Host / Storage' '.NET Directory.GetFiles' $root.Name -Variables @{ rootPath = $root.Value } -Action {
        $files = [System.IO.Directory]::GetFiles($rootPath)
        $dirs = [System.IO.Directory]::GetDirectories($rootPath)
        $firstFile = if ($files.Length -gt 0) { Mask-ProbePath -Path $files[0] } else { '(none)' }
        $firstDir = if ($dirs.Length -gt 0) { Mask-ProbePath -Path $dirs[0] } else { '(none)' }
        "Files=$($files.Length); Directories=$($dirs.Length); FirstFile=$firstFile; FirstDirectory=$firstDir"
    }
}

# ============================================================================
# Test runner: create xlsm, inject code, save, optionally run, clean up
# ============================================================================

function Test-VbaCode {
    param(
        [string]$Level,
        [string]$Category,
        [string]$Pattern,
        [string]$Target,
        [string]$VbaCode,
        [string]$RunMacro = '',     # macro name to execute after save (empty = save-only test)
        [switch]$ExpectSaveFail     # if true, save failure = expected (e.g. EDR blocks Declare)
    )

    $testFile = Join-Path $tempDir "probe_$([guid]::NewGuid().ToString('N').Substring(0,8)).xlsm"
    $excel = $null
    $wb = $null
    $excelPid = $null
    $beforePids = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $itemSw = [System.Diagnostics.Stopwatch]::StartNew()
    $phase = 'Setup'

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false

        # Track the Excel PID we just created for cleanup on timeout
        $afterPids = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $excelPid = @($afterPids | Where-Object { $_ -notin $beforePids }) | Select-Object -First 1

        if (Test-SkipRequested) { throw [System.OperationCanceledException]::new("$phase") }
        if ($itemSw.Elapsed.TotalSeconds -ge $script:ItemTimeoutSeconds) { throw [System.TimeoutException]::new("$phase") }

        $wb = $excel.Workbooks.Add()

        # Inject VBA code
        $phase = 'Inject'
        try {
            $mod = $wb.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
            $mod.Name = 'ProbeTest'
            $mod.CodeModule.AddFromString($VbaCode)
        } catch {
            if ($itemSw.Elapsed.TotalSeconds -ge $script:ItemTimeoutSeconds) { throw [System.TimeoutException]::new("$phase") }
            Add-ProbeResult $Level $Category $Pattern $Target 'FAIL' 'Inject' 0 $_.Exception.Message
            return
        }

        if (Test-SkipRequested) { throw [System.OperationCanceledException]::new("$phase") }
        if ($itemSw.Elapsed.TotalSeconds -ge $script:ItemTimeoutSeconds) { throw [System.TimeoutException]::new("$phase") }

        # Save
        $phase = 'Save'
        try {
            $wb.SaveAs($testFile, 52)  # xlOpenXMLWorkbookMacroEnabled
        } catch {
            if ($itemSw.Elapsed.TotalSeconds -ge $script:ItemTimeoutSeconds) { throw [System.TimeoutException]::new("$phase") }
            if ($ExpectSaveFail) {
                Add-ProbeResult $Level $Category $Pattern $Target 'BLOCKED' 'Save' 0 $_.Exception.Message 'EDR blocked as expected'
            } else {
                Add-ProbeResult $Level $Category $Pattern $Target 'FAIL' 'Save' 0 $_.Exception.Message
            }
            return
        }

        if (Test-SkipRequested) { throw [System.OperationCanceledException]::new("$phase") }
        if ($itemSw.Elapsed.TotalSeconds -ge $script:ItemTimeoutSeconds) { throw [System.TimeoutException]::new("$phase") }

        # Close and reopen to check if EDR corrupts the file on open
        $phase = 'Reopen'
        $wb.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
        $wb = $null

        try {
            $wb = $excel.Workbooks.Open($testFile, 0, $false)
            # Try accessing VBProject to verify file integrity
            $compCount = $wb.VBProject.VBComponents.Count
        } catch {
            if ($itemSw.Elapsed.TotalSeconds -ge $script:ItemTimeoutSeconds) { throw [System.TimeoutException]::new("$phase") }
            if ($ExpectSaveFail) {
                Add-ProbeResult $Level $Category $Pattern $Target 'BLOCKED' 'Reopen' 0 $_.Exception.Message 'EDR corrupted file on save/open'
            } else {
                Add-ProbeResult $Level $Category $Pattern $Target 'FAIL' 'Reopen' 0 $_.Exception.Message 'File corrupted on reopen'
            }
            return
        }

        if (Test-SkipRequested) { throw [System.OperationCanceledException]::new("$phase") }
        if ($itemSw.Elapsed.TotalSeconds -ge $script:ItemTimeoutSeconds) { throw [System.TimeoutException]::new("$phase") }

        if ($ExpectSaveFail) {
            # Save + reopen both succeeded: EDR did not block this pattern
            Add-ProbeResult $Level $Category $Pattern $Target 'OK' 'Save+Reopen' 0 '' 'EDR did not block'
        }

        # Run macro if specified — interpret return value
        if ($RunMacro) {
            $phase = 'Run'
            try {
                $runResult = [string]($excel.Run($RunMacro))
                if ($itemSw.Elapsed.TotalSeconds -ge $script:ItemTimeoutSeconds) { throw [System.TimeoutException]::new("$phase") }
                if ($runResult.StartsWith('OK')) {
                    Add-ProbeResult $Level $Category $Pattern $Target 'OK' 'Run' 0 '' $runResult
                } else {
                    # Macro returned FAIL: or unexpected value
                    Add-ProbeResult $Level $Category $Pattern $Target 'FAIL' 'Run' 0 $runResult 'Macro reported failure'
                }
            } catch {
                if ($_.Exception -is [System.TimeoutException] -or $_.Exception -is [System.OperationCanceledException]) { throw }
                if ($itemSw.Elapsed.TotalSeconds -ge $script:ItemTimeoutSeconds) {
                    throw [System.TimeoutException]::new("$phase")
                }
                Add-ProbeResult $Level $Category $Pattern $Target 'FAIL' 'Run' 0 $_.Exception.Message
            }
        } elseif (-not $ExpectSaveFail) {
            Add-ProbeResult $Level $Category $Pattern $Target 'OK' 'Save+Reopen' 0 '' 'Code accepted'
        }

    } catch {
        if ($_.Exception -is [System.OperationCanceledException]) {
            Add-ProbeResult $Level $Category $Pattern $Target 'SKIP' $phase 0 '' 'Skipped by user (Ctrl+C)'
        } elseif ($_.Exception -is [System.TimeoutException] -or $itemSw.Elapsed.TotalSeconds -ge $script:ItemTimeoutSeconds) {
            Add-ProbeResult $Level $Category $Pattern $Target 'TIMEOUT' $phase 0 '' "Exceeded ${script:ItemTimeoutSeconds}s"
        } else {
            Add-ProbeResult $Level $Category $Pattern $Target 'FAIL' 'Setup' 0 $_.Exception.Message
        }
    } finally {
        try { if ($wb) { $wb.Close($false) } } catch {}
        try { if ($excel) { $excel.Quit() } } catch {}
        if ($wb) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
        if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        Remove-Item $testFile -Force -ErrorAction SilentlyContinue
        # Force kill orphaned Excel process on timeout
        if ($excelPid) {
            try {
                $proc = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
                if ($proc -and -not $proc.HasExited) {
                    Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
                }
            } catch {}
        }
    }
}


# ============================================================================
# Generate SharePoint Storage Probe xlsm
# ============================================================================

function Generate-StorageProbe {
    $scriptRoot = Split-Path $PSScriptRoot -Parent
    $outDir = if ($OutputDirectory) { $OutputDirectory } else { Join-Path $scriptRoot 'output' }
    if (-not (Test-Path $outDir)) { New-Item $outDir -ItemType Directory -Force | Out-Null }
    $outPath = Join-Path $outDir 'probe_storage.xlsm'

    Write-Host "Generating probe_storage.xlsm..." -ForegroundColor Cyan

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    try {
        $wb = $excel.Workbooks.Add()
        $mod = $wb.VBProject.VBComponents.Add(1)
        $mod.Name = 'StorageProbe'

        $vbaCode = @'
Option Explicit

Private Type ProbeResult
    TestName As String
    Result As String
    Detail As String
End Type

Private m_results() As ProbeResult
Private m_count As Long

Private Function MaskPath(p As String) As String
    ' Mask a path to hide directory names but keep type and depth identifiable.
    ' Examples:
    '   C:\Users\xxx\Documents\folder        -> Local:C:\***\***\***\***(depth=4)
    '   https://company.sharepoint.com/a/b/c  -> URL:https://***/***/***/***(depth=4)
    '   \\server\share\folder                 -> UNC:\\***\***\***(depth=3)
    Dim prefix As String
    Dim sep As String
    Dim parts() As String
    Dim depth As Long

    If Len(p) = 0 Then
        MaskPath = "(empty)"
        Exit Function
    End If

    If Left(p, 8) = "https://" Or Left(p, 7) = "http://" Then
        prefix = "URL:"
        Dim scheme As String
        If Left(p, 8) = "https://" Then
            scheme = "https://"
        Else
            scheme = "http://"
        End If
        Dim rest As String
        rest = Mid(p, Len(scheme) + 1)
        parts = Split(rest, "/")
        depth = 0
        Dim urlMasked As String
        urlMasked = scheme
        Dim j As Long
        For j = LBound(parts) To UBound(parts)
            If Len(parts(j)) > 0 Then
                depth = depth + 1
                If depth > 1 Then urlMasked = urlMasked & "/"
                urlMasked = urlMasked & "***"
            End If
        Next j
        MaskPath = prefix & urlMasked & "(depth=" & depth & ")"
    ElseIf Left(p, 2) = "\\" Then
        prefix = "UNC:\\"
        rest = Mid(p, 3)
        parts = Split(rest, "\")
        depth = 0
        Dim uncMasked As String
        uncMasked = ""
        For j = LBound(parts) To UBound(parts)
            If Len(parts(j)) > 0 Then
                depth = depth + 1
                If depth > 1 Then uncMasked = uncMasked & "\"
                uncMasked = uncMasked & "***"
            End If
        Next j
        MaskPath = prefix & uncMasked & "(depth=" & depth & ")"
    Else
        ' Local path: extract drive letter
        Dim drive As String
        If Mid(p, 2, 1) = ":" Then
            drive = Left(p, 2) & "\"
            rest = Mid(p, 4)
        Else
            drive = ""
            rest = p
        End If
        parts = Split(rest, "\")
        depth = 0
        Dim localMasked As String
        localMasked = ""
        For j = LBound(parts) To UBound(parts)
            If Len(parts(j)) > 0 Then
                depth = depth + 1
                If depth > 1 Then localMasked = localMasked & "\"
                localMasked = localMasked & "***"
            End If
        Next j
        MaskPath = "Local:" & drive & localMasked & "(depth=" & depth & ")"
    End If
End Function

Private Function SafeEnv(name As String) As String
    SafeEnv = Trim$(Environ$(name))
End Function

Private Function GetPreferredSyncRoot(ByRef sourceName As String) As String
    Dim names As Variant
    Dim i As Long
    Dim value As String

    names = Array("OneDriveCommercial", "OneDrive", "OneDriveConsumer")
    sourceName = ""

    For i = LBound(names) To UBound(names)
        value = SafeEnv(CStr(names(i)))
        If Len(value) > 0 Then
            sourceName = CStr(names(i))
            GetPreferredSyncRoot = value
            Exit Function
        End If
    Next i
End Function

Private Function EnsureSyncRoot(scenario As String, sourceName As String, root As String) As Boolean
    If Len(root) > 0 Then
        EnsureSyncRoot = True
        Exit Function
    End If

    If scenario = "OneDrive_Synced" Or scenario = "SharePoint_OpenInApp" Then
        AddResult "Local Sync Root", "FAIL", "NoOneDriveEnv"
    Else
        AddResult "Local Sync Root", "OK", "NoOneDriveEnv"
    End If
End Function

Private Sub AddResult(name As String, result As String, detail As String)
    m_count = m_count + 1
    If m_count > UBound(m_results) Then ReDim Preserve m_results(1 To m_count + 10)
    m_results(m_count).TestName = name
    m_results(m_count).Result = result
    m_results(m_count).Detail = detail
End Sub

Public Sub Probe_Run()
    ReDim m_results(1 To 40)
    m_count = 0

    Dim scenario As String
    scenario = InputBox("Scenario:" & vbCrLf & "1 = Local" & vbCrLf & "2 = SharePoint Open-in-App" & vbCrLf & "3 = OneDrive Synced", "Storage Probe", "2")
    Select Case scenario
        Case "1": scenario = "Local"
        Case "2": scenario = "SharePoint_OpenInApp"
        Case "3": scenario = "OneDrive_Synced"
        Case Else: scenario = "Unknown"
    End Select

    AddResult "Scenario", "OK", scenario
    AddResult "Office Version", "OK", Application.Version
    #If Win64 Then
        AddResult "Office Bitness", "OK", "64-bit"
    #Else
        AddResult "Office Bitness", "OK", "32-bit"
    #End If

    ' Path observations (masked to hide confidential directory names)
    AddResult "ThisWorkbook.Path", "OK", MaskPath(ThisWorkbook.Path)
    AddResult "ThisWorkbook.FullName", "OK", MaskPath(ThisWorkbook.FullName)

    On Error Resume Next
    AddResult "CurDir", "OK", MaskPath(CurDir)
    Err.Clear

    ' Path type
    Dim p As String: p = ThisWorkbook.FullName
    Dim pathType As String
    If Left(p, 5) = "https" Then
        pathType = "URL"
    ElseIf InStr(p, "OneDrive") > 0 Then
        pathType = "OneDrive"
    ElseIf Left(p, 2) = "\\" Then
        pathType = "UNC"
    Else
        pathType = "Local"
    End If
    AddResult "Path Type", "OK", pathType

    ' Relative path
    Dim firstFile As String
    firstFile = Dir(ThisWorkbook.Path & "\*.*")
    If Err.Number <> 0 Then
        AddResult "Relative Path (Dir)", "FAIL", Err.Description
        Err.Clear
    Else
        AddResult "Relative Path (Dir)", "OK", "FirstFile=" & firstFile
    End If

    Dim commercialRoot As String
    commercialRoot = SafeEnv("OneDriveCommercial")
    If Len(commercialRoot) = 0 Then
        AddResult "OneDriveCommercial Env", "FAIL", "Empty"
    Else
        AddResult "OneDriveCommercial Env", "OK", MaskPath(commercialRoot)
    End If

    If Len(commercialRoot) = 0 Then
        AddResult "OneDriveCommercial Dir", "FAIL", "Empty"
    Else
        firstFile = Dir(commercialRoot & "\*.*")
        If Err.Number <> 0 Then
            AddResult "OneDriveCommercial Dir", "FAIL", Err.Description
            Err.Clear
        ElseIf Len(firstFile) = 0 Then
            AddResult "OneDriveCommercial Dir", "OK", "(empty)"
        Else
            AddResult "OneDriveCommercial Dir", "OK", "FirstEntry=" & firstFile
        End If
    End If

    AddResult "OneDrive Env Vars", "OK", _
        "OneDriveCommercial=" & MaskPath(SafeEnv("OneDriveCommercial")) & _
        "; OneDrive=" & MaskPath(SafeEnv("OneDrive")) & _
        "; OneDriveConsumer=" & MaskPath(SafeEnv("OneDriveConsumer"))

    Dim syncSource As String
    Dim syncRoot As String
    syncRoot = GetPreferredSyncRoot(syncSource)
    If EnsureSyncRoot(scenario, syncSource, syncRoot) Then
        AddResult "Local Sync Root", "OK", "Source=" & syncSource & "; Root=" & MaskPath(syncRoot)

        firstFile = Dir(syncRoot & "\*.*")
        If Err.Number <> 0 Then
            AddResult "Local Sync Root (Dir)", "FAIL", Err.Description
            Err.Clear
        ElseIf Len(firstFile) = 0 Then
            AddResult "Local Sync Root (Dir)", "OK", "Source=" & syncSource & "; FirstEntry=(empty)"
        Else
            AddResult "Local Sync Root (Dir)", "OK", "Source=" & syncSource & "; FirstEntry=" & firstFile
        End If

        Dim firstWorkbook As String
        firstWorkbook = Dir(syncRoot & "\*.xls*")
        If Err.Number <> 0 Then
            AddResult "Local Sync Root (Dir *.xls*)", "FAIL", Err.Description
            Err.Clear
        ElseIf Len(firstWorkbook) = 0 Then
            AddResult "Local Sync Root (Dir *.xls*)", "OK", "Source=" & syncSource & "; FirstWorkbook=(none at root)"
        Else
            AddResult "Local Sync Root (Dir *.xls*)", "OK", "Source=" & syncSource & "; FirstWorkbook=" & firstWorkbook
        End If

        Dim syncFso As Object
        Dim syncFolder As Object
        Set syncFso = CreateObject("Scripting.FileSystemObject")
        Set syncFolder = syncFso.GetFolder(syncRoot)
        If Err.Number <> 0 Then
            AddResult "Local Sync Root (FSO)", "FAIL", Err.Description
            Err.Clear
        Else
            AddResult "Local Sync Root (FSO)", "OK", "Source=" & syncSource & "; Files=" & syncFolder.Files.Count & "; SubFolders=" & syncFolder.SubFolders.Count
        End If
    End If

    ' AutoSave
    Dim autoSave As Boolean
    autoSave = ThisWorkbook.AutoSaveOn
    If Err.Number <> 0 Then
        AddResult "AutoSave", "OK", "Not supported"
        Err.Clear
    Else
        AddResult "AutoSave", "OK", CStr(autoSave)
    End If

    ' === Workbooks.Open: resolution path detail ===

    ' Test 1: Open via ThisWorkbook.Path (the main concern)
    Dim tmpPath As String
    tmpPath = ThisWorkbook.Path & "\probe_open_test.xlsx"
    AddResult "Open: input path", "OK", MaskPath(tmpPath)

    Dim tmpWb As Workbook
    Set tmpWb = Application.Workbooks.Add
    tmpWb.SaveAs tmpPath, 51
    If Err.Number <> 0 Then
        AddResult "Open: create adjacent", "FAIL", Err.Description
        Err.Clear
    Else
        tmpWb.Close False
        Dim openedWb As Workbook
        Set openedWb = Application.Workbooks.Open(tmpPath)
        If Err.Number <> 0 Then
            AddResult "Open: via TWB.Path", "FAIL", Err.Description
            Err.Clear
        Else
            ' Record how Excel resolved the path
            AddResult "Open: resolved Path", "OK", MaskPath(openedWb.Path)
            AddResult "Open: resolved FullName", "OK", MaskPath(openedWb.FullName)
            openedWb.Close False
        End If
        On Error Resume Next
        Kill tmpPath
        Err.Clear
    End If

    ' Test 2: Open via Environ TEMP (always local)
    Dim tmpPath2 As String
    tmpPath2 = Environ$("TEMP") & "\probe_open_local.xlsx"
    Set tmpWb = Application.Workbooks.Add
    tmpWb.SaveAs tmpPath2, 51
    tmpWb.Close False
    If Err.Number <> 0 Then
        AddResult "Open: local TEMP", "FAIL", Err.Description
        Err.Clear
    Else
        Set openedWb = Application.Workbooks.Open(tmpPath2)
        If Err.Number <> 0 Then
            AddResult "Open: local TEMP", "FAIL", Err.Description
            Err.Clear
        Else
            AddResult "Open: local TEMP", "OK", MaskPath(openedWb.FullName)
            openedWb.Close False
        End If
        Kill tmpPath2
        If Err.Number <> 0 Then Err.Clear
    End If

    ' === SaveAs: where does it actually save? ===

    ' SaveAs to TEMP
    Dim savePath As String
    savePath = Environ$("TEMP") & "\probe_saveas_" & Format(Now, "yyyymmddhhnnss") & ".xlsm"
    Dim saveWb As Workbook
    Set saveWb = Application.Workbooks.Add
    saveWb.SaveAs savePath, 52
    If Err.Number <> 0 Then
        AddResult "SaveAs: to TEMP", "FAIL", Err.Description
        Err.Clear
    Else
        ' Check path AFTER save
        AddResult "SaveAs: to TEMP", "OK", MaskPath(saveWb.FullName)
        saveWb.Close False
        Kill savePath
        If Err.Number <> 0 Then Err.Clear
    End If

    ' SaveAs to same folder as ThisWorkbook
    Dim savePath2 As String
    savePath2 = ThisWorkbook.Path & "\probe_saveas_here.xlsm"
    Set saveWb = Application.Workbooks.Add
    saveWb.SaveAs savePath2, 52
    If Err.Number <> 0 Then
        AddResult "SaveAs: to TWB.Path", "FAIL", Err.Description
        Err.Clear
    Else
        AddResult "SaveAs: to TWB.Path", "OK", MaskPath(saveWb.FullName)
        saveWb.Close False
        Kill savePath2
        If Err.Number <> 0 Then Err.Clear
    End If

    ' === External link test ===

    ' Create a workbook with a formula linking to this workbook
    Dim linkWb As Workbook
    Set linkWb = Application.Workbooks.Add
    Dim linkFormula As String
    ' Reference cell A1 of Sheet1 in THIS workbook
    linkFormula = "='" & ThisWorkbook.FullName & "'!Sheet1!A1"
    linkWb.Sheets(1).Cells(1, 1).Value = "source"  ' put something in A1 first
    ThisWorkbook.Sheets(1).Cells(1, 1).Value = "probe_link_test"

    linkWb.Sheets(1).Cells(2, 1).Formula = "=[" & ThisWorkbook.Name & "]Sheet1!A1"
    If Err.Number <> 0 Then
        AddResult "ExtLink: formula set", "FAIL", Err.Description
        Err.Clear
    Else
        Dim linkVal As String
        linkVal = CStr(linkWb.Sheets(1).Cells(2, 1).Value)
        If linkVal = "probe_link_test" Then
            AddResult "ExtLink: resolve", "OK", "Value=" & linkVal
        Else
            AddResult "ExtLink: resolve", "FAIL", "Expected probe_link_test, got: " & linkVal
        End If
    End If

    ' Check link sources
    Dim linkSources As Variant
    linkSources = linkWb.LinkSources(1)  ' xlExcelLinks
    If Err.Number <> 0 Then
        AddResult "ExtLink: sources", "FAIL", Err.Description
        Err.Clear
    Else
        If IsArray(linkSources) Then
            AddResult "ExtLink: sources", "OK", MaskPath(CStr(linkSources(1)))
        Else
            AddResult "ExtLink: sources", "OK", "No array returned"
        End If
    End If
    linkWb.Close False

    ' Clean up
    ThisWorkbook.Sheets(1).Cells(1, 1).Value = ""

    ' COM tests (safe ones only, no Declare)
    TestCOM "Scripting.FileSystemObject"
    TestCOM "Scripting.Dictionary"
    TestCOM "ADODB.Connection"
    TestCOM "MSXML2.XMLHTTP.6.0"

    On Error GoTo 0

    ' Output
    WriteResults scenario
    MsgBox "Storage Probe complete." & vbCrLf & m_count & " tests run.", vbInformation
End Sub

Private Sub TestCOM(progId As String)
    On Error Resume Next
    Dim o As Object: Set o = CreateObject(progId)
    If Err.Number <> 0 Then
        AddResult "COM: " & progId, "FAIL", Err.Description
    Else
        AddResult "COM: " & progId, "OK", TypeName(o)
    End If
    Set o = Nothing
    Err.Clear
End Sub

Private Sub WriteResults(scenario As String)
    ' Write results to a scenario-named sheet (accumulates across runs)
    Dim sheetName As String: sheetName = scenario
    If Len(sheetName) > 31 Then sheetName = Left(sheetName, 31)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    ws.Cells(1, 1).Value = "Storage Probe: " & scenario
    ws.Cells(2, 1).Value = "Date: " & Format(Now, "yyyy-mm-dd hh:nn:ss")

    ws.Cells(4, 1).Value = "TestName"
    ws.Cells(4, 2).Value = "Result"
    ws.Cells(4, 3).Value = "Detail"
    ws.Range("A4:C4").Font.Bold = True

    Dim i As Long
    For i = 1 To m_count
        ws.Cells(4 + i, 1).Value = m_results(i).TestName
        ws.Cells(4 + i, 2).Value = m_results(i).Result
        ws.Cells(4 + i, 3).Value = m_results(i).Detail
        If m_results(i).Result = "OK" Then
            ws.Cells(4 + i, 2).Interior.Color = RGB(200, 255, 200)
        ElseIf m_results(i).Result = "FAIL" Then
            ws.Cells(4 + i, 2).Interior.Color = RGB(255, 200, 200)
        End If
    Next i

    ws.Columns("A:C").AutoFit
    ws.Activate
    MsgBox scenario & ": " & m_count & " tests written to sheet.", vbInformation
End Sub

' === Export all scenario sheets to text files ===
Public Sub Probe_Export()
    Dim outFolder As String
    outFolder = ThisWorkbook.Path & "\probe-export"
    On Error Resume Next
    MkDir outFolder
    On Error GoTo 0

    Dim exported As Long
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Skip default sheets (Sheet1 etc)
        If Left(ws.Cells(1, 1).Value, 14) = "Storage Probe:" Then
            Dim scenario As String
            scenario = Mid(ws.Cells(1, 1).Value, 16)

            Dim outPath As String
            outPath = outFolder & "\storage_" & scenario & ".txt"

            Dim f As Long: f = FreeFile
            Open outPath For Output As #f
            Print #f, ws.Cells(1, 1).Value
            Print #f, ws.Cells(2, 1).Value
            Print #f, ""

            Dim lastRow As Long
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            Dim r As Long
            For r = 4 To lastRow
                Print #f, ws.Cells(r, 1).Value & vbTab & ws.Cells(r, 2).Value & vbTab & ws.Cells(r, 3).Value
            Next r
            Close #f
            exported = exported + 1
        End If
    Next ws

    If exported > 0 Then
        MsgBox exported & " file(s) exported to:" & vbCrLf & outFolder, vbInformation, "Export"
    Else
        MsgBox "No probe result sheets found.", vbExclamation
    End If
End Sub
'@

        $mod.CodeModule.AddFromString($vbaCode)
        $wb.SaveAs($outPath, 52)
        $wb.Close($false)

        Write-Host "Generated: $outPath" -ForegroundColor Green
        Write-Host ""
        Write-Host "Usage:" -ForegroundColor Gray
        Write-Host "  1. Upload probe_storage.xlsm to SharePoint" -ForegroundColor Gray
        Write-Host "  2. Open in App (not browser)" -ForegroundColor Gray
        Write-Host "  3. Alt+F8 > Probe_Run" -ForegroundColor Gray
        Write-Host "  4. Select scenario" -ForegroundColor Gray
        Write-Host "  5. Results saved as text file" -ForegroundColor Gray
    } finally {
        if ($wb) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
}

# ============================================================================
# Mode selection
# ============================================================================

Write-Host "=== Environment Probe ===" -ForegroundColor Cyan
Write-Host ""

$runExtended = $false
$scenarioLabel = $Scenario

if ($Mode -eq 'Interactive') {
    $modeInput = Read-Host "Run mode? (B=Basic only, E=Basic+Extended, G=Generate SharePoint probe xlsm, Q=Quit)"
    if ($modeInput -eq 'Q' -or $modeInput -eq 'q') { exit 0 }

    if ($modeInput -eq 'G' -or $modeInput -eq 'g') {
        Generate-StorageProbe
        exit 0
    }

    $runExtended = ($modeInput -eq 'E' -or $modeInput -eq 'e')

    Write-Host ""
    Write-Host "Scenario (how is this file opened?):"
    Write-Host "  1 = Local (desktop, local drive)"
    Write-Host "  2 = OneDrive Synced (local path via sync)"
    $scenarioNum = Read-Host "Scenario (1/2)"
    $scenarioLabel = switch ($scenarioNum) {
        '1' { 'Local' }
        '2' { 'OneDrive_Synced' }
        default { 'Local' }
    }
} else {
    if ($Mode -eq 'GenerateStorage') {
        Write-Host "Mode: Generate SharePoint probe xlsm" -ForegroundColor Gray
        Write-Host ""
        Generate-StorageProbe
        exit 0
    }

    $runExtended = ($Mode -eq 'Extended')
    $scenarioLabel = $Scenario
}

Write-Host ""
Write-Host "Mode: $(if($runExtended){'Basic + Extended'}else{'Basic only'}) | Scenario: $scenarioLabel" -ForegroundColor Gray
Write-Host "  Ctrl+C = skip item | Ctrl+Esc = abort" -ForegroundColor DarkGray
Write-Host ""
$scenario = $scenarioLabel

# Capture Ctrl+C as input so we can handle skip/abort ourselves
[Console]::TreatControlCAsInput = $true

# ============================================================================
# System Info
# ============================================================================

Write-Host "--- System Info ---" -ForegroundColor Cyan
# Mask personal info in output
$maskedPC = '(masked)'
$maskedUser = '(masked)'
Add-ProbeResult 'Aux' 'SystemInfo' 'Computer' 'Environ' 'OK' '' 0 '' $maskedPC
Add-ProbeResult 'Aux' 'SystemInfo' 'User' 'Environ' 'OK' '' 0 '' $maskedUser

# Get Office info via COM
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $ver = $xl.Version
    $build = $xl.Build
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Add-ProbeResult 'Aux' 'SystemInfo' 'Office Version' 'Excel.Application' 'OK' '' 0 '' "Version: $ver Build: $build"
} catch {
    Add-ProbeResult 'Aux' 'SystemInfo' 'Office Version' 'Excel.Application' 'FAIL' '' 0 $_.Exception.Message
}

# ============================================================================
# Basic Tests
# ============================================================================

# All tests inject VBA into a temp xlsm, save, reopen, and run.
# This ensures we test the actual VBA runtime behavior, not PowerShell COM.

# ----------------------------------------------------------------------------
# 1. VBA Baseline — confirms VBA itself works before testing anything else
# ----------------------------------------------------------------------------

Write-Host ""
Write-Host "--- VBA Baseline ---" -ForegroundColor Cyan

Test-VbaCode 'Basic' 'Baseline' 'VBA Baseline' 'Pure VBA (no external)' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    Dim x As Long: x = 1 + 2
    RunTest = "OK:" & x
End Function
'@

# ----------------------------------------------------------------------------
# 2. EDR / Security — Win32 API, COM, File I/O, Registry, Clipboard, Shell
# ----------------------------------------------------------------------------

Write-Host ""
Write-Host "--- EDR / Security Tests ---" -ForegroundColor Cyan

# Win32 API Declare (call)
Test-VbaCode 'Basic' 'EDR' 'Win32 API (Declare)' 'Declare PtrSafe Function' -ExpectSaveFail -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Public Function RunTest() As String
    RunTest = "OK:" & GetTickCount()
End Function
'@

# Win32 API Declare (no call — does mere presence break the file?)
Test-VbaCode 'Basic' 'EDR' 'Win32 API (Declare)' 'Declare only (no call)' -ExpectSaveFail -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Public Function RunTest() As String
    RunTest = "OK:DeclarePresent"
End Function
'@

# COM / CreateObject
Test-VbaCode 'Basic' 'EDR' 'COM / CreateObject' 'Scripting.FileSystemObject' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    Dim o As Object: Set o = CreateObject("Scripting.FileSystemObject")
    RunTest = "OK:" & o.GetTempName()
End Function
'@

Test-VbaCode 'Basic' 'EDR' 'COM / CreateObject' 'Scripting.Dictionary' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    Dim o As Object: Set o = CreateObject("Scripting.Dictionary")
    o.Add "k", "v"
    RunTest = "OK:" & o.Count
End Function
'@

Test-VbaCode 'Basic' 'EDR' 'COM / CreateObject' 'ADODB.Connection' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    Dim o As Object: Set o = CreateObject("ADODB.Connection")
    RunTest = "OK:" & TypeName(o)
End Function
'@

Test-VbaCode 'Basic' 'EDR' 'COM / CreateObject' 'ADODB.Recordset' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim o As Object: Set o = CreateObject("ADODB.Recordset")
    If Err.Number <> 0 Then RunTest = "FAIL:" & Err.Description Else RunTest = "OK:" & TypeName(o)
End Function
'@

Test-VbaCode 'Basic' 'EDR' 'COM / CreateObject' 'MSXML2.XMLHTTP.6.0' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    Dim o As Object: Set o = CreateObject("MSXML2.XMLHTTP.6.0")
    RunTest = "OK:" & TypeName(o)
End Function
'@

Test-VbaCode 'Basic' 'EDR' 'COM / CreateObject' 'MSXML2.DOMDocument.6.0' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim o As Object: Set o = CreateObject("MSXML2.DOMDocument.6.0")
    If Err.Number <> 0 Then RunTest = "FAIL:" & Err.Description Else RunTest = "OK:" & TypeName(o)
End Function
'@

Test-VbaCode 'Basic' 'EDR' 'COM / CreateObject' 'Shell.Application' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim o As Object: Set o = CreateObject("Shell.Application")
    If Err.Number <> 0 Then RunTest = "FAIL:" & Err.Description Else RunTest = "OK:" & TypeName(o)
End Function
'@

Test-VbaCode 'Basic' 'EDR' 'COM / CreateObject' 'WinHttp.WinHttpRequest.5.1' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim o As Object: Set o = CreateObject("WinHttp.WinHttpRequest.5.1")
    If Err.Number <> 0 Then RunTest = "FAIL:" & Err.Description Else RunTest = "OK:" & TypeName(o)
End Function
'@

# File I/O
Test-VbaCode 'Basic' 'EDR' 'File I/O' 'Open/Write/Kill' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    Dim p As String: p = Environ$("TEMP") & "\probe_" & Format(Now, "yyyymmddhhnnss") & ".txt"
    Dim f As Long: f = FreeFile
    Open p For Output As #f: Print #f, "test": Close #f
    Kill p
    RunTest = "OK"
End Function
'@

# FileSystemObject methods
Test-VbaCode 'Basic' 'EDR' 'FileSystemObject' 'FSO methods' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    RunTest = "OK:" & fso.FileExists(Environ$("TEMP") & "\nonexistent.xyz")
End Function
'@

# Registry
Test-VbaCode 'Basic' 'EDR' 'Registry' 'GetSetting' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    Dim v As String: v = GetSetting("ProbeTest", "S", "K", "default")
    RunTest = "OK:" & v
End Function
'@

# Environment
Test-VbaCode 'Basic' 'EDR' 'Environment' 'Environ$' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    RunTest = "OK:" & Len(Environ$("USERNAME"))
End Function
'@

# Clipboard
Test-VbaCode 'Basic' 'EDR' 'Clipboard' 'MSForms.DataObject' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim d As Object
    Set d = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    If Err.Number <> 0 Then
        RunTest = "FAIL:" & Err.Description
    Else
        d.SetText "probe": d.PutInClipboard
        RunTest = "OK"
    End If
End Function
'@

# Shell / process (create only, no Run)
Test-VbaCode 'Basic' 'EDR' 'Shell / process' 'WScript.Shell (create only)' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        RunTest = "FAIL:CreateObject:" & Err.Description
    Else
        RunTest = "OK:CreateObject"
    End If
End Function
'@

# ----------------------------------------------------------------------------
# 3. Workbook Context — path observations, open/save behavior, references
# ----------------------------------------------------------------------------

Write-Host ""
Write-Host "--- Workbook Context Tests ---" -ForegroundColor Cyan

# ThisWorkbook.Path observation (masked)
Test-VbaCode 'Basic' 'Info' 'ThisWorkbook.Path' 'Path value' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Private Function MaskPath(p As String) As String
    Dim prefix As String, sep As String, parts() As String, rest As String
    Dim depth As Long, masked As String, j As Long, scheme As String, drive As String
    If Len(p) = 0 Then MaskPath = "(empty)": Exit Function
    If Left(p, 8) = "https://" Or Left(p, 7) = "http://" Then
        If Left(p, 8) = "https://" Then scheme = "https://" Else scheme = "http://"
        rest = Mid(p, Len(scheme) + 1): parts = Split(rest, "/"): masked = scheme
        For j = LBound(parts) To UBound(parts)
            If Len(parts(j)) > 0 Then depth = depth + 1: If depth > 1 Then masked = masked & "/": masked = masked & "***"
        Next j
        MaskPath = "URL:" & masked & "(depth=" & depth & ")"
    ElseIf Left(p, 2) = "\\" Then
        rest = Mid(p, 3): parts = Split(rest, "\")
        For j = LBound(parts) To UBound(parts)
            If Len(parts(j)) > 0 Then depth = depth + 1: If depth > 1 Then masked = masked & "\": masked = masked & "***"
        Next j
        MaskPath = "UNC:\\" & masked & "(depth=" & depth & ")"
    Else
        If Mid(p, 2, 1) = ":" Then drive = Left(p, 2) & "\": rest = Mid(p, 4) Else drive = "": rest = p
        parts = Split(rest, "\")
        For j = LBound(parts) To UBound(parts)
            If Len(parts(j)) > 0 Then depth = depth + 1: If depth > 1 Then masked = masked & "\": masked = masked & "***"
        Next j
        MaskPath = "Local:" & drive & masked & "(depth=" & depth & ")"
    End If
End Function
Public Function RunTest() As String
    RunTest = "OK:" & MaskPath(ThisWorkbook.Path)
End Function
'@

# ThisWorkbook.FullName observation (masked)
Test-VbaCode 'Basic' 'Info' 'ThisWorkbook.FullName' 'FullName value' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Private Function MaskPath(p As String) As String
    Dim prefix As String, sep As String, parts() As String, rest As String
    Dim depth As Long, masked As String, j As Long, scheme As String, drive As String
    If Len(p) = 0 Then MaskPath = "(empty)": Exit Function
    If Left(p, 8) = "https://" Or Left(p, 7) = "http://" Then
        If Left(p, 8) = "https://" Then scheme = "https://" Else scheme = "http://"
        rest = Mid(p, Len(scheme) + 1): parts = Split(rest, "/"): masked = scheme
        For j = LBound(parts) To UBound(parts)
            If Len(parts(j)) > 0 Then depth = depth + 1: If depth > 1 Then masked = masked & "/": masked = masked & "***"
        Next j
        MaskPath = "URL:" & masked & "(depth=" & depth & ")"
    ElseIf Left(p, 2) = "\\" Then
        rest = Mid(p, 3): parts = Split(rest, "\")
        For j = LBound(parts) To UBound(parts)
            If Len(parts(j)) > 0 Then depth = depth + 1: If depth > 1 Then masked = masked & "\": masked = masked & "***"
        Next j
        MaskPath = "UNC:\\" & masked & "(depth=" & depth & ")"
    Else
        If Mid(p, 2, 1) = ":" Then drive = Left(p, 2) & "\": rest = Mid(p, 4) Else drive = "": rest = p
        parts = Split(rest, "\")
        For j = LBound(parts) To UBound(parts)
            If Len(parts(j)) > 0 Then depth = depth + 1: If depth > 1 Then masked = masked & "\": masked = masked & "***"
        Next j
        MaskPath = "Local:" & drive & masked & "(depth=" & depth & ")"
    End If
End Function
Public Function RunTest() As String
    RunTest = "OK:" & MaskPath(ThisWorkbook.FullName)
End Function
'@

# Workbooks.Open (adjacent file)
Test-VbaCode 'Basic' 'Info' 'Workbooks.Open' 'Open adjacent file' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim p As String: p = ThisWorkbook.Path & "\probe_open_test.xlsx"
    Dim wb As Workbook
    ' Create a temp xlsx to open
    Dim xlApp As Object: Set xlApp = Application
    Dim tmpWb As Workbook: Set tmpWb = xlApp.Workbooks.Add
    tmpWb.SaveAs p, 51  ' xlOpenXMLWorkbook
    tmpWb.Close False
    If Err.Number <> 0 Then
        RunTest = "FAIL:CreateTemp:" & Err.Description
        Exit Function
    End If
    ' Open it
    Set wb = xlApp.Workbooks.Open(p)
    If Err.Number <> 0 Then
        RunTest = "FAIL:Open:" & Err.Description
    Else
        wb.Close False
        RunTest = "OK:Opened"
    End If
    Kill p
End Function
'@

# SaveAs test (path masked)
Test-VbaCode 'Basic' 'Info' 'SaveAs' 'SaveAs to temp' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Private Function MaskPath(p As String) As String
    Dim parts() As String, rest As String
    Dim depth As Long, masked As String, j As Long, drive As String
    If Len(p) = 0 Then MaskPath = "(empty)": Exit Function
    If Mid(p, 2, 1) = ":" Then drive = Left(p, 2) & "\": rest = Mid(p, 4) Else drive = "": rest = p
    parts = Split(rest, "\")
    For j = LBound(parts) To UBound(parts)
        If Len(parts(j)) > 0 Then depth = depth + 1: If depth > 1 Then masked = masked & "\": masked = masked & "***"
    Next j
    MaskPath = "Local:" & drive & masked & "(depth=" & depth & ")"
End Function
Public Function RunTest() As String
    On Error Resume Next
    Dim p As String: p = Environ$("TEMP") & "\probe_saveas_" & Format(Now, "yyyymmddhhnnss") & ".xlsm"
    ThisWorkbook.SaveAs p, 52
    If Err.Number <> 0 Then
        RunTest = "FAIL:" & Err.Description
    Else
        RunTest = "OK:" & MaskPath(p)
        Kill p
    End If
End Function
'@

# Reference tests (auxiliary, SKIP if VBIDE access denied)
Test-VbaCode 'Basic' 'Info' 'VBProject.References' 'Reference list' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim refs As Object: Set refs = ThisWorkbook.VBProject.References
    If Err.Number <> 0 Then
        RunTest = "FAIL:VBIDE:" & Err.Description
        Exit Function
    End If
    Dim s As String, r As Object, mc As Long
    For Each r In refs
        s = s & r.Name & "(" & r.Major & "." & r.Minor & ")"
        If r.IsBroken Then s = s & "[MISSING]": mc = mc + 1
        s = s & "; "
    Next r
    If mc > 0 Then
        RunTest = "FAIL:Missing=" & mc & ":" & s
    Else
        RunTest = "OK:" & s
    End If
End Function
'@

# ----------------------------------------------------------------------------
# 4. Legacy / Compatibility — deprecated COM objects
# ----------------------------------------------------------------------------

Write-Host ""
Write-Host "--- Legacy / Compatibility Tests ---" -ForegroundColor Cyan

# DAO
Test-VbaCode 'Basic' 'Compat' 'Deprecated: DAO' 'DAO.DBEngine.36' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim o As Object: Set o = CreateObject("DAO.DBEngine.36")
    If Err.Number <> 0 Then
        RunTest = "FAIL:" & Err.Description
    Else
        RunTest = "OK:" & TypeName(o)
    End If
End Function
'@

# Legacy Controls - CommonDialog
Test-VbaCode 'Basic' 'Compat' 'Deprecated: Legacy Controls' 'MSComDlg.CommonDialog' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim o As Object: Set o = CreateObject("MSComDlg.CommonDialog")
    If Err.Number <> 0 Then
        RunTest = "FAIL:" & Err.Description
    Else
        RunTest = "OK"
    End If
End Function
'@

# Legacy Controls - Calendar
Test-VbaCode 'Basic' 'Compat' 'Deprecated: Legacy Controls' 'MSCAL.Calendar' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim o As Object: Set o = CreateObject("MSCAL.Calendar")
    If Err.Number <> 0 Then
        RunTest = "FAIL:" & Err.Description
    Else
        RunTest = "OK"
    End If
End Function
'@

# IE Automation
Test-VbaCode 'Basic' 'Compat' 'Deprecated: IE Automation' 'InternetExplorer.Application' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim o As Object: Set o = CreateObject("InternetExplorer.Application")
    If Err.Number <> 0 Then
        RunTest = "FAIL:" & Err.Description
    Else
        o.Quit
        RunTest = "OK"
    End If
End Function
'@

# ============================================================================
# 5. Storage / Path Tests (path resolution, save behavior, file references)
# ============================================================================

$storageVbaHelpers = @'
Option Explicit
Private Const SCENARIO_LABEL As String = "__SCENARIO__"

Private Function MaskPath(p As String) As String
    Dim parts() As String, rest As String
    Dim depth As Long, masked As String, j As Long, scheme As String, drive As String
    If Len(p) = 0 Then MaskPath = "(empty)": Exit Function
    If Left$(p, 8) = "https://" Or Left$(p, 7) = "http://" Then
        If Left$(p, 8) = "https://" Then scheme = "https://" Else scheme = "http://"
        rest = Mid$(p, Len(scheme) + 1)
        parts = Split(rest, "/")
        For j = LBound(parts) To UBound(parts)
            If Len(parts(j)) > 0 Then
                depth = depth + 1
                If depth > 1 Then masked = masked & "/"
                masked = masked & "***"
            End If
        Next j
        MaskPath = "URL:" & scheme & masked & "(depth=" & depth & ")"
    ElseIf Left$(p, 2) = "\\" Then
        rest = Mid$(p, 3)
        parts = Split(rest, "\")
        For j = LBound(parts) To UBound(parts)
            If Len(parts(j)) > 0 Then
                depth = depth + 1
                If depth > 1 Then masked = masked & "\"
                masked = masked & "***"
            End If
        Next j
        MaskPath = "UNC:\\" & masked & "(depth=" & depth & ")"
    Else
        If Mid$(p, 2, 1) = ":" Then
            drive = Left$(p, 2) & "\"
            rest = Mid$(p, 4)
        Else
            drive = ""
            rest = p
        End If
        parts = Split(rest, "\")
        For j = LBound(parts) To UBound(parts)
            If Len(parts(j)) > 0 Then
                depth = depth + 1
                If depth > 1 Then masked = masked & "\"
                masked = masked & "***"
            End If
        Next j
        MaskPath = "Local:" & drive & masked & "(depth=" & depth & ")"
    End If
End Function

Private Function SafeEnv(name As String) As String
    SafeEnv = Trim$(Environ$(name))
End Function

Private Function GetPreferredSyncRoot(ByRef sourceName As String) As String
    Dim names As Variant, i As Long, value As String
    names = Array("OneDriveCommercial", "OneDrive", "OneDriveConsumer")
    sourceName = ""
    For i = LBound(names) To UBound(names)
        value = SafeEnv(CStr(names(i)))
        If Len(value) > 0 Then
            sourceName = CStr(names(i))
            GetPreferredSyncRoot = value
            Exit Function
        End If
    Next i
End Function

Private Function EnsureSyncRoot(sourceName As String, root As String) As String
    If Len(root) > 0 Then
        EnsureSyncRoot = ""
        Exit Function
    End If

    If SCENARIO_LABEL = "OneDrive_Synced" Then
        EnsureSyncRoot = "FAIL:NoOneDriveEnv"
    Else
        EnsureSyncRoot = "OK:NoOneDriveEnv"
    End If
End Function
'@.Replace('__SCENARIO__', $scenario)

Write-Host ""
Write-Host "--- Storage / Path Tests (Scenario: $scenario) ---" -ForegroundColor Cyan

# CurDir observation (masked)
Test-VbaCode 'Basic' 'Storage' 'CurDir' 'Current directory' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Private Function MaskPath(p As String) As String
    Dim prefix As String, sep As String, parts() As String, rest As String
    Dim depth As Long, masked As String, j As Long, scheme As String, drive As String
    If Len(p) = 0 Then MaskPath = "(empty)": Exit Function
    If Left(p, 8) = "https://" Or Left(p, 7) = "http://" Then
        If Left(p, 8) = "https://" Then scheme = "https://" Else scheme = "http://"
        rest = Mid(p, Len(scheme) + 1): parts = Split(rest, "/"): masked = scheme
        For j = LBound(parts) To UBound(parts)
            If Len(parts(j)) > 0 Then depth = depth + 1: If depth > 1 Then masked = masked & "/": masked = masked & "***"
        Next j
        MaskPath = "URL:" & masked & "(depth=" & depth & ")"
    ElseIf Left(p, 2) = "\\" Then
        rest = Mid(p, 3): parts = Split(rest, "\")
        For j = LBound(parts) To UBound(parts)
            If Len(parts(j)) > 0 Then depth = depth + 1: If depth > 1 Then masked = masked & "\": masked = masked & "***"
        Next j
        MaskPath = "UNC:\\" & masked & "(depth=" & depth & ")"
    Else
        If Mid(p, 2, 1) = ":" Then drive = Left(p, 2) & "\": rest = Mid(p, 4) Else drive = "": rest = p
        parts = Split(rest, "\")
        For j = LBound(parts) To UBound(parts)
            If Len(parts(j)) > 0 Then depth = depth + 1: If depth > 1 Then masked = masked & "\": masked = masked & "***"
        Next j
        MaskPath = "Local:" & drive & masked & "(depth=" & depth & ")"
    End If
End Function
Public Function RunTest() As String
    RunTest = "OK:" & MaskPath(CurDir)
End Function
'@

# Relative path resolution
Test-VbaCode 'Basic' 'Storage' 'Relative Path' 'Dir(ThisWorkbook.Path)' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim p As String: p = ThisWorkbook.Path
    If Len(p) = 0 Then
        RunTest = "FAIL:EmptyPath"
        Exit Function
    End If
    Dim f As String: f = Dir(p & "\*.*")
    If Err.Number <> 0 Then
        RunTest = "FAIL:" & Err.Description
    Else
        RunTest = "OK:FirstFile=" & f
    End If
End Function
'@

# AutoSave state observation
Test-VbaCode 'Basic' 'Storage' 'AutoSave' 'AutoSaveOn state' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim a As Boolean: a = ThisWorkbook.AutoSaveOn
    If Err.Number <> 0 Then
        RunTest = "OK:NotSupported"
    Else
        RunTest = "OK:AutoSave=" & a
    End If
End Function
'@

# Path type detection (local vs URL vs OneDrive)
Test-VbaCode 'Basic' 'Storage' 'Path Type' 'Local vs URL' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    Dim p As String: p = ThisWorkbook.FullName
    If Left(p, 5) = "https" Then
        RunTest = "OK:URL"
    ElseIf InStr(p, "OneDrive") > 0 Then
        RunTest = "OK:OneDrive"
    ElseIf Left(p, 2) = "\\" Then
        RunTest = "OK:UNC"
    Else
        RunTest = "OK:Local"
    End If
End Function
'@

# Minimal Environ$ tests (no helper functions — isolate EDR trigger)
Test-VbaCode 'Basic' 'Storage' 'Environ$ (minimal)' 'OneDriveCommercial' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    Dim v As String: v = Environ$("OneDriveCommercial")
    If Len(v) = 0 Then
        RunTest = "OK:Empty"
    Else
        RunTest = "OK:Len=" & Len(v)
    End If
End Function
'@

Test-VbaCode 'Basic' 'Storage' 'Environ$ (minimal)' 'OneDrive' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    Dim v As String: v = Environ$("OneDrive")
    If Len(v) = 0 Then
        RunTest = "OK:Empty"
    Else
        RunTest = "OK:Len=" & Len(v)
    End If
End Function
'@

Test-VbaCode 'Basic' 'Storage' 'Environ$ (minimal)' 'OneDriveConsumer' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    Dim v As String: v = Environ$("OneDriveConsumer")
    If Len(v) = 0 Then
        RunTest = "OK:Empty"
    Else
        RunTest = "OK:Len=" & Len(v)
    End If
End Function
'@

Test-VbaCode 'Basic' 'Storage' 'OneDriveCommercial Env' 'Environ$("OneDriveCommercial")' -RunMacro 'ProbeTest.RunTest' -VbaCode ($storageVbaHelpers + @'
Public Function RunTest() As String
    Dim root As String
    root = SafeEnv("OneDriveCommercial")
    If Len(root) = 0 Then
        RunTest = "FAIL:Empty"
    Else
        RunTest = "OK:" & MaskPath(root)
    End If
End Function
'@)

Test-VbaCode 'Basic' 'Storage' 'OneDriveCommercial Dir' 'Dir(Environ$("OneDriveCommercial") & "\*.*")' -RunMacro 'ProbeTest.RunTest' -VbaCode ($storageVbaHelpers + @'
Public Function RunTest() As String
    On Error Resume Next
    Dim root As String
    Dim firstEntry As String
    root = SafeEnv("OneDriveCommercial")
    If Len(root) = 0 Then
        RunTest = "FAIL:Empty"
        Exit Function
    End If

    firstEntry = Dir(root & "\*.*")
    If Err.Number <> 0 Then
        RunTest = "FAIL:" & Err.Description
    ElseIf Len(firstEntry) = 0 Then
        RunTest = "OK:(empty)"
    Else
        RunTest = "OK:FirstEntry=" & firstEntry
    End If
End Function
'@)

Test-VbaCode 'Basic' 'Storage' 'OneDrive Environment' 'OneDrive / OneDriveCommercial / OneDriveConsumer' -RunMacro 'ProbeTest.RunTest' -VbaCode ($storageVbaHelpers + @'
Public Function RunTest() As String
    RunTest = "OK:OneDriveCommercial=" & MaskPath(SafeEnv("OneDriveCommercial")) & _
        "; OneDrive=" & MaskPath(SafeEnv("OneDrive")) & _
        "; OneDriveConsumer=" & MaskPath(SafeEnv("OneDriveConsumer"))
End Function
'@)

Test-VbaCode 'Basic' 'Storage' 'Local Sync Root' 'Preferred sync root' -RunMacro 'ProbeTest.RunTest' -VbaCode ($storageVbaHelpers + @'
Public Function RunTest() As String
    Dim sourceName As String, root As String, gate As String
    root = GetPreferredSyncRoot(sourceName)
    gate = EnsureSyncRoot(sourceName, root)
    If Len(gate) > 0 Then
        RunTest = gate
        Exit Function
    End If

    RunTest = "OK:Source=" & sourceName & "; Root=" & MaskPath(root)
End Function
'@)

Test-VbaCode 'Basic' 'Storage' 'Local Sync Root Enumeration' 'Dir(local sync root)' -RunMacro 'ProbeTest.RunTest' -VbaCode ($storageVbaHelpers + @'
Public Function RunTest() As String
    On Error Resume Next
    Dim sourceName As String, root As String, gate As String, firstEntry As String
    root = GetPreferredSyncRoot(sourceName)
    gate = EnsureSyncRoot(sourceName, root)
    If Len(gate) > 0 Then
        RunTest = gate
        Exit Function
    End If

    firstEntry = Dir(root & "\*.*")
    If Err.Number <> 0 Then
        RunTest = "FAIL:" & Err.Description
    ElseIf Len(firstEntry) = 0 Then
        RunTest = "OK:Source=" & sourceName & "; FirstEntry=(empty)"
    Else
        RunTest = "OK:Source=" & sourceName & "; FirstEntry=" & firstEntry
    End If
End Function
'@)

Test-VbaCode 'Basic' 'Storage' 'Local Sync Root Enumeration' 'Dir(local sync root, *.xls*)' -RunMacro 'ProbeTest.RunTest' -VbaCode ($storageVbaHelpers + @'
Public Function RunTest() As String
    On Error Resume Next
    Dim sourceName As String, root As String, gate As String, firstWorkbook As String
    root = GetPreferredSyncRoot(sourceName)
    gate = EnsureSyncRoot(sourceName, root)
    If Len(gate) > 0 Then
        RunTest = gate
        Exit Function
    End If

    firstWorkbook = Dir(root & "\*.xls*")
    If Err.Number <> 0 Then
        RunTest = "FAIL:" & Err.Description
    ElseIf Len(firstWorkbook) = 0 Then
        RunTest = "OK:Source=" & sourceName & "; FirstWorkbook=(none at root)"
    Else
        RunTest = "OK:Source=" & sourceName & "; FirstWorkbook=" & firstWorkbook
    End If
End Function
'@)

Test-VbaCode 'Basic' 'Storage' 'Local Sync Root Enumeration' 'FSO.GetFolder(local sync root)' -RunMacro 'ProbeTest.RunTest' -VbaCode ($storageVbaHelpers + @'
Public Function RunTest() As String
    On Error Resume Next
    Dim sourceName As String, root As String, gate As String
    Dim fso As Object, folder As Object
    root = GetPreferredSyncRoot(sourceName)
    gate = EnsureSyncRoot(sourceName, root)
    If Len(gate) > 0 Then
        RunTest = gate
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(root)
    If Err.Number <> 0 Then
        RunTest = "FAIL:" & Err.Description
    Else
        RunTest = "OK:Source=" & sourceName & "; Files=" & folder.Files.Count & "; SubFolders=" & folder.SubFolders.Count
    End If
End Function
'@)

# ============================================================================
# Extended Tests
# ============================================================================

if ($runExtended) {

    # ----------------------------------------------------------------------------
    # 6. Extended EDR — Shell execution, WMI, SendKeys, AppActivate
    # ----------------------------------------------------------------------------

    Write-Host ""
    Write-Host "--- Extended EDR / Automation Tests ---" -ForegroundColor Cyan

    # Shell via VBA (WScript.Shell.Run)
    Test-VbaCode 'Extended' 'EDR' 'Shell / process' 'WScript.Shell cmd' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        RunTest = "FAIL:CreateObject:" & Err.Description
        Exit Function
    End If
    Dim rc As Long: rc = wsh.Run("cmd /c echo probe", 0, True)
    If Err.Number <> 0 Then
        RunTest = "FAIL:Run:" & Err.Description
    Else
        RunTest = "OK:ExitCode=" & rc
    End If
End Function
'@

    # PowerShell via VBA
    Test-VbaCode 'Extended' 'EDR' 'PowerShell / WScript' 'powershell via VBA' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        RunTest = "FAIL:CreateObject:" & Err.Description
        Exit Function
    End If
    Dim rc As Long: rc = wsh.Run("powershell -Command exit", 0, True)
    If Err.Number <> 0 Then
        RunTest = "FAIL:Run:" & Err.Description
    Else
        RunTest = "OK:ExitCode=" & rc
    End If
End Function
'@

    # WMI via VBA
    Test-VbaCode 'Extended' 'EDR' 'Process / WMI' 'GetObject winmgmts' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim wmi As Object: Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    If Err.Number <> 0 Then
        RunTest = "FAIL:GetObject:" & Err.Description
        Exit Function
    End If
    Dim rs As Object: Set rs = wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE ProcessId = 4")
    If Err.Number <> 0 Then
        RunTest = "FAIL:ExecQuery:" & Err.Description
    Else
        RunTest = "OK:Count=" & rs.Count
    End If
End Function
'@

    # ----------------------------------------------------------------------------
    # 7. Extended Legacy / Compatibility
    # ----------------------------------------------------------------------------

    Write-Host ""
    Write-Host "--- Extended Legacy / Compatibility Tests ---" -ForegroundColor Cyan

    # DDE via VBA
    Test-VbaCode 'Extended' 'Compat' 'Deprecated: DDE' 'DDEInitiate' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim ch As Long
    ch = DDEInitiate("Excel", "Sheet1")
    If Err.Number <> 0 Then
        RunTest = "FAIL:" & Err.Description
    Else
        DDETerminate ch
        RunTest = "OK"
    End If
End Function
'@

    # SendKeys via VBA
    Test-VbaCode 'Extended' 'EDR' 'SendKeys' 'SendKeys (empty)' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    SendKeys ""
    If Err.Number <> 0 Then
        RunTest = "FAIL:" & Err.Description
    Else
        RunTest = "OK"
    End If
End Function
'@

    Test-VbaCode 'Extended' 'EDR' 'AppActivate' 'AppActivate Application.Caption' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    AppActivate Application.Caption
    If Err.Number <> 0 Then
        RunTest = "FAIL:" & Err.Description
    Else
        RunTest = "OK:Activated"
    End If
End Function
'@

    Test-VbaCode 'Extended' 'EDR' 'WScript.Shell AppActivate' 'WScript.Shell AppActivate Excel' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        RunTest = "FAIL:CreateObject:" & Err.Description
        Exit Function
    End If
    Dim ok As Boolean
    ok = wsh.AppActivate(Application.Caption)
    If Err.Number <> 0 Then
        RunTest = "FAIL:AppActivate:" & Err.Description
    Else
        RunTest = "OK:Result=" & ok
    End If
End Function
'@

    Test-VbaCode 'Extended' 'EDR' 'WScript.Shell SendKeys' 'WScript.Shell SendKeys (empty)' -RunMacro 'ProbeTest.RunTest' -VbaCode @'
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        RunTest = "FAIL:CreateObject:" & Err.Description
        Exit Function
    End If
    Call wsh.AppActivate(Application.Caption)
    wsh.SendKeys ""
    If Err.Number <> 0 Then
        RunTest = "FAIL:SendKeys:" & Err.Description
    Else
        RunTest = "OK:SentEmpty"
    End If
End Function
'@

    # ----------------------------------------------------------------------------
    # 8. Cross-process HTTP roundtrip — VBA (XMLHTTP) → PS (HttpListener)
    #    Tests the actual casedesk architecture: VBA calls a local backend.
    # ----------------------------------------------------------------------------

    Write-Host ""
    Write-Host "--- Cross-process HTTP Roundtrip Test ---" -ForegroundColor Cyan

    $httpTestPort = Get-Random -Minimum 20000 -Maximum 45000
    $httpTestPrefix = "http://127.0.0.1:${httpTestPort}/"
    $httpServerJob = $null

    try {
        # Start a minimal HTTP server in a background job (separate process)
        $httpServerJob = Start-Job -ArgumentList $httpTestPrefix -ScriptBlock {
            param($prefix)
            $listener = [System.Net.HttpListener]::new()
            $listener.Prefixes.Add($prefix)
            $listener.Start()

            # Handle one request then exit
            $context = $listener.GetContext()
            $body = (New-Object System.IO.StreamReader($context.Request.InputStream)).ReadToEnd()
            $response = '{"status":"ok","echo":' + '"' + $body.Replace('"','\"') + '"}'
            $buffer = [System.Text.Encoding]::UTF8.GetBytes($response)
            $context.Response.ContentType = 'application/json'
            $context.Response.ContentLength64 = $buffer.Length
            $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
            $context.Response.Close()
            $listener.Stop()
            $listener.Close()
        }

        # Give the server a moment to start
        Start-Sleep -Milliseconds 500

        if ($httpServerJob.State -eq 'Failed') {
            $jobError = Receive-Job $httpServerJob -ErrorAction SilentlyContinue
            Add-ProbeResult 'Extended' 'IPC' 'HTTP server (background)' 'HttpListener (separate process)' 'FAIL' 'Host' 0 ([string]$jobError)
        } else {
            Add-ProbeResult 'Extended' 'IPC' 'HTTP server (background)' 'HttpListener (separate process)' 'OK' 'Host' 0 '' "Listening on $httpTestPrefix"

            # VBA sends HTTP POST to the background server
            Test-VbaCode 'Extended' 'IPC' 'HTTP roundtrip (VBA→PS)' "XMLHTTP POST $httpTestPrefix" -RunMacro 'ProbeTest.RunTest' -VbaCode @"
Option Explicit
Public Function RunTest() As String
    On Error Resume Next
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    If Err.Number <> 0 Then
        RunTest = "FAIL:CreateObject:" & Err.Description
        Exit Function
    End If
    http.Open "POST", "$httpTestPrefix", False
    http.setRequestHeader "Content-Type", "text/plain"
    http.send "ping"
    If Err.Number <> 0 Then
        RunTest = "FAIL:Send:" & Err.Description
        Exit Function
    End If
    If http.Status = 200 Then
        RunTest = "OK:Status=200; Body=" & http.responseText
    Else
        RunTest = "FAIL:Status=" & http.Status
    End If
End Function
"@
        }
    } finally {
        if ($httpServerJob) {
            $httpServerJob | Stop-Job -ErrorAction SilentlyContinue
            $httpServerJob | Remove-Job -Force -ErrorAction SilentlyContinue
        }
    }

    # ----------------------------------------------------------------------------
    # 9. Extended Host-side — Storage paths, Window automation, IPC
    # ----------------------------------------------------------------------------

    Write-Host ""
    Write-Host "--- Extended Host-side Tests ---" -ForegroundColor Cyan
    Test-HostStoragePaths
    Test-HostWindowAutomation
}

# ============================================================================
# Output
# ============================================================================

Write-Host ""
Write-Host "--- Writing Results ---" -ForegroundColor Cyan

$scriptRoot = Split-Path $PSScriptRoot -Parent
$outDir = if ($OutputDirectory) { $OutputDirectory } else { Join-Path $scriptRoot 'output' }
if (-not (Test-Path $outDir)) { New-Item $outDir -ItemType Directory -Force | Out-Null }
$outPath = if ($OutputDirectory) { Join-Path $outDir $ProbeReportName } else { Join-Path $outDir "probe_result_$(Get-Date -Format yyyyMMdd_HHmmss).txt" }

$sb = [System.Text.StringBuilder]::new()
[void]$sb.AppendLine("# Environment Probe Results")
[void]$sb.AppendLine("# Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$sb.AppendLine("# Computer: $maskedPC")
[void]$sb.AppendLine("# User: $maskedUser")
[void]$sb.AppendLine("# Mode: $(if($runExtended){'Basic + Extended'}else{'Basic only'})")
[void]$sb.AppendLine("# Scenario: $scenario")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("Level`tCategory`tPattern`tTarget`tResult`tPhase`tErrMsg`tDetail")

foreach ($r in $results) {
    [void]$sb.AppendLine("$($r.Level)`t$($r.Category)`t$($r.Pattern)`t$($r.Target)`t$($r.Result)`t$($r.Phase)`t$($r.ErrMsg)`t$($r.Detail)")
}

$utf8Bom = New-Object System.Text.UTF8Encoding $true
if ($ReturnTextFile) {
    [IO.File]::WriteAllText($ReturnTextFile, $sb.ToString(), $utf8Bom)
} else {
    [IO.File]::WriteAllText($outPath, $sb.ToString(), $utf8Bom)
}

# Restore console Ctrl+C handling
[Console]::TreatControlCAsInput = $false

# Cleanup
Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue

# Summary
$sw.Stop()
$okCount = @($results | Where-Object { $_.Result -eq 'OK' }).Count
$failCount = @($results | Where-Object { $_.Result -eq 'FAIL' }).Count
$blockedCount = @($results | Where-Object { $_.Result -eq 'BLOCKED' }).Count
$timeoutCount = @($results | Where-Object { $_.Result -eq 'TIMEOUT' }).Count
$skipCount = @($results | Where-Object { $_.Result -eq 'SKIP' }).Count

Write-Host ""
Write-Host "=== Results ===" -ForegroundColor Cyan
Write-Host "  OK:      $okCount" -ForegroundColor Green
Write-Host "  FAIL:    $failCount" -ForegroundColor $(if($failCount -gt 0){'Red'}else{'Green'})
Write-Host "  BLOCKED: $blockedCount" -ForegroundColor $(if($blockedCount -gt 0){'Yellow'}else{'Green'})
Write-Host "  TIMEOUT: $timeoutCount" -ForegroundColor $(if($timeoutCount -gt 0){'Red'}else{'Green'})
Write-Host "  SKIP:    $skipCount" -ForegroundColor DarkGray
Write-Host "  Time: $([Math]::Round($sw.Elapsed.TotalSeconds, 1))s" -ForegroundColor Gray
if (-not $ReturnTextFile) {
    Write-Host "  Output: $outPath" -ForegroundColor Gray
}

# Force exit — background listeners/jobs may prevent graceful shutdown
[Environment]::Exit(0)
