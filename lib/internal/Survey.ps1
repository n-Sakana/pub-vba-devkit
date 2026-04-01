param(
    [string]$OutputDirectory,
    [string]$ReturnTextFile
)

$ErrorActionPreference = 'Stop'
$script:SurveyDiagnostics = [System.Collections.Generic.List[object]]::new()
$script:SurveyStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

function Write-SurveySection {
    param([string]$Title)

    Write-Host ''
    Write-Host "--- $Title ---" -ForegroundColor Cyan
}

function Write-SurveyItem {
    param(
        [string]$Name,
        [string]$Status = 'INFO',
        [string]$Detail = ''
    )

    $color = switch ($Status) {
        'OK' { 'Green' }
        'FAIL' { 'Red' }
        'WARN' { 'Yellow' }
        default { 'Gray' }
    }

    $suffix = if ([string]::IsNullOrWhiteSpace($Detail)) { '' } else { " - $Detail" }
    Write-Host "  [$Status] $Name$suffix" -ForegroundColor $color
}

function Add-SurveyDiagnostic {
    param(
        [string]$Category,
        [string]$Target,
        [string]$Message,
        [bool]$TimedOut = $false
    )

    $maskedMessage = Mask-EmbeddedPaths -Text $Message
    $script:SurveyDiagnostics.Add([ordered]@{
        Category = $Category
        Target   = $Target
        TimedOut = $TimedOut
        Message  = $maskedMessage
    }) | Out-Null

    Write-SurveyItem -Name "$Category / $Target" -Status 'WARN' -Detail $maskedMessage
}

function Get-DefaultRegistryValue {
    param([string]$Path)

    try {
        $item = Get-Item -Path $Path -ErrorAction Stop
        return $item.GetValue('')
    } catch {
        return $null
    }
}

function Get-RegistryValue {
    param(
        [string]$Path,
        [string]$Name
    )

    try {
        return (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name
    } catch {
        return $null
    }
}

function Get-FileVersionSummary {
    param([string]$Path)

    if (-not $Path) { return $null }
    if (-not (Test-Path -LiteralPath $Path)) { return $null }

    try {
        $item = Get-Item -LiteralPath $Path -ErrorAction Stop
        return [ordered]@{
            ProductVersion = $item.VersionInfo.ProductVersion
            FileVersion    = $item.VersionInfo.FileVersion
        }
    } catch {
        return $null
    }
}

function Resolve-VersionText {
    param(
        [string]$Text,
        [string[]]$Patterns
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $value = $Text.Trim()
    $candidatePatterns = @()
    if ($Patterns) {
        $candidatePatterns += $Patterns
    }
    $candidatePatterns += @(
        'version\s+"([^"]+)"',
        'Python\s+([0-9][0-9A-Za-z\.\-_]*)',
        '^v([0-9][0-9A-Za-z\.\-_]*)',
        '([0-9]+(?:\.[0-9A-Za-z\-_]+)+)'
    )

    foreach ($pattern in $candidatePatterns) {
        if ($value -match $pattern) {
            return $matches[1]
        }
    }

    return $value
}

function Get-SafeCimInstance {
    param(
        [string]$ClassName,
        [string]$Filter,
        [string]$SourceTag,
        [int]$OperationTimeoutSec = 5
    )

    try {
        if ($Filter) {
            return Get-CimInstance -ClassName $ClassName -Filter $Filter -OperationTimeoutSec $OperationTimeoutSec -ErrorAction Stop
        }

        return Get-CimInstance -ClassName $ClassName -OperationTimeoutSec $OperationTimeoutSec -ErrorAction Stop
    } catch {
        $message = $_.Exception.Message
        Add-SurveyDiagnostic -Category 'CIM' -Target $(if ($SourceTag) { $SourceTag } else { $ClassName }) -Message $message -TimedOut ([bool]($message -match '(?i)timed out|timeout'))
        return $null
    }
}

function Mask-Path {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    $value = $Path.Trim()

    if ($value -match '^https?://') {
        $parts = $value -split '/'
        $depth = [Math]::Max($parts.Count - 3, 0)
        $masked = @($parts[0], '', '') + (@('***') * $depth)
        return "URL:$($masked -join '/') (depth=$depth)"
    }

    if ($value -match '^\\\\') {
        $parts = $value.TrimStart('\') -split '\\'
        $depth = $parts.Count
        return "UNC:\\$((@('***') * $depth) -join '\') (depth=$depth)"
    }

    if ($value -match '^[A-Za-z]:\\') {
        $parts = $value -split '\\'
        $depth = [Math]::Max($parts.Count - 1, 0)
        $masked = @($parts[0])
        if ($parts.Count -gt 1) {
            $masked += @('***') * ($parts.Count - 1)
        }
        return "Local:$($masked -join '\') (depth=$depth)"
    }

    return '(masked)'
}

function Mask-EmbeddedPaths {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $Text
    }

    $result = [regex]::Replace($Text, '[A-Za-z]:\\[^\]\r\n]+', {
        param($match)
        Mask-Path -Path $match.Value
    })

    $result = [regex]::Replace($result, '\\\\[^\]\r\n]+', {
        param($match)
        Mask-Path -Path $match.Value
    })

    return $result
}

function Get-RegistryBiosInfo {
    try {
        return Get-ItemProperty -Path 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -ErrorAction Stop
    } catch {
        return $null
    }
}

function Get-TotalMemoryGBFallback {
    try {
        Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
        $info = New-Object Microsoft.VisualBasic.Devices.ComputerInfo
        return [Math]::Round($info.TotalPhysicalMemory / 1GB, 1)
    } catch {
        return $null
    }
}

function Convert-ToSurveyDateText {
    param(
        $Value,
        [string]$Format,
        [string]$SourceTag
    )

    if ($null -eq $Value) {
        return $null
    }

    try {
        if ($Value -is [datetime]) {
            return $Value.ToString($Format)
        }
        if ($Value -is [datetimeoffset]) {
            return $Value.ToString($Format)
        }

        $text = [string]$Value
        if ([string]::IsNullOrWhiteSpace($text)) {
            return $null
        }

        if ($text -match '^\d{14}\.\d{6}(?:[+-]\d{3}|\*{3})$') {
            return ([Management.ManagementDateTimeConverter]::ToDateTime($text)).ToString($Format)
        }

        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse($text, [ref]$parsed)) {
            return $parsed.ToString($Format)
        }
    } catch {
        Add-SurveyDiagnostic -Category 'DateParse' -Target $SourceTag -Message $_.Exception.Message
        return $null
    }

    Add-SurveyDiagnostic -Category 'DateParse' -Target $SourceTag -Message "Unsupported date value: $([string]$Value)"
    return $null
}

function Convert-RegistryDisplayText {
    param($Value)

    if ($null -eq $Value) {
        return $null
    }

    if (($Value -is [System.Array]) -and -not ($Value -is [string])) {
        $elements = @($Value)
        if ($elements.Count -gt 0 -and @($elements | Where-Object { $_ -isnot [byte] -and $_ -isnot [int] -and $_ -isnot [uint16] -and $_ -isnot [int16] }).Count -eq 0) {
            $bytes = New-Object byte[] ($elements.Count)
            for ($i = 0; $i -lt $elements.Count; $i++) {
                $bytes[$i] = [byte]$elements[$i]
            }

            $text = [System.Text.Encoding]::Unicode.GetString($bytes).Trim([char]0).Trim()
            if ([string]::IsNullOrWhiteSpace($text)) {
                $text = [System.Text.Encoding]::ASCII.GetString($bytes).Trim([char]0).Trim()
            }
            return $text
        }
    }

    if ($Value -is [byte[]]) {
        $text = [System.Text.Encoding]::Unicode.GetString($Value).Trim([char]0).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) {
            $text = [System.Text.Encoding]::ASCII.GetString($Value).Trim([char]0).Trim()
        }
        return $text
    }

    return ([string]$Value).Trim()
}

function Mask-ProxyValue {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    if ($Value -match '(?i)direct access|直接アクセス|プロキシ サーバーなし|no proxy') {
        return $Value.Trim()
    }

    $segments = $Value -split ';'
    $maskedSegments = foreach ($segment in $segments) {
        $part = $segment.Trim()
        if ([string]::IsNullOrWhiteSpace($part)) {
            continue
        }

        if ($part -match '^(?<prefix>[^=]+)=(?<endpoint>.+)$') {
            $prefix = $matches['prefix']
            $endpoint = $matches['endpoint'].Trim()
            if ($endpoint -match '^https?://') {
                "$prefix=$(Mask-Path -Path $endpoint)"
            } elseif ($endpoint -match '^(?<host>[^:]+):(?<port>\d+)$') {
                "$prefix=***:$($matches['port'])"
            } elseif ($endpoint -eq '<local>') {
                "$prefix=<local>"
            } else {
                "$prefix=(masked)"
            }
            continue
        }

        if ($part -match '^https?://') {
            Mask-Path -Path $part
        } elseif ($part -match '^(?<host>[^:]+):(?<port>\d+)$') {
            "***:$($matches['port'])"
        } elseif ($part -eq '<local>') {
            '<local>'
        } else {
            '(masked)'
        }
    }

    return ($maskedSegments -join ';')
}

function Get-DriveInventory {
    $cimDrives = @(Get-SafeCimInstance -ClassName 'Win32_LogicalDisk' -Filter "DriveType = 3" -SourceTag 'CIM:Win32_LogicalDisk' | Where-Object { $_ })
    if ($cimDrives.Count -gt 0) {
        return [ordered]@{
            Source = 'CIM:Win32_LogicalDisk'
            Items  = @($cimDrives | ForEach-Object {
                [ordered]@{
                    DeviceId   = $_.DeviceID
                    FileSystem = $_.FileSystem
                    SizeGB     = [Math]::Round($_.Size / 1GB, 1)
                    FreeGB     = [Math]::Round($_.FreeSpace / 1GB, 1)
                }
            })
        }
    }

    return [ordered]@{
        Source = 'System.IO.DriveInfo'
        Items  = @([System.IO.DriveInfo]::GetDrives() | Where-Object { $_.DriveType -eq 'Fixed' -and $_.IsReady } | ForEach-Object {
            [ordered]@{
                DeviceId   = $_.Name
                FileSystem = $_.DriveFormat
                SizeGB     = [Math]::Round($_.TotalSize / 1GB, 1)
                FreeGB     = [Math]::Round($_.AvailableFreeSpace / 1GB, 1)
            }
        })
    }
}

function Get-GpuInventory {
    $cimRows = @(Get-SafeCimInstance -ClassName 'Win32_VideoController' -SourceTag 'CIM:Win32_VideoController')
    $cimNames = @($cimRows | ForEach-Object { $_.Name } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    if ($cimNames.Count -gt 0) {
        return [ordered]@{
            Source = 'CIM:Win32_VideoController'
            Items  = $cimNames
        }
    }

    if (Get-Command Get-PnpDevice -ErrorAction SilentlyContinue) {
        try {
            $pnpNames = @(Get-PnpDevice -Class Display -PresentOnly -ErrorAction Stop |
                ForEach-Object { $_.FriendlyName } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique)
            if ($pnpNames.Count -gt 0) {
                return [ordered]@{
                    Source = 'PnpDevice:Display'
                    Items  = $pnpNames
                }
            }
        } catch {
            Add-SurveyDiagnostic -Category 'PnpDevice' -Target 'Display' -Message $_.Exception.Message
        }
    }

    $classKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}'
    $registryNames = [System.Collections.Generic.List[string]]::new()
    if (Test-Path -LiteralPath $classKey) {
        foreach ($subKey in @(Get-ChildItem -Path $classKey -ErrorAction SilentlyContinue)) {
            try {
                $props = Get-ItemProperty -Path $subKey.PSPath -ErrorAction Stop
                foreach ($name in @(
                    (Convert-RegistryDisplayText -Value $props.DriverDesc),
                    (Convert-RegistryDisplayText -Value $props.'HardwareInformation.AdapterString')
                )) {
                    if (-not [string]::IsNullOrWhiteSpace($name)) {
                        $registryNames.Add([string]$name) | Out-Null
                    }
                }
            } catch {
                Add-SurveyDiagnostic -Category 'Registry' -Target $subKey.PSChildName -Message $_.Exception.Message
            }
        }
    }

    $fallbackNames = @($registryNames | Sort-Object -Unique)
    if ($fallbackNames.Count -gt 0) {
        return [ordered]@{
            Source = "Registry:$classKey"
            Items  = $fallbackNames
        }
    }

    return [ordered]@{
        Source = $null
        Items  = @()
    }
}

function Get-CurrentUserProxySettings {
    $path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'

    try {
        $item = Get-ItemProperty -Path $path -ErrorAction Stop
        return [ordered]@{
            Source        = "Registry:$path"
            ProxyEnabled  = [bool]$item.ProxyEnable
            ProxyServer   = Mask-ProxyValue -Value $item.ProxyServer
            AutoConfigUrl = Mask-Path -Path $item.AutoConfigURL
            AutoDetect    = [bool]$item.AutoDetect
        }
    } catch {
        return [ordered]@{
            Source        = "Registry:$path"
            ProxyEnabled  = $false
            ProxyServer   = $null
            AutoConfigUrl = $null
            AutoDetect    = $false
        }
    }
}

function Get-WinHttpProxySettings {
    $lines = @()
    try {
        $prev = [Console]::OutputEncoding
        [Console]::OutputEncoding = [Text.Encoding]::UTF8
        $lines = @(& netsh winhttp show proxy 2>&1 | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        [Console]::OutputEncoding = $prev
    } catch {
        $lines = @()
    }

    return [ordered]@{
        Source = 'netsh winhttp show proxy'
        Lines  = @($lines | ForEach-Object {
            $line = [string]$_
            $separatorIndex = $line.IndexOf(':')
            if ($separatorIndex -lt 0) {
                $separatorIndex = $line.IndexOf([char]0xFF1A)
            }

            if ($separatorIndex -gt -1 -and $separatorIndex -lt ($line.Length - 1)) {
                $label = $line.Substring(0, $separatorIndex + 1)
                $value = $line.Substring($separatorIndex + 1).Trim()
                "$label $(Mask-ProxyValue -Value $value)"
            } else {
                Mask-EmbeddedPaths -Text $line
            }
        })
    }
}

function Get-OneDriveInventory {
    $exePath = Resolve-AppPath -ExeName 'OneDrive.exe'
    $accountRoot = 'HKCU:\Software\Microsoft\OneDrive\Accounts'
    $accounts = @()

    if (Test-Path -LiteralPath $accountRoot) {
        foreach ($key in @(Get-ChildItem -Path $accountRoot -ErrorAction SilentlyContinue)) {
            $props = $null
            try {
                $props = Get-ItemProperty -Path $key.PSPath -ErrorAction Stop
            } catch {
                $props = $null
            }

            $accounts += [ordered]@{
                AccountKey = $key.PSChildName
                Business   = [bool]($key.PSChildName -like 'Business*')
                UserFolder = Mask-Path -Path $(if ($props) { $props.UserFolder } else { $null })
            }
        }
    }

    $envRows = foreach ($name in @('OneDriveCommercial', 'OneDrive', 'OneDriveConsumer')) {
        $value = [Environment]::GetEnvironmentVariable($name)
        [ordered]@{
            Name   = $name
            Path   = Mask-Path -Path $value
            Exists = if ([string]::IsNullOrWhiteSpace($value)) { $false } else { [bool](Test-Path -LiteralPath $value -ErrorAction SilentlyContinue) }
        }
    }

    return [ordered]@{
        Source               = "AppPaths + Registry:$accountRoot + Environment"
        Installed            = [bool]$exePath
        ExecutablePath       = Mask-Path -Path $exePath
        AccountCount         = $accounts.Count
        BusinessAccountCount = @($accounts | Where-Object { $_.Business }).Count
        SharePointSyncLikely = [bool](@($accounts | Where-Object { $_.Business }).Count -gt 0)
        SyncRoots            = @($accounts | ForEach-Object { $_.UserFolder } | Where-Object { $_ })
        EnvironmentVariables = @($envRows)
        EnvironmentPaths     = @($envRows | Where-Object { $_.Path -and $_.Path -ne '(empty)' } | ForEach-Object { $_.Path })
    }
}

function Resolve-AppPath {
    param([string]$ExeName)

    $candidates = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$ExeName",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\$ExeName"
    )

    foreach ($path in $candidates) {
        $value = Get-DefaultRegistryValue -Path $path
        if ($value) { return $value }
    }

    $cmd = Get-Command $ExeName -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($cmd) {
        return $cmd.Source
    }

    return $null
}

function Get-UninstallEntries {
    $paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $rows = foreach ($path in $paths) {
        Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName } |
            Select-Object DisplayName, DisplayVersion, Publisher, InstallLocation, InstallDate, UninstallString
    }

    return @($rows)
}

function Find-UninstallEntry {
    param([string[]]$Patterns)

    $entries = Get-UninstallEntries
    foreach ($pattern in $Patterns) {
        $match = $entries | Where-Object { $_.DisplayName -like $pattern } | Select-Object -First 1
        if ($match) { return $match }
    }

    return $null
}

function Get-CommandInventory {
    param(
        [string]$Name,
        [string[]]$VersionArgs = @('--version'),
        [string[]]$VersionPatterns = @()
    )

    $cmd = Get-Command $Name -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $cmd) {
        return [ordered]@{
            Name      = $Name
            Available = $false
            Path      = $null
            Version   = $null
        }
    }

    $version = $null
    $versionRaw = $null
    try {
        $versionOutput = @(& $cmd.Source @VersionArgs 2>&1)
        if ($versionOutput.Count -gt 0) {
            $versionRaw = (($versionOutput | ForEach-Object { [string]$_ }) -join "`n").Trim()
            $version = Resolve-VersionText -Text $versionRaw -Patterns $VersionPatterns
            if ($versionRaw) { $versionRaw = Mask-EmbeddedPaths -Text $versionRaw }
        }
    } catch {
        $version = $null
        $versionRaw = $null
    }

    return [ordered]@{
        Name      = $Name
        Available = $true
        Path      = Mask-Path -Path $cmd.Source
        Version   = $version
        VersionRaw = $versionRaw
    }
}

function Convert-DotNetRelease {
    param([int]$Release)

    if ($Release -ge 533320) { return '4.8.1' }
    if ($Release -ge 528040) { return '4.8' }
    if ($Release -ge 461808) { return '4.7.2' }
    if ($Release -ge 461308) { return '4.7.1' }
    if ($Release -ge 460798) { return '4.7' }
    if ($Release -ge 394802) { return '4.6.2' }
    if ($Release -ge 394254) { return '4.6.1' }
    if ($Release -ge 393295) { return '4.6' }
    if ($Release -ge 379893) { return '4.5.2' }
    if ($Release -ge 378675) { return '4.5.1' }
    if ($Release -ge 378389) { return '4.5' }
    return "Unknown ($Release)"
}

function Get-DotNetFrameworkInventory {
    $release = Get-RegistryValue -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name 'Release'
    if (-not $release) {
        return [ordered]@{
            Installed = $false
            Version   = $null
            Release   = $null
        }
    }

    return [ordered]@{
        Installed = $true
        Version   = Convert-DotNetRelease -Release ([int]$release)
        Release   = [int]$release
    }
}

function Get-DotNetRuntimeInventory {
    $dotnet = Get-Command dotnet -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $dotnet) {
        return [ordered]@{
            Available = $false
            Path      = $null
            Version   = $null
            Runtimes  = @()
            Sdks      = @()
        }
    }

    $version = $null
    $runtimes = @()
    $sdks = @()

    try { $version = (& $dotnet.Source --version 2>$null | Select-Object -First 1) } catch {}
    try { $runtimes = @(& $dotnet.Source --list-runtimes 2>$null) } catch {}
    try { $sdks = @(& $dotnet.Source --list-sdks 2>$null) } catch {}

    return [ordered]@{
        Available = $true
        Path      = Mask-Path -Path $dotnet.Source
        Version   = $version
        Runtimes  = @($runtimes | ForEach-Object { Mask-EmbeddedPaths -Text $_ })
        Sdks      = @($sdks | ForEach-Object { Mask-EmbeddedPaths -Text $_ })
    }
}

function Test-AddTypeCapability {
    try {
        Add-Type -TypeDefinition @'
public static class SurveyAddTypeCheck
{
    public static string Ping()
    {
        return "OK";
    }
}
'@ -Language CSharp -ErrorAction Stop

        return [ordered]@{
            Available = $true
            Detail    = [SurveyAddTypeCheck]::Ping()
        }
    } catch {
        return [ordered]@{
            Available = $false
            Detail    = $_.Exception.Message
        }
    }
}

function Test-AssemblyCapability {
    param(
        [string]$Name,
        [string]$AssemblyName,
        [string]$TypeName
    )

    try {
        if ($AssemblyName) {
            Add-Type -AssemblyName $AssemblyName -ErrorAction Stop
        }

        $resolved = if ($TypeName) { $TypeName -as [type] } else { $null }

        return [ordered]@{
            Name      = $Name
            Available = [bool]$resolved
            Detail    = if ($resolved) { $resolved.FullName } else { $null }
        }
    } catch {
        return [ordered]@{
            Name      = $Name
            Available = $false
            Detail    = $_.Exception.Message
        }
    }
}

function Test-ObjectCapability {
    param(
        [string]$Name,
        [scriptblock]$Action
    )

    try {
        $detail = & $Action
        return [ordered]@{
            Name      = $Name
            Available = $true
            Detail    = [string]$detail
        }
    } catch {
        return [ordered]@{
            Name      = $Name
            Available = $false
            Detail    = $_.Exception.Message
        }
    }
}

function Get-RegisteredComServer {
    param([string]$ProgId)

    $base = "Registry::HKEY_CLASSES_ROOT\$ProgId"
    if (-not (Test-Path -LiteralPath $base)) {
        return [ordered]@{
            ProgId      = $ProgId
            Registered  = $false
            Clsid       = $null
            Server      = $null
            ServerPath  = $null
            FileVersion = $null
        }
    }

    $clsid = Get-DefaultRegistryValue -Path "$base\CLSID"
    $localServer = $null
    $inprocServer = $null

    if ($clsid) {
        $localServer = Get-DefaultRegistryValue -Path "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\LocalServer32"
        $inprocServer = Get-DefaultRegistryValue -Path "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\InprocServer32"
    }

    $server = if ($localServer) { $localServer } else { $inprocServer }
    $serverPath = $null
    if ($server) {
        if ($server -match '^[\s"]*([^"]+?\.exe)') {
            $serverPath = $matches[1]
        } elseif ($server -match '^[\s"]*([^"]+?\.dll)') {
            $serverPath = $matches[1]
        }
    }

    $version = Get-FileVersionSummary -Path $serverPath

    return [ordered]@{
        ProgId      = $ProgId
        Registered  = $true
        Clsid       = $clsid
        Server      = Mask-Path -Path $server
        ServerPath  = Mask-Path -Path $serverPath
        FileVersion = if ($version) { $version.ProductVersion } else { $null }
    }
}

function Format-SizeGB {
    param([double]$Bytes)
    return ('{0:N1} GB' -f ($Bytes / 1GB))
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$scriptRoot = Split-Path -Parent $PSScriptRoot
$outDir = if ($OutputDirectory) { $OutputDirectory } else { Join-Path $scriptRoot "output\${timestamp}_survey" }
New-Item -Path $outDir -ItemType Directory -Force | Out-Null

Write-Host '=== Environment Survey ===' -ForegroundColor Cyan
Write-Host "Output folder: $outDir" -ForegroundColor Gray

$computer = Get-SafeCimInstance -ClassName 'Win32_ComputerSystem' -SourceTag 'CIM:Win32_ComputerSystem'
$os = Get-SafeCimInstance -ClassName 'Win32_OperatingSystem' -SourceTag 'CIM:Win32_OperatingSystem'
$bios = Get-SafeCimInstance -ClassName 'Win32_BIOS' -SourceTag 'CIM:Win32_BIOS'
$cpu = @(Get-SafeCimInstance -ClassName 'Win32_Processor' -SourceTag 'CIM:Win32_Processor' | Select-Object -First 1)[0]
$gpuInventory = Get-GpuInventory
$gpus = @($gpuInventory.Items)
$driveInventory = Get-DriveInventory
$drives = $driveInventory.Items
$biosReg = Get-RegistryBiosInfo
$ntCurrent = try { Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop } catch { $null }
$currentUserProxy = Get-CurrentUserProxySettings
$winHttpProxy = Get-WinHttpProxySettings
$oneDrive = Get-OneDriveInventory

$machine = [ordered]@{
    ComputerName       = '(masked)'
    Manufacturer       = if ($computer) { $computer.Manufacturer } else { $biosReg.SystemManufacturer }
    Model              = if ($computer) { $computer.Model } else { $biosReg.SystemProductName }
    SystemType         = if ($computer) { $computer.SystemType } else { if ([Environment]::Is64BitOperatingSystem) { 'x64-based PC' } else { 'x86-based PC' } }
    BiosVersion        = if ($bios) { ($bios.SMBIOSBIOSVersion -join ', ') } else { ($biosReg.BIOSVersion -join ', ') }
    TotalMemoryGB      = if ($computer) { [Math]::Round($computer.TotalPhysicalMemory / 1GB, 1) } else { Get-TotalMemoryGBFallback }
    CpuName            = if ($cpu) { $cpu.Name } else { Get-RegistryValue -Path 'HKLM:\HARDWARE\DESCRIPTION\System\CentralProcessor\0' -Name 'ProcessorNameString' }
    CpuCores           = if ($cpu) { $cpu.NumberOfCores } else { $null }
    CpuLogicalProcessors = if ($cpu) { $cpu.NumberOfLogicalProcessors } else { $null }
    GPUs               = @($gpus)
    Drives             = $drives
}

Write-SurveySection 'Machine'
Write-SurveyItem -Name 'Manufacturer' -Status $(if ($machine.Manufacturer) { 'OK' } else { 'WARN' }) -Detail ([string]$machine.Manufacturer)
Write-SurveyItem -Name 'Model' -Status $(if ($machine.Model) { 'OK' } else { 'WARN' }) -Detail ([string]$machine.Model)
Write-SurveyItem -Name 'System Type' -Status 'OK' -Detail ([string]$machine.SystemType)
Write-SurveyItem -Name 'BIOS' -Status $(if ($machine.BiosVersion) { 'OK' } else { 'WARN' }) -Detail ([string]$machine.BiosVersion)
Write-SurveyItem -Name 'Memory' -Status $(if ($null -ne $machine.TotalMemoryGB) { 'OK' } else { 'WARN' }) -Detail "$(if ($null -ne $machine.TotalMemoryGB) { $machine.TotalMemoryGB } else { '?' }) GB"
Write-SurveyItem -Name 'CPU' -Status $(if ($machine.CpuName) { 'OK' } else { 'WARN' }) -Detail "$(if ($machine.CpuName) { $machine.CpuName } else { '' })$(if ($machine.CpuCores -or $machine.CpuLogicalProcessors) { " ($($machine.CpuCores)/$($machine.CpuLogicalProcessors))" } else { '' })"
Write-SurveyItem -Name 'GPU' -Status $(if (@($machine.GPUs).Count -gt 0) { 'OK' } else { 'WARN' }) -Detail "$(if (@($machine.GPUs).Count -gt 0) { $machine.GPUs -join '; ' } else { '(none reported)' })"
Write-SurveyItem -Name 'Drives' -Status $(if (@($machine.Drives).Count -gt 0) { 'OK' } else { 'WARN' }) -Detail "$(if (@($machine.Drives).Count -gt 0) { (@($machine.Drives | ForEach-Object { $_.DeviceId }) -join '; ') } else { '(none reported)' })"

$osInfo = [ordered]@{
    Caption       = if ($os) { $os.Caption } else { $ntCurrent.ProductName }
    Version       = if ($os) { $os.Version } else { $ntCurrent.DisplayVersion }
    BuildNumber   = if ($os) { $os.BuildNumber } else { $ntCurrent.CurrentBuild }
    Architecture  = if ($os) { $os.OSArchitecture } else { if ([Environment]::Is64BitOperatingSystem) { '64-bit' } else { '32-bit' } }
    InstallDate   = if ($os) { Convert-ToSurveyDateText -Value $os.InstallDate -Format 'yyyy-MM-dd' -SourceTag 'CIM:Win32_OperatingSystem.InstallDate' } elseif ($ntCurrent.InstallDate) { ([DateTimeOffset]::FromUnixTimeSeconds([int64]$ntCurrent.InstallDate)).ToString('yyyy-MM-dd') } else { $null }
    LastBoot      = if ($os) { Convert-ToSurveyDateText -Value $os.LastBootUpTime -Format 'yyyy-MM-dd HH:mm:ss' -SourceTag 'CIM:Win32_OperatingSystem.LastBootUpTime' } else { $null }
}

Write-SurveySection 'OS'
Write-SurveyItem -Name 'Caption' -Status $(if ($osInfo.Caption) { 'OK' } else { 'WARN' }) -Detail ([string]$osInfo.Caption)
Write-SurveyItem -Name 'Version' -Status $(if ($osInfo.Version) { 'OK' } else { 'WARN' }) -Detail "$($osInfo.Version) build=$($osInfo.BuildNumber)"
Write-SurveyItem -Name 'Architecture' -Status 'OK' -Detail ([string]$osInfo.Architecture)
Write-SurveyItem -Name 'Installed' -Status $(if ($osInfo.InstallDate) { 'OK' } else { 'WARN' }) -Detail ([string]$osInfo.InstallDate)
Write-SurveyItem -Name 'Last Boot' -Status $(if ($osInfo.LastBoot) { 'OK' } else { 'WARN' }) -Detail ([string]$osInfo.LastBoot)

$network = [ordered]@{
    DomainJoined     = if ($computer) { [bool]$computer.PartOfDomain } else { $null }
    DomainName       = if ($computer -and $computer.PartOfDomain) { '(masked)' } else { $null }
    CurrentUserProxy = $currentUserProxy
    WinHttpProxy     = $winHttpProxy
    OneDrive         = $oneDrive
}

Write-SurveySection 'Network'
Write-SurveyItem -Name 'Domain Joined' -Status 'INFO' -Detail "$(if ($network.DomainJoined -eq $true) { 'yes' } elseif ($network.DomainJoined -eq $false) { 'no' } else { 'unknown' })"
Write-SurveyItem -Name 'Current User Proxy' -Status 'INFO' -Detail "enabled=$($network.CurrentUserProxy.ProxyEnabled) autodetect=$($network.CurrentUserProxy.AutoDetect) server=$($network.CurrentUserProxy.ProxyServer)"
Write-SurveyItem -Name 'WinHTTP Proxy' -Status 'INFO' -Detail "$(if (@($network.WinHttpProxy.Lines).Count -gt 0) { $network.WinHttpProxy.Lines[0] } else { 'not reported' })"
Write-SurveyItem -Name 'OneDrive' -Status 'INFO' -Detail "installed=$($network.OneDrive.Installed) accounts=$($network.OneDrive.AccountCount) business=$($network.OneDrive.BusinessAccountCount)"
Write-SurveyItem -Name 'OneDrive Env Vars' -Status 'INFO' -Detail "$((@($network.OneDrive.EnvironmentVariables | Where-Object { $_.Path -and $_.Path -ne '(empty)' }).Count)) configured"
Write-SurveyItem -Name 'SharePoint Sync Likely' -Status 'INFO' -Detail ([string]$network.OneDrive.SharePointSyncLikely)

$runtimes = [ordered]@{
    PowerShell51 = [ordered]@{
        Available = $true
        Path      = Mask-Path -Path ((Get-Command powershell.exe | Select-Object -First 1).Source)
        Version   = $PSVersionTable.PSVersion.ToString()
        Edition   = $PSVersionTable.PSEdition
    }
    PowerShell7       = Get-CommandInventory -Name 'pwsh' -VersionArgs @('-NoLogo', '-NoProfile', '-Command', '$PSVersionTable.PSVersion.ToString()')
    DotNetFramework   = Get-DotNetFrameworkInventory
    DotNet            = Get-DotNetRuntimeInventory
    CSharpViaAddType  = Test-AddTypeCapability
    WindowsForms      = Test-AssemblyCapability -Name 'WindowsForms' -AssemblyName 'System.Windows.Forms' -TypeName 'System.Windows.Forms.Form'
    WPF               = Test-AssemblyCapability -Name 'WPF' -AssemblyName 'PresentationFramework' -TypeName 'System.Windows.Window'
    UIAutomation      = Test-AssemblyCapability -Name 'UIAutomationClient' -AssemblyName 'UIAutomationClient' -TypeName 'System.Windows.Automation.AutomationElement'
    NamedPipe         = Test-ObjectCapability -Name 'NamedPipe' -Action { [System.IO.Pipes.NamedPipeServerStream]::new('survey-probe', [System.IO.Pipes.PipeDirection]::InOut, 1).Dispose(); 'NamedPipeServerStream' }
    TcpListener       = Test-ObjectCapability -Name 'TcpListener' -Action { $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0); try { $listener.Start(); ([System.Net.IPEndPoint]$listener.LocalEndpoint).Port } finally { $listener.Stop() } }
    HttpListener      = Test-ObjectCapability -Name 'HttpListener' -Action { $listener = [System.Net.HttpListener]::new(); try { 'Constructed' } finally { $listener.Close() } }
    Python            = Get-CommandInventory -Name 'python' -VersionArgs @('--version') -VersionPatterns @('Python\s+([0-9][0-9A-Za-z\.\-_]*)')
    PythonLauncher    = Get-CommandInventory -Name 'py' -VersionArgs @('--version') -VersionPatterns @('Python\s+([0-9][0-9A-Za-z\.\-_]*)')
    Node              = Get-CommandInventory -Name 'node' -VersionArgs @('--version') -VersionPatterns @('^v([0-9][0-9A-Za-z\.\-_]*)')
    Java              = Get-CommandInventory -Name 'java' -VersionArgs @('-version') -VersionPatterns @('version\s+"([^"]+)"')
    CScript           = [ordered]@{
        Name      = 'cscript'
        Available = [bool](Get-Command cscript.exe -ErrorAction SilentlyContinue)
        Path      = Mask-Path -Path ((Get-Command cscript.exe -ErrorAction SilentlyContinue | Select-Object -First 1).Source)
        Version   = $null
    }
    WScript           = [ordered]@{
        Name      = 'wscript'
        Available = [bool](Get-Command wscript.exe -ErrorAction SilentlyContinue)
        Path      = Mask-Path -Path ((Get-Command wscript.exe -ErrorAction SilentlyContinue | Select-Object -First 1).Source)
        Version   = $null
    }
}

Write-SurveySection 'Runtimes'
Write-SurveyItem -Name 'Windows PowerShell' -Status 'OK' -Detail "$($runtimes.PowerShell51.Version) [$($runtimes.PowerShell51.Path)]"
Write-SurveyItem -Name 'PowerShell 7' -Status $(if ($runtimes.PowerShell7.Available) { 'OK' } else { 'INFO' }) -Detail "$(if ($runtimes.PowerShell7.Available) { $runtimes.PowerShell7.Version } else { 'not found' })"
Write-SurveyItem -Name '.NET Framework' -Status $(if ($runtimes.DotNetFramework.Installed) { 'OK' } else { 'INFO' }) -Detail "$(if ($runtimes.DotNetFramework.Installed) { $runtimes.DotNetFramework.Version } else { 'not found' })"
Write-SurveyItem -Name '.NET' -Status $(if ($runtimes.DotNet.Available) { 'OK' } else { 'INFO' }) -Detail "$(if ($runtimes.DotNet.Available) { $runtimes.DotNet.Version } else { 'not found' })"
Write-SurveyItem -Name 'C# via Add-Type' -Status $(if ($runtimes.CSharpViaAddType.Available) { 'OK' } else { 'WARN' }) -Detail "$(if ($runtimes.CSharpViaAddType.Available) { 'available' } else { $runtimes.CSharpViaAddType.Detail })"
Write-SurveyItem -Name 'Windows Forms' -Status $(if ($runtimes.WindowsForms.Available) { 'OK' } else { 'INFO' }) -Detail "$(if ($runtimes.WindowsForms.Available) { $runtimes.WindowsForms.Detail } else { $runtimes.WindowsForms.Detail })"
Write-SurveyItem -Name 'WPF' -Status $(if ($runtimes.WPF.Available) { 'OK' } else { 'INFO' }) -Detail "$(if ($runtimes.WPF.Available) { $runtimes.WPF.Detail } else { $runtimes.WPF.Detail })"
Write-SurveyItem -Name 'UIAutomation' -Status $(if ($runtimes.UIAutomation.Available) { 'OK' } else { 'INFO' }) -Detail "$(if ($runtimes.UIAutomation.Available) { $runtimes.UIAutomation.Detail } else { $runtimes.UIAutomation.Detail })"
Write-SurveyItem -Name 'NamedPipe' -Status $(if ($runtimes.NamedPipe.Available) { 'OK' } else { 'WARN' }) -Detail "$($runtimes.NamedPipe.Detail)"
Write-SurveyItem -Name 'TcpListener' -Status $(if ($runtimes.TcpListener.Available) { 'OK' } else { 'WARN' }) -Detail "$($runtimes.TcpListener.Detail)"
Write-SurveyItem -Name 'HttpListener' -Status $(if ($runtimes.HttpListener.Available) { 'OK' } else { 'WARN' }) -Detail "$($runtimes.HttpListener.Detail)"
Write-SurveyItem -Name 'Python' -Status $(if ($runtimes.Python.Available) { 'OK' } else { 'INFO' }) -Detail "$(if ($runtimes.Python.Available) { $runtimes.Python.Version } else { 'not found' })"
Write-SurveyItem -Name 'py launcher' -Status $(if ($runtimes.PythonLauncher.Available) { 'OK' } else { 'INFO' }) -Detail "$(if ($runtimes.PythonLauncher.Available) { $runtimes.PythonLauncher.Version } else { 'not found' })"
Write-SurveyItem -Name 'Node.js' -Status $(if ($runtimes.Node.Available) { 'OK' } else { 'INFO' }) -Detail "$(if ($runtimes.Node.Available) { $runtimes.Node.Version } else { 'not found' })"
Write-SurveyItem -Name 'Java' -Status $(if ($runtimes.Java.Available) { 'OK' } else { 'INFO' }) -Detail "$(if ($runtimes.Java.Available) { $runtimes.Java.Version } else { 'not found' })"
Write-SurveyItem -Name 'CScript' -Status $(if ($runtimes.CScript.Available) { 'OK' } else { 'INFO' }) -Detail "$(if ($runtimes.CScript.Available) { $runtimes.CScript.Path } else { 'not found' })"
Write-SurveyItem -Name 'WScript' -Status $(if ($runtimes.WScript.Available) { 'OK' } else { 'INFO' }) -Detail "$(if ($runtimes.WScript.Available) { $runtimes.WScript.Path } else { 'not found' })"

$officeApps = @(
    [ordered]@{ Name = 'Excel';   Exe = 'excel.exe';   ProgId = 'Excel.Application' },
    [ordered]@{ Name = 'Word';    Exe = 'winword.exe'; ProgId = 'Word.Application' },
    [ordered]@{ Name = 'Outlook'; Exe = 'outlook.exe'; ProgId = 'Outlook.Application' },
    [ordered]@{ Name = 'Access';  Exe = 'msaccess.exe'; ProgId = 'Access.Application' },
    [ordered]@{ Name = 'PowerPoint'; Exe = 'powerpnt.exe'; ProgId = 'PowerPoint.Application' }
) | ForEach-Object {
    $path = Resolve-AppPath -ExeName $_.Exe
    $version = Get-FileVersionSummary -Path $path
    $com = Get-RegisteredComServer -ProgId $_.ProgId
    [ordered]@{
        Name             = $_.Name
        Installed        = [bool]$path
        Path             = Mask-Path -Path $path
        Version          = if ($version) { $version.ProductVersion } else { $null }
        AutomationProgId = $_.ProgId
        Automation       = $com.Registered
        Discovery        = 'AppPaths/Get-Command + COM registry'
    }
}

Write-SurveySection 'Office Hosts'
foreach ($app in $officeApps) {
    $status = if ($app.Installed -and $app.Automation) { 'OK' } elseif ($app.Installed) { 'WARN' } else { 'INFO' }
    Write-SurveyItem -Name $app.Name -Status $status -Detail "$(if ($app.Installed) { "installed version=$($app.Version) automation=$(if ($app.Automation) { 'registered' } else { 'not registered' })" } else { 'not found' })"
}

$acrobatInstall = Find-UninstallEntry -Patterns @('*Adobe Acrobat*', '*Acrobat DC*', '*Acrobat Pro*', '*Adobe Acrobat Reader*')
$acrobatPath = Resolve-AppPath -ExeName 'Acrobat.exe'
if (-not $acrobatPath) {
    $acrobatPath = Resolve-AppPath -ExeName 'AcroRd32.exe'
}
if (-not $acrobatPath -and $acrobatInstall -and $acrobatInstall.InstallLocation) {
    foreach ($exeName in @('Acrobat.exe', 'AcroRd32.exe')) {
        $candidate = Join-Path $acrobatInstall.InstallLocation $exeName
        if (Test-Path -LiteralPath $candidate) {
            $acrobatPath = $candidate
            break
        }
    }
}
$acrobatVersion = Get-FileVersionSummary -Path $acrobatPath
$acrobatProgIds = @('AcroExch.App', 'AcroExch.AVDoc', 'AcroExch.PDDoc') | ForEach-Object {
    Get-RegisteredComServer -ProgId $_
}

$acrobat = [ordered]@{
    InstalledProduct = if ($acrobatInstall) { $acrobatInstall.DisplayName } else { $null }
    InstalledVersion = if ($acrobatInstall) { $acrobatInstall.DisplayVersion } else { $null }
    ExecutablePath   = Mask-Path -Path $acrobatPath
    ExecutableVersion = if ($acrobatVersion) { $acrobatVersion.ProductVersion } else { $null }
    Automation       = $acrobatProgIds
    Discovery        = 'Uninstall registry + App Paths + COM registry'
}

Write-SurveySection 'Acrobat'
Write-SurveyItem -Name 'Product' -Status $(if ($acrobat.InstalledProduct) { 'OK' } else { 'INFO' }) -Detail "$(if ($acrobat.InstalledProduct) { "$($acrobat.InstalledProduct) $($acrobat.InstalledVersion)" } else { 'not found' })"
Write-SurveyItem -Name 'Executable' -Status $(if ($acrobat.ExecutablePath) { 'OK' } else { 'INFO' }) -Detail "$(if ($acrobat.ExecutablePath) { $acrobat.ExecutablePath } else { 'not found' })"
foreach ($api in $acrobat.Automation) {
    Write-SurveyItem -Name $api.ProgId -Status $(if ($api.Registered) { 'OK' } else { 'INFO' }) -Detail "$(if ($api.Registered) { $api.ServerPath } else { 'not registered' })"
}

$sources = [ordered]@{
    Machine = [ordered]@{
        Manufacturer = if ($computer) { 'CIM:Win32_ComputerSystem' } else { 'Registry:HKLM\\HARDWARE\\DESCRIPTION\\System\\BIOS' }
        Model = if ($computer) { 'CIM:Win32_ComputerSystem' } else { 'Registry:HKLM\\HARDWARE\\DESCRIPTION\\System\\BIOS' }
        SystemType = if ($computer) { 'CIM:Win32_ComputerSystem' } else { '.NET Environment' }
        BiosVersion = if ($bios) { 'CIM:Win32_BIOS' } else { 'Registry:HKLM\\HARDWARE\\DESCRIPTION\\System\\BIOS' }
        TotalMemoryGB = if ($computer) { 'CIM:Win32_ComputerSystem' } else { 'Microsoft.VisualBasic.Devices.ComputerInfo' }
        CpuName = if ($cpu) { 'CIM:Win32_Processor' } else { 'Registry:HKLM\\HARDWARE\\DESCRIPTION\\System\\CentralProcessor\\0' }
        GPUs = $gpuInventory.Source
        Drives = $driveInventory.Source
    }
    OS = [ordered]@{
        Caption = if ($os) { 'CIM:Win32_OperatingSystem' } else { 'Registry:HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion' }
        Version = if ($os) { 'CIM:Win32_OperatingSystem' } else { 'Registry:HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion' }
        BuildNumber = if ($os) { 'CIM:Win32_OperatingSystem' } else { 'Registry:HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion' }
        Architecture = if ($os) { 'CIM:Win32_OperatingSystem' } else { '.NET Environment' }
        InstallDate = if ($os) { 'CIM:Win32_OperatingSystem' } else { 'Registry:HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion' }
        LastBoot = if ($os) { 'CIM:Win32_OperatingSystem' } else { $null }
    }
    Network = [ordered]@{
        DomainJoined = if ($computer) { 'CIM:Win32_ComputerSystem' } else { $null }
        CurrentUserProxy = $currentUserProxy.Source
        WinHttpProxy = $winHttpProxy.Source
        OneDrive = $oneDrive.Source
    }
    Office = 'AppPaths/Get-Command + COM registry'
    Acrobat = $acrobat.Discovery
}

$report = [ordered]@{
    GeneratedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Machine     = $machine
    OS          = $osInfo
    Network     = $network
    Runtimes    = $runtimes
    Office      = $officeApps
    Acrobat     = $acrobat
    Sources     = $sources
    Diagnostics = [ordered]@{
        CimTimeoutSec = 5
        Items         = @($script:SurveyDiagnostics)
    }
}

$textLines = [System.Collections.Generic.List[string]]::new()
$textLines.Add('# Environment Survey')
$textLines.Add("Date: $($report.GeneratedAt)")
$textLines.Add('')
$textLines.Add('## Machine')
$textLines.Add("Computer: $($machine.ComputerName)")
$textLines.Add("Manufacturer: $($machine.Manufacturer)")
$textLines.Add("Model: $($machine.Model)")
$textLines.Add("System Type: $($machine.SystemType)")
$textLines.Add("BIOS: $($machine.BiosVersion)")
$textLines.Add("Memory: $($machine.TotalMemoryGB) GB")
$textLines.Add("CPU: $($machine.CpuName)")
$textLines.Add("CPU Cores/Threads: $($machine.CpuCores) / $($machine.CpuLogicalProcessors)")
$textLines.Add("GPU: $(if($machine.GPUs.Count -gt 0){($machine.GPUs -join '; ')}else{'(none reported)'})")
$textLines.Add("Machine Sources: memory=$($report.Sources.Machine.TotalMemoryGB), gpu=$(if($report.Sources.Machine.GPUs){$report.Sources.Machine.GPUs}else{'(none)'}), drives=$($report.Sources.Machine.Drives)")
$textLines.Add('')
$textLines.Add('Drives:')
foreach ($drive in $machine.Drives) {
    $textLines.Add("- $($drive.DeviceId) $($drive.FileSystem) Total=$($drive.SizeGB)GB Free=$($drive.FreeGB)GB")
}
$textLines.Add('')
$textLines.Add('## OS')
$textLines.Add("Caption: $($osInfo.Caption)")
$textLines.Add("Version: $($osInfo.Version)")
$textLines.Add("Build: $($osInfo.BuildNumber)")
$textLines.Add("Architecture: $($osInfo.Architecture)")
$textLines.Add("Installed: $($osInfo.InstallDate)")
$textLines.Add("Last Boot: $($osInfo.LastBoot)")
$textLines.Add("OS Sources: caption=$($report.Sources.OS.Caption), version=$($report.Sources.OS.Version)")
$textLines.Add('')
$textLines.Add('## Network')
$textLines.Add("Domain Joined: $(if($network.DomainJoined -eq $true){'yes'}elseif($network.DomainJoined -eq $false){'no'}else{'unknown'})")
if ($network.DomainName) {
    $textLines.Add("Domain Name: $($network.DomainName)")
}
$textLines.Add("Current User Proxy: enabled=$($network.CurrentUserProxy.ProxyEnabled) autodetect=$($network.CurrentUserProxy.AutoDetect) server=$($network.CurrentUserProxy.ProxyServer)")
if ($network.CurrentUserProxy.AutoConfigUrl) {
    $textLines.Add("Current User Proxy PAC: $($network.CurrentUserProxy.AutoConfigUrl)")
}
$textLines.Add("OneDrive Installed: $($network.OneDrive.Installed)")
$textLines.Add("OneDrive Accounts: total=$($network.OneDrive.AccountCount) business=$($network.OneDrive.BusinessAccountCount)")
$textLines.Add("SharePoint Sync Likely: $($network.OneDrive.SharePointSyncLikely)")
if ($network.OneDrive.SyncRoots.Count -gt 0) {
    $textLines.Add('OneDrive Sync Roots:')
    foreach ($root in $network.OneDrive.SyncRoots) {
        $textLines.Add("- $root")
    }
}
if ($network.OneDrive.EnvironmentVariables.Count -gt 0) {
    $textLines.Add('OneDrive Environment Variables:')
    foreach ($row in $network.OneDrive.EnvironmentVariables) {
        $textLines.Add("- $($row.Name): $($row.Path) exists=$($row.Exists)")
    }
}
if ($network.WinHttpProxy.Lines.Count -gt 0) {
    $textLines.Add('WinHTTP Proxy:')
    foreach ($line in $network.WinHttpProxy.Lines) {
        $textLines.Add("- $line")
    }
}
$textLines.Add("Network Sources: domain=$($report.Sources.Network.DomainJoined), proxy=$($report.Sources.Network.CurrentUserProxy), onedrive=$($report.Sources.Network.OneDrive)")
$textLines.Add('')
$textLines.Add('## Diagnostics')
$textLines.Add("CIM Timeout: $($report.Diagnostics.CimTimeoutSec) sec")
if ($report.Diagnostics.Items.Count -gt 0) {
    foreach ($item in $report.Diagnostics.Items) {
        $textLines.Add("- [$($item.Category)] $($item.Target): $($item.Message)")
    }
} else {
    $textLines.Add('- none')
}
$textLines.Add('')
$textLines.Add('## Runtimes')
$textLines.Add("Windows PowerShell: $($runtimes.PowerShell51.Version) [$($runtimes.PowerShell51.Path)]")
$textLines.Add("PowerShell 7: $(if($runtimes.PowerShell7.Available){$runtimes.PowerShell7.Version}else{'not found'})")
$textLines.Add(".NET Framework: $(if($runtimes.DotNetFramework.Installed){$runtimes.DotNetFramework.Version}else{'not found'})")
$textLines.Add(".NET: $(if($runtimes.DotNet.Available){$runtimes.DotNet.Version}else{'not found'})")
$textLines.Add("C# via Add-Type: $(if($runtimes.CSharpViaAddType.Available){'available'}else{'not available'})")
$textLines.Add("Windows Forms: $(if($runtimes.WindowsForms.Available){$runtimes.WindowsForms.Detail}else{'not available'})")
$textLines.Add("WPF: $(if($runtimes.WPF.Available){$runtimes.WPF.Detail}else{'not available'})")
$textLines.Add("UIAutomation: $(if($runtimes.UIAutomation.Available){$runtimes.UIAutomation.Detail}else{'not available'})")
$textLines.Add("NamedPipe: $(if($runtimes.NamedPipe.Available){$runtimes.NamedPipe.Detail}else{'not available'})")
$textLines.Add("TcpListener: $(if($runtimes.TcpListener.Available){$runtimes.TcpListener.Detail}else{'not available'})")
$textLines.Add("HttpListener: $(if($runtimes.HttpListener.Available){$runtimes.HttpListener.Detail}else{'not available'})")
$textLines.Add("Python: $(if($runtimes.Python.Available){$runtimes.Python.Version}else{'not found'})")
$textLines.Add("py launcher: $(if($runtimes.PythonLauncher.Available){$runtimes.PythonLauncher.Version}else{'not found'})")
$textLines.Add("Node.js: $(if($runtimes.Node.Available){$runtimes.Node.Version}else{'not found'})")
$textLines.Add("Java: $(if($runtimes.Java.Available){$runtimes.Java.Version}else{'not found'})")
$textLines.Add("CScript: $(if($runtimes.CScript.Available){$runtimes.CScript.Path}else{'not found'})")
$textLines.Add("WScript: $(if($runtimes.WScript.Available){$runtimes.WScript.Path}else{'not found'})")

if ($runtimes.DotNet.Available -and $runtimes.DotNet.Runtimes.Count -gt 0) {
    $textLines.Add('')
    $textLines.Add('.NET runtimes:')
    foreach ($runtime in $runtimes.DotNet.Runtimes) {
        $textLines.Add("- $runtime")
    }
}

if ($runtimes.DotNet.Available -and $runtimes.DotNet.Sdks.Count -gt 0) {
    $textLines.Add('')
    $textLines.Add('.NET SDKs:')
    foreach ($sdk in $runtimes.DotNet.Sdks) {
        $textLines.Add("- $sdk")
    }
}

$textLines.Add('')
$textLines.Add('## Office Hosts')
foreach ($app in $officeApps) {
    $status = if ($app.Installed) { 'installed' } else { 'not found' }
    $automation = if ($app.Automation) { 'registered' } else { 'not registered' }
    $textLines.Add("- $($app.Name): $status / version=$($app.Version) / automation=$automation")
    if ($app.Path) {
        $textLines.Add("  path: $($app.Path)")
    }
}

$textLines.Add('')
$textLines.Add('## Acrobat')
$textLines.Add("Product: $(if($acrobat.InstalledProduct){$acrobat.InstalledProduct}else{'not found'})")
$textLines.Add("Version: $($acrobat.InstalledVersion)")
$textLines.Add("Executable: $($acrobat.ExecutablePath)")
$textLines.Add("Executable Version: $($acrobat.ExecutableVersion)")
$textLines.Add("Discovery: $($acrobat.Discovery)")
$textLines.Add('Automation:')
foreach ($api in $acrobat.Automation) {
    $textLines.Add("- $($api.ProgId): $(if($api.Registered){'registered'}else{'not registered'})")
    if ($api.ServerPath) {
        $textLines.Add("  server: $($api.ServerPath)")
    }
}

Write-SurveySection 'Diagnostics'
if ($report.Diagnostics.Items.Count -gt 0) {
    Write-SurveyItem -Name 'Diagnostics' -Status 'WARN' -Detail "$($report.Diagnostics.Items.Count) item(s)"
} else {
    Write-SurveyItem -Name 'Diagnostics' -Status 'OK' -Detail 'none'
}

$utf8Bom = New-Object System.Text.UTF8Encoding $true
if ($ReturnTextFile) {
    [System.IO.File]::WriteAllText($ReturnTextFile, ($textLines -join "`r`n"), $utf8Bom)
} else {
    $txtPath = Join-Path $outDir 'survey.txt'
    [System.IO.File]::WriteAllLines($txtPath, $textLines, $utf8Bom)
}

Write-Host ''
Write-Host '=== Survey Complete ===' -ForegroundColor Cyan
Write-Host "Time        : $([Math]::Round($script:SurveyStopwatch.Elapsed.TotalSeconds, 1))s" -ForegroundColor Gray
if (-not $ReturnTextFile) {
    Write-Host "Output folder: $outDir" -ForegroundColor Gray
    Write-Host "Report      : $txtPath" -ForegroundColor Green
}
