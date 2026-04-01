param([int]$Port = 8899)

$ErrorActionPreference = 'Stop'

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class Win32Demo
{
    [StructLayout(LayoutKind.Sequential)]
    public struct POINT { public int X; public int Y; }

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out POINT pt);

    [DllImport("user32.dll")]
    public static extern int GetSystemMetrics(int nIndex);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder sb, int maxCount);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

    [DllImport("user32.dll")]
    public static extern ulong GetTickCount64();

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr FindWindowW(string className, string windowName);

    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
}
'@ -Language CSharp

$prefix = "http://127.0.0.1:${Port}/"
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add($prefix)
$listener.Start()

Write-Host "=== Backend Server ===" -ForegroundColor Cyan
Write-Host "Listening on $prefix" -ForegroundColor Green
Write-Host "Endpoints:" -ForegroundColor Gray
Write-Host "  GET  /api/sysinfo  - System info via Win32 API" -ForegroundColor Gray
Write-Host "  POST /api/wiggle   - Wiggle Excel window via SetWindowPos" -ForegroundColor Gray
Write-Host "  GET  /api/shutdown - Stop server" -ForegroundColor Gray
Write-Host ""

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $path = $context.Request.Url.AbsolutePath
        $response = $context.Response
        $json = '{}'

        switch ($path) {
            '/api/sysinfo' {
                $cursor = New-Object Win32Demo+POINT
                [Win32Demo]::GetCursorPos([ref]$cursor) | Out-Null
                $screenW = [Win32Demo]::GetSystemMetrics(0)
                $screenH = [Win32Demo]::GetSystemMetrics(1)
                $fg = [Win32Demo]::GetForegroundWindow()
                $sb = New-Object System.Text.StringBuilder 256
                [Win32Demo]::GetWindowText($fg, $sb, $sb.Capacity) | Out-Null
                $ticks = [Win32Demo]::GetTickCount64()

                $title = $sb.ToString().Replace('\', '\\').Replace('"', '\"')
                $json = "{`"screenWidth`":$screenW,`"screenHeight`":$screenH,`"cursorX`":$($cursor.X),`"cursorY`":$($cursor.Y),`"foregroundWindow`":`"$title`",`"uptimeMs`":$ticks}"
            }
            '/api/wiggle' {
                $excelHwnd = [Win32Demo]::FindWindowW('XLMAIN', $null)
                if ($excelHwnd -ne [IntPtr]::Zero) {
                    $rect = New-Object Win32Demo+RECT
                    [Win32Demo]::GetWindowRect($excelHwnd, [ref]$rect) | Out-Null
                    $flags = 0x0004 -bor 0x0001  # SWP_NOZORDER | SWP_NOSIZE
                    # Gentle wiggle: right, left, back to original
                    [Win32Demo]::SetWindowPos($excelHwnd, [IntPtr]::Zero, $rect.Left + 15, $rect.Top, 0, 0, $flags) | Out-Null
                    Start-Sleep -Milliseconds 60
                    [Win32Demo]::SetWindowPos($excelHwnd, [IntPtr]::Zero, $rect.Left - 15, $rect.Top, 0, 0, $flags) | Out-Null
                    Start-Sleep -Milliseconds 60
                    [Win32Demo]::SetWindowPos($excelHwnd, [IntPtr]::Zero, $rect.Left + 8, $rect.Top, 0, 0, $flags) | Out-Null
                    Start-Sleep -Milliseconds 50
                    [Win32Demo]::SetWindowPos($excelHwnd, [IntPtr]::Zero, $rect.Left - 8, $rect.Top, 0, 0, $flags) | Out-Null
                    Start-Sleep -Milliseconds 50
                    [Win32Demo]::SetWindowPos($excelHwnd, [IntPtr]::Zero, $rect.Left, $rect.Top, 0, 0, $flags) | Out-Null
                    $json = '{"result":"ok","message":"Window wiggled"}'
                } else {
                    $json = '{"result":"fail","message":"Excel window not found"}'
                }
            }
            '/api/shutdown' {
                $json = '{"result":"ok"}'
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
                $response.ContentType = 'application/json'
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.Close()
                Write-Host "  $(Get-Date -Format 'HH:mm:ss') /api/shutdown" -ForegroundColor Yellow
                break
            }
            default {
                $json = '{"error":"unknown endpoint"}'
                $response.StatusCode = 404
            }
        }

        if ($path -eq '/api/shutdown') { break }

        $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
        $response.ContentType = 'application/json'
        $response.ContentLength64 = $buffer.Length
        $response.OutputStream.Write($buffer, 0, $buffer.Length)
        $response.Close()

        Write-Host "  $(Get-Date -Format 'HH:mm:ss') $path" -ForegroundColor DarkGray
    }
} finally {
    $listener.Stop()
    $listener.Close()
    Write-Host "Server stopped." -ForegroundColor Yellow
}
