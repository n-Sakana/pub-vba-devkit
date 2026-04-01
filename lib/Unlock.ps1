param([Parameter(Mandatory)][string]$FilePath)
$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force -DisableNameChecking

$FilePath = Resolve-VbaFilePath $FilePath
$fileName = [IO.Path]::GetFileName($FilePath)
$ext = [IO.Path]::GetExtension($FilePath).ToLower()
$sw = [System.Diagnostics.Stopwatch]::StartNew()

Write-VbaHeader 'Unlock' $fileName
Write-VbaLog 'Unlock' $FilePath 'Started'

# Non-destructive: copy to output folder
$outDir = New-VbaOutputDir $FilePath 'unlock'
$copyPath = Join-Path $outDir $fileName
Copy-Item $FilePath $copyPath -Force
Write-VbaStatus 'Unlock' $fileName "Copy created in output folder"

function Find-DPB([byte[]]$data) {
    $pattern = [System.Text.Encoding]::ASCII.GetBytes('DPB=')
    for ($i = 0; $i -le $data.Length - $pattern.Length; $i++) {
        $match = $true
        for ($j = 0; $j -lt $pattern.Length; $j++) {
            if ($data[$i + $j] -ne $pattern[$j]) { $match = $false; break }
        }
        if ($match) { return $i }
    }
    return -1
}

function Patch-XlsFile([string]$path) {
    $data = [IO.File]::ReadAllBytes($path)
    $pos = Find-DPB $data
    if ($pos -eq -1) { return 'not_found' }
    $data[$pos + 2] = 0x78  # Change 'B' to 'x' in DPB= -> DPx= to invalidate password hash
    [IO.File]::WriteAllBytes($path, $data)
    return 'patched'
}

if ($ext -eq '.xls') {
    $result = Patch-XlsFile $copyPath
} else {
    $tempXls = Join-Path ([IO.Path]::GetTempPath()) "VBAUnlock_$(Get-Date -Format yyyyMMddHHmmss).xls"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    try {
        $wb = $excel.Workbooks.Open($copyPath, 0, $false)
        $wb.SaveAs($tempXls, 56)  # xlExcel8
        $wb.Close($false)

        $result = Patch-XlsFile $tempXls

        if ($result -eq 'patched') {
            $wb = $excel.Workbooks.Open($tempXls, 0, $false)
            if ($ext -eq '.xlam') { $wb.SaveAs($copyPath, 55) }
            else { $wb.SaveAs($copyPath, 52) }
            $wb.Close($false)
        }
    }
    catch {
        $result = 'error'
        $errorMsg = $_.Exception.Message
    }
    finally {
        if ($wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null; $wb = $null }
        try { $excel.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        Remove-Item $tempXls -Force -ErrorAction SilentlyContinue
    }
}

$sw.Stop()

if ($result -eq 'patched') {
    Write-VbaResult 'Unlock' $fileName "Password protection disabled" $outDir $sw.Elapsed.TotalSeconds
    Write-Host ""
    Write-Host "  To fully remove, open the unlocked copy and:" -ForegroundColor Gray
    Write-Host "  1. Open VBE (Alt+F11)" -ForegroundColor Gray
    Write-Host "  2. Tools > VBAProject Properties > Protection tab" -ForegroundColor Gray
    Write-Host "  3. Clear the password fields and click OK" -ForegroundColor Gray
    Write-Host "  4. Save the file" -ForegroundColor Gray
    Write-VbaLog 'Unlock' $FilePath "Patched | -> $outDir"
} elseif ($result -eq 'error') {
    Remove-Item $copyPath -Force -ErrorAction SilentlyContinue
    Write-VbaError 'Unlock' $fileName "Failed to process: $errorMsg"
} else {
    Remove-Item $copyPath -Force -ErrorAction SilentlyContinue
    Write-VbaStatus 'Unlock' $fileName "No VBA password hash (DPB=) found"
    Write-VbaLog 'Unlock' $FilePath 'No DPB= found'
}
