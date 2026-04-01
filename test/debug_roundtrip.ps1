param([string]$InputFile = 'test\test_sample.xlsm')

$ErrorActionPreference = 'Stop'
$modulePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'lib\VBAToolkit.psm1'
Import-Module $modulePath -Force -DisableNameChecking

$inputPath = (Resolve-Path $InputFile).Path
$baseName = [IO.Path]::GetFileNameWithoutExtension($inputPath)
$outDir = Join-Path $PSScriptRoot 'debug_output'
if (-not (Test-Path $outDir)) { New-Item $outDir -ItemType Directory -Force | Out-Null }

Write-Host "=== OLE2 Round-trip Debug ===" -ForegroundColor Cyan
Write-Host "Input: $inputPath" -ForegroundColor Gray

# Read project with raw data
$project = Get-AllModuleCode $inputPath -IncludeRawData
$encoding = [System.Text.Encoding]::GetEncoding($project.Codepage)

Write-Host "`nModules:" -ForegroundColor Yellow
foreach ($modName in $project.Modules.Keys) {
    $md = $project.Modules[$modName]
    if (-not $md.Entry) { Write-Host "  $modName : (no entry)"; continue }
    $streamLen = $md.StreamData.Length
    Write-Host "  $modName : Offset=$($md.Offset) Stream=$streamLen P-code=$($md.Offset) Compressed=$($streamLen - $md.Offset)"

    $decompressed = Decompress-VBA $md.StreamData $md.Offset
    $recompressed = Compress-VBA $decompressed
    $origCompLen = $streamLen - $md.Offset
    Write-Host "    Decompressed=$($decompressed.Length) Recompressed=$($recompressed.Length) OrigCompressed=$origCompLen"
    Write-Host "    Size match: $(if ($recompressed.Length -eq $origCompLen) { 'YES' } else { 'NO (diff=' + ($recompressed.Length - $origCompLen) + ')' })"
}

# --- Test A: Pure copy ---
Write-Host "`n--- A: Pure copy ---" -ForegroundColor Cyan
$testA = Join-Path $outDir "${baseName}_A_copy.xlsm"
Copy-Item $inputPath $testA -Force

# --- Test B: Keep p-code, recompress source ---
Write-Host "--- B: Keep p-code, recompress ---" -ForegroundColor Cyan
$testB = Join-Path $outDir "${baseName}_B_keepPcode.xlsm"
Copy-Item $inputPath $testB -Force
$ole2B = [byte[]]$project.Ole2Bytes.Clone()

foreach ($modName in $project.Modules.Keys) {
    $md = $project.Modules[$modName]
    if (-not $md.Entry) { continue }
    $decompressed = Decompress-VBA $md.StreamData $md.Offset
    $recompressed = Compress-VBA $decompressed

    $newStream = New-Object byte[] ($md.Offset + $recompressed.Length)
    [Array]::Copy($md.StreamData, 0, $newStream, 0, $md.Offset)  # keep p-code
    [Array]::Copy($recompressed, 0, $newStream, $md.Offset, $recompressed.Length)

    Write-Host "  $modName : $($md.StreamData.Length) -> $($newStream.Length)"
    Write-Ole2Stream $ole2B $project.Ole2 $md.Entry $newStream
}
Save-VbaProjectBytes $testB $ole2B $project.IsZip

# --- Test C: Zero p-code, recompress ---
Write-Host "--- C: Zero p-code, recompress ---" -ForegroundColor Cyan
$testC = Join-Path $outDir "${baseName}_C_zeroPcode.xlsm"
Copy-Item $inputPath $testC -Force
$ole2C = [byte[]]$project.Ole2Bytes.Clone()

foreach ($modName in $project.Modules.Keys) {
    $md = $project.Modules[$modName]
    if (-not $md.Entry) { continue }
    $decompressed = Decompress-VBA $md.StreamData $md.Offset
    $recompressed = Compress-VBA $decompressed

    $newStream = New-Object byte[] ($md.Offset + $recompressed.Length)
    # p-code stays zero
    [Array]::Copy($recompressed, 0, $newStream, $md.Offset, $recompressed.Length)

    Write-Host "  $modName : $($md.StreamData.Length) -> $($newStream.Length) (p-code zeroed)"
    Write-Ole2Stream $ole2C $project.Ole2 $md.Entry $newStream
}
Save-VbaProjectBytes $testC $ole2C $project.IsZip

# --- Test D: Keep entire original stream, only save/reload OLE2 ---
Write-Host "--- D: Write original streams back (identity write) ---" -ForegroundColor Cyan
$testD = Join-Path $outDir "${baseName}_D_identity.xlsm"
Copy-Item $inputPath $testD -Force
$ole2D = [byte[]]$project.Ole2Bytes.Clone()

foreach ($modName in $project.Modules.Keys) {
    $md = $project.Modules[$modName]
    if (-not $md.Entry) { continue }
    # Write back exact same stream data
    Write-Ole2Stream $ole2D $project.Ole2 $md.Entry $md.StreamData
    Write-Host "  $modName : $($md.StreamData.Length) (unchanged)"
}
Save-VbaProjectBytes $testD $ole2D $project.IsZip

# --- Verify all ---
Write-Host "`n=== Verification (Excel GUI with alerts) ===" -ForegroundColor Cyan
foreach ($info in @(
    @{ Label = 'A: pure copy'; Path = $testA },
    @{ Label = 'B: keep p-code + recompress'; Path = $testB },
    @{ Label = 'C: zero p-code + recompress'; Path = $testC },
    @{ Label = 'D: identity write'; Path = $testD }
)) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $true
    $status = ''
    try {
        $wb = $excel.Workbooks.Open($info.Path)
        try {
            $cnt = $wb.VBProject.VBComponents.Count
            $status = "OK ($cnt components)"
        } catch {
            $status = "Open OK, VBA: $($_.Exception.Message)"
        }
        $wb.Close($false)
    } catch {
        $status = "FAIL: $($_.Exception.Message)"
    } finally {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    $color = if ($status.StartsWith('OK')) { 'Green' } else { 'Red' }
    Write-Host "  $($info.Label): $status" -ForegroundColor $color
}

Write-Host "`nFiles: $outDir" -ForegroundColor Gray
