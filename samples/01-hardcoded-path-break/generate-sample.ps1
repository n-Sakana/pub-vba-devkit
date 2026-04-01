$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outDir = Join-Path $scriptDir 'out'
$modulePath = Join-Path $scriptDir 'vba\PathBreakDemo.bas'

$syncRoot = if ($env:OneDriveCommercial) { $env:OneDriveCommercial } elseif ($env:OneDrive) { $env:OneDrive } else { throw 'OneDriveCommercial / OneDrive が見つかりません。' }
$actualDataPath = Join-Path $syncRoot '_vba_devkit_samples\SharePointDemo\Shared Documents\案件データ'
$workbookFolder = Join-Path $syncRoot '_vba_devkit_samples\SharePointDemo\Shared Documents\MacroHost'
$workbookPath = Join-Path $workbookFolder 'HardcodedPathBreakDemo.xlsm'

New-Item -ItemType Directory -Force -Path $outDir, $actualDataPath, $workbookFolder | Out-Null
Set-Content -Path (Join-Path $actualDataPath '案件A.txt') -Value 'sample-a' -Encoding UTF8
Set-Content -Path (Join-Path $actualDataPath '案件B.txt') -Value 'sample-b' -Encoding UTF8

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Demo'
    $ws.Range('A1').Value2 = 'Run PathBreakDemo.RunHardcodedPathDemo'
    $ws.Range('A2').Value2 = 'This sample intentionally fails for hard-coded and ThisWorkbook.Path-based folder resolution.'
    $null = $wb.VBProject.VBComponents.Import($modulePath)
    $fileFormat = 52
    $wb.SaveAs($workbookPath, $fileFormat)
    $wb.Close($true)
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host "Created: $workbookPath"
Write-Host "Actual synced data path: $actualDataPath"
Write-Host "Workbook folder: $workbookFolder"
Write-Host 'Open the workbook and run RunHardcodedPathDemo.'
