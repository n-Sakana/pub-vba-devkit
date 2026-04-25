$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outDir = Join-Path $scriptDir 'out'
$workbookPath = Join-Path $outDir 'IpcHttpClientDemo.xlsm'
$modulePath = Join-Path $scriptDir 'vba\IpcHttpClientDemo.bas'

New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Demo'
    $ws.Range('A1').Value2 = 'Run IpcHttpClientDemo.RunVisualWin32Demo'
    $ws.Range('A2').Value2 = 'Requires helper server on http://127.0.0.1:8765/'
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
Write-Host 'Open the workbook after starting the helper, then run RunVisualWin32Demo.'
