param([string]$OutputDirectory)

$ErrorActionPreference = 'Stop'

$outDir = if ($OutputDirectory) { $OutputDirectory } else { Join-Path $PSScriptRoot 'output' }
if (-not (Test-Path $outDir)) { New-Item $outDir -ItemType Directory -Force | Out-Null }
$outDir = (Resolve-Path $outDir).Path
$outPath = Join-Path $outDir 'demo_environ_solution.xlsm'

$vbaCode = @'
Option Explicit

Sub Demo_EnvironSolution()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.Clear

    ws.Cells(1, 1).Value = "Environ$ Solution Demo"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Range("A1:D1").Merge

    ws.Cells(2, 1).Value = "Resolve OneDrive sync root via Environ$ instead of ThisWorkbook.Path"
    ws.Cells(2, 1).Font.Color = RGB(128, 128, 128)
    ws.Range("A2:D2").Merge

    Dim r As Long: r = 4
    ws.Cells(r, 1).Value = "Test": ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Value = "Code": ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 3).Value = "Result": ws.Cells(r, 3).Font.Bold = True
    ws.Cells(r, 4).Value = "Note": ws.Cells(r, 4).Font.Bold = True
    r = r + 1

    On Error Resume Next

    ' Get sync root
    Dim syncRoot As String
    syncRoot = Environ$("OneDriveCommercial")
    ws.Cells(r, 1).Value = "OneDriveCommercial"
    ws.Cells(r, 2).Value = "Environ$(""OneDriveCommercial"")"
    If Len(syncRoot) > 0 Then
        ws.Cells(r, 3).Value = syncRoot
        ws.Cells(r, 3).Interior.Color = RGB(200, 255, 200)
        ws.Cells(r, 4).Value = "OK - sync root found"
    Else
        ws.Cells(r, 3).Value = "(empty)"
        ws.Cells(r, 3).Interior.Color = RGB(255, 255, 200)
    End If
    r = r + 1

    Dim odEnv As String: odEnv = Environ$("OneDrive")
    ws.Cells(r, 1).Value = "OneDrive"
    ws.Cells(r, 2).Value = "Environ$(""OneDrive"")"
    If Len(odEnv) > 0 Then
        ws.Cells(r, 3).Value = odEnv
        ws.Cells(r, 3).Interior.Color = RGB(200, 255, 200)
    Else
        ws.Cells(r, 3).Value = "(empty)"
        ws.Cells(r, 3).Interior.Color = RGB(255, 255, 200)
    End If
    r = r + 1

    ' Use best available
    If Len(syncRoot) = 0 Then syncRoot = odEnv
    If Len(syncRoot) = 0 Then
        ws.Cells(r, 1).Value = "Result"
        ws.Cells(r, 3).Value = "No OneDrive environment variable set"
        ws.Cells(r, 3).Interior.Color = RGB(255, 200, 200)
        ws.Columns("A:D").AutoFit
        Exit Sub
    End If

    r = r + 1

    ' Dir via Environ$
    ws.Cells(r, 1).Value = "Dir(syncRoot)"
    ws.Cells(r, 2).Value = "Dir(Environ$(""..."") & ""\*.*"")"
    Dim f As String: f = Dir(syncRoot & "\*.*")
    If Err.Number <> 0 Then
        ws.Cells(r, 3).Value = "ERROR: " & Err.Description
        ws.Cells(r, 3).Interior.Color = RGB(255, 200, 200)
        Err.Clear
    Else
        ws.Cells(r, 3).Value = f
        ws.Cells(r, 3).Interior.Color = RGB(200, 255, 200)
        ws.Cells(r, 4).Value = "Dir() works with local sync path"
    End If
    r = r + 1

    ' FSO via Environ$
    ws.Cells(r, 1).Value = "FSO(syncRoot)"
    ws.Cells(r, 2).Value = "FSO.GetFolder(Environ$(""...""))"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object: Set folder = fso.GetFolder(syncRoot)
    If Err.Number <> 0 Then
        ws.Cells(r, 3).Value = "ERROR: " & Err.Description
        ws.Cells(r, 3).Interior.Color = RGB(255, 200, 200)
        Err.Clear
    Else
        ws.Cells(r, 3).Value = "Files=" & folder.Files.Count & " Folders=" & folder.SubFolders.Count
        ws.Cells(r, 3).Interior.Color = RGB(200, 255, 200)
        ws.Cells(r, 4).Value = "FSO works with local sync path"
    End If
    Set folder = Nothing: Set fso = Nothing
    r = r + 2

    ' File listing
    ws.Cells(r, 1).Value = "Files in sync root (first 10):"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    Dim cnt As Long: cnt = 0
    Dim entry As String: entry = Dir(syncRoot & "\*.*")
    Do While Len(entry) > 0 And cnt < 10
        cnt = cnt + 1
        ws.Cells(r, 1).Value = cnt
        ws.Cells(r, 2).Value = entry
        r = r + 1
        entry = Dir()
    Loop

    ws.Columns("A:D").AutoFit
    ws.Columns("C:C").ColumnWidth = 50
End Sub
'@

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false

try {
    $wb = $excel.Workbooks.Add()
    $mod = $wb.VBProject.VBComponents.Add(1)
    $mod.Name = 'DemoModule'
    $mod.CodeModule.AddFromString($vbaCode)

    $ws = $wb.Sheets(1)
    $ws.Name = 'EnvironSolution'
    $ws.Cells.Item(1, 1) = 'Run Demo_EnvironSolution (Alt+F8)'

    $wb.SaveAs($outPath, 52)
    $wb.Close($false)
    Write-Host "Generated: $outPath" -ForegroundColor Green
} finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
