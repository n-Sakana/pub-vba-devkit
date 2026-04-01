param([string]$OutputDirectory)

$ErrorActionPreference = 'Stop'

$outDir = if ($OutputDirectory) { $OutputDirectory } else { Join-Path $PSScriptRoot 'output' }
if (-not (Test-Path $outDir)) { New-Item $outDir -ItemType Directory -Force | Out-Null }
$outDir = (Resolve-Path $outDir).Path
$outPath = Join-Path $outDir 'demo_path_problem.xlsm'

$vbaCode = @'
Option Explicit

Sub Demo_PathProblem()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.Clear

    ws.Cells(1, 1).Value = "Path Problem Demo"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Range("A1:D1").Merge

    ws.Cells(2, 1).Value = "OneDrive sync folder: ThisWorkbook.Path / hardcoded path issues"
    ws.Cells(2, 1).Font.Color = RGB(128, 128, 128)
    ws.Range("A2:D2").Merge

    Dim r As Long: r = 4
    ws.Cells(r, 1).Value = "Test": ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Value = "Code": ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 3).Value = "Result": ws.Cells(r, 3).Font.Bold = True
    ws.Cells(r, 4).Value = "Note": ws.Cells(r, 4).Font.Bold = True
    r = r + 1

    On Error Resume Next

    ' Test 1: ThisWorkbook.Path
    ws.Cells(r, 1).Value = "ThisWorkbook.Path"
    ws.Cells(r, 2).Value = "ThisWorkbook.Path"
    ws.Cells(r, 3).Value = ThisWorkbook.Path
    If Left(ThisWorkbook.Path, 5) = "https" Then
        ws.Cells(r, 4).Value = "URL (OneDrive/SharePoint)"
        ws.Cells(r, 3).Interior.Color = RGB(255, 255, 200)
    Else
        ws.Cells(r, 4).Value = "Local path"
        ws.Cells(r, 3).Interior.Color = RGB(200, 255, 200)
    End If
    r = r + 1

    ' Test 2: Dir(ThisWorkbook.Path)
    ws.Cells(r, 1).Value = "Dir(TWB.Path)"
    ws.Cells(r, 2).Value = "Dir(ThisWorkbook.Path & ""\*.*"")"
    Dim f As String: f = Dir(ThisWorkbook.Path & "\*.*")
    If Err.Number <> 0 Then
        ws.Cells(r, 3).Value = "ERROR: " & Err.Description
        ws.Cells(r, 3).Interior.Color = RGB(255, 200, 200)
        ws.Cells(r, 4).Value = "Dir() fails on URL path"
        Err.Clear
    Else
        ws.Cells(r, 3).Value = f
        ws.Cells(r, 3).Interior.Color = RGB(200, 255, 200)
    End If
    r = r + 1

    ' Test 3: FSO with ThisWorkbook.Path
    ws.Cells(r, 1).Value = "FSO(TWB.Path)"
    ws.Cells(r, 2).Value = "FSO.GetFolder(ThisWorkbook.Path)"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object: Set folder = fso.GetFolder(ThisWorkbook.Path)
    If Err.Number <> 0 Then
        ws.Cells(r, 3).Value = "ERROR: " & Err.Description
        ws.Cells(r, 3).Interior.Color = RGB(255, 200, 200)
        ws.Cells(r, 4).Value = "FSO fails on URL path"
        Err.Clear
    Else
        ws.Cells(r, 3).Value = "Files=" & folder.Files.Count & " Folders=" & folder.SubFolders.Count
        ws.Cells(r, 3).Interior.Color = RGB(200, 255, 200)
    End If
    Set folder = Nothing: Set fso = Nothing
    r = r + 1

    ' Test 4: Hardcoded path
    ws.Cells(r, 1).Value = "Hardcoded path"
    ws.Cells(r, 2).Value = "Dir(""C:\Users\"" & user & ""\Desktop\*.*"")"
    Dim desktop As String: desktop = "C:\Users\" & Environ$("USERNAME") & "\Desktop"
    f = Dir(desktop & "\*.*")
    If Err.Number <> 0 Then
        ws.Cells(r, 3).Value = "ERROR: " & Err.Description
        ws.Cells(r, 3).Interior.Color = RGB(255, 200, 200)
        ws.Cells(r, 4).Value = "Desktop may be redirected to OneDrive"
        Err.Clear
    Else
        ws.Cells(r, 3).Value = f
        ws.Cells(r, 3).Interior.Color = RGB(200, 255, 200)
        ws.Cells(r, 4).Value = "Works locally but environment-dependent"
    End If

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
    $ws.Name = 'PathProblem'
    $ws.Cells.Item(1, 1) = 'Run Demo_PathProblem (Alt+F8)'

    $wb.SaveAs($outPath, 52)
    $wb.Close($false)
    Write-Host "Generated: $outPath" -ForegroundColor Green
} finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
