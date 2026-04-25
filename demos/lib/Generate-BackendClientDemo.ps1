param([string]$OutputDirectory)

$ErrorActionPreference = 'Stop'

$outDir = if ($OutputDirectory) { $OutputDirectory } else { Join-Path $PSScriptRoot 'output' }
if (-not (Test-Path $outDir)) { New-Item $outDir -ItemType Directory -Force | Out-Null }
$outDir = (Resolve-Path $outDir).Path
$outPath = Join-Path $outDir 'demo_http_backend.xlsm'

$vbaCode = @'
Option Explicit

Private Const SERVER_URL As String = "http://127.0.0.1:8899"

Private Function HttpRequest(method As String, endpoint As String, Optional body As String = "") As String
    On Error Resume Next
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    If Err.Number <> 0 Then
        HttpRequest = "ERR:CreateObject:" & Err.Description
        Exit Function
    End If
    http.Open method, SERVER_URL & endpoint, False
    If Len(body) > 0 Then
        http.setRequestHeader "Content-Type", "text/plain"
        http.send body
    Else
        http.send
    End If
    If Err.Number <> 0 Then
        HttpRequest = "ERR:Send:" & Err.Description
        Exit Function
    End If
    If http.Status = 200 Then
        HttpRequest = http.responseText
    Else
        HttpRequest = "ERR:HTTP " & http.Status
    End If
End Function

Private Function JVal(json As String, key As String) As String
    Dim pat As String: pat = """" & key & """:"
    Dim p As Long: p = InStr(json, pat)
    If p = 0 Then JVal = "": Exit Function
    p = p + Len(pat)
    If Mid(json, p, 1) = """" Then
        p = p + 1
        Dim e As Long: e = InStr(p, json, """")
        JVal = Mid(json, p, e - p)
    Else
        Dim e2 As Long: e2 = p
        Do While e2 <= Len(json)
            Dim ch As String: ch = Mid(json, e2, 1)
            If ch >= "0" And ch <= "9" Or ch = "." Then e2 = e2 + 1 Else Exit Do
        Loop
        JVal = Mid(json, p, e2 - p)
    End If
End Function

Sub Demo_Win32Api()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.Clear

    ' Title
    ws.Cells(1, 1).Value = "Win32 API Demo via HTTP Backend"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Range("A1:E1").Merge

    ws.Cells(2, 1).Value = "No Declare statements - all Win32 API calls go through PS+C# server"
    ws.Cells(2, 1).Font.Color = RGB(128, 128, 128)
    ws.Range("A2:E2").Merge

    ' Call server
    Dim json As String: json = HttpRequest("GET", "/api/sysinfo")
    If Left(json, 4) = "ERR:" Then
        ws.Cells(4, 1).Value = "Cannot connect to backend server"
        ws.Cells(4, 1).Font.Color = RGB(255, 0, 0)
        ws.Cells(5, 1).Value = "Run 03_http_backend.bat first"
        ws.Cells(5, 1).Font.Color = RGB(128, 128, 128)
        ws.Cells(6, 1).Value = "Error: " & json
        Exit Sub
    End If

    Dim r As Long: r = 4
    ws.Cells(r, 1).Value = "Item": ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Value = "Value": ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 3).Value = "Win32 API used": ws.Cells(r, 3).Font.Bold = True
    r = r + 1

    ' Screen resolution
    Dim sw As Long, sh As Long
    sw = CLng(JVal(json, "screenWidth"))
    sh = CLng(JVal(json, "screenHeight"))
    ws.Cells(r, 1).Value = "Screen resolution"
    ws.Cells(r, 2).Value = sw & " x " & sh
    ws.Cells(r, 3).Value = "GetSystemMetrics(SM_CXSCREEN / SM_CYSCREEN)"
    r = r + 1

    ' Cursor position
    Dim cx As Long, cy As Long
    cx = CLng(JVal(json, "cursorX"))
    cy = CLng(JVal(json, "cursorY"))
    ws.Cells(r, 1).Value = "Cursor position"
    ws.Cells(r, 2).Value = cx & ", " & cy
    ws.Cells(r, 3).Value = "GetCursorPos"
    r = r + 1

    ' Foreground window
    ws.Cells(r, 1).Value = "Foreground window"
    ws.Cells(r, 2).Value = JVal(json, "foregroundWindow")
    ws.Cells(r, 3).Value = "GetForegroundWindow + GetWindowText"
    r = r + 1

    ' Uptime
    Dim uptimeMs As Double: uptimeMs = CDbl(JVal(json, "uptimeMs"))
    Dim uptimeH As Long: uptimeH = Int(uptimeMs / 3600000)
    Dim uptimeM As Long: uptimeM = Int((uptimeMs - uptimeH * 3600000) / 60000)
    ws.Cells(r, 1).Value = "System uptime"
    ws.Cells(r, 2).Value = uptimeH & "h " & uptimeM & "m"
    ws.Cells(r, 3).Value = "GetTickCount64"
    r = r + 2

    ' Color bar: cursor X mapped to rainbow
    ws.Cells(r, 1).Value = "Cursor color (run again to update)"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1

    Dim ratio As Double: ratio = cx / IIf(sw > 0, sw, 1)
    Dim hue As Long: hue = Int(ratio * 255)
    Dim i As Long
    For i = 1 To 12
        Dim h As Long: h = (hue + (i - 1) * 20) Mod 256
        Dim rr As Long, gg As Long, bb As Long
        If h < 85 Then
            rr = 255 - h * 3: gg = h * 3: bb = 0
        ElseIf h < 170 Then
            rr = 0: gg = 255 - (h - 85) * 3: bb = (h - 85) * 3
        Else
            rr = (h - 170) * 3: gg = 0: bb = 255 - (h - 170) * 3
        End If
        If rr < 0 Then rr = 0
        If rr > 255 Then rr = 255
        If gg < 0 Then gg = 0
        If gg > 255 Then gg = 255
        If bb < 0 Then bb = 0
        If bb > 255 Then bb = 255
        ws.Cells(r, i).Interior.Color = RGB(rr, gg, bb)
        ws.Cells(r, i).ColumnWidth = 4
    Next i
    r = r + 1
    ws.Cells(r, 1).Value = "^ Move mouse and re-run to see color change"
    ws.Cells(r, 1).Font.Color = RGB(128, 128, 128)

    ws.Columns("A:A").ColumnWidth = 22
    ws.Columns("B:B").ColumnWidth = 35
    ws.Columns("C:C").ColumnWidth = 50
End Sub

Sub Demo_Wiggle()
    Dim json As String: json = HttpRequest("POST", "/api/wiggle")
    If Left(json, 4) = "ERR:" Then
        MsgBox "Cannot connect to backend server." & vbCrLf & "Run 03_http_backend.bat first.", vbExclamation
    End If
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
    $ws.Name = 'Win32ApiDemo'
    $ws.Cells.Item(1, 1) = 'Run Demo_Win32Api or Demo_Wiggle (Alt+F8)'

    $wb.SaveAs($outPath, 52)
    $wb.Close($false)
    Write-Host "Generated: $outPath" -ForegroundColor Green
} finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
