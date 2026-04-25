Attribute VB_Name = "IpcHttpClientDemo"
Option Explicit

Public Sub RunVisualWin32Demo()
    Dim http As Object
    Dim responseText As String
    Dim parts() As String
    Dim ws As Worksheet
    Dim panel As Shape
    Dim r As Long, g As Long, b As Long
    
    Set ws = ThisWorkbook.Worksheets("Demo")
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("Field", "Value")
    
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", "http://127.0.0.1:8765/api/foreground", False
    http.send
    
    If http.Status <> 200 Then
        ws.Range("A3").Value = "HTTP Status"
        ws.Range("B3").Value = http.Status
        ws.Range("B3").Interior.Color = RGB(255, 220, 220)
        Exit Sub
    End If
    
    responseText = http.responseText
    parts = Split(responseText, "|")
    If UBound(parts) < 5 Then
        ws.Range("A3").Value = "Response"
        ws.Range("B3").Value = responseText
        ws.Range("B3").Interior.Color = RGB(255, 220, 220)
        Exit Sub
    End If
    
    ws.Range("A3").Value = "Foreground title"
    ws.Range("B3").Value = parts(0)
    ws.Range("A4").Value = "Window class"
    ws.Range("B4").Value = parts(1)
    ws.Range("A5").Value = "Width"
    ws.Range("B5").Value = CLng(parts(2))
    ws.Range("A6").Value = "Height"
    ws.Range("B6").Value = CLng(parts(3))
    ws.Range("A7").Value = "Left"
    ws.Range("B7").Value = CLng(parts(4))
    ws.Range("A8").Value = "Top"
    ws.Range("B8").Value = CLng(parts(5))
    
    r = (CLng(parts(2)) * 3) Mod 256
    g = (CLng(parts(3)) * 5) Mod 256
    b = ((CLng(parts(4)) + CLng(parts(5))) * 7) Mod 256
    
    On Error Resume Next
    ws.Shapes("Win32Panel").Delete
    On Error GoTo 0
    
    Set panel = ws.Shapes.AddShape(msoShapeRoundedRectangle, 220, 40, 260, 120)
    panel.Name = "Win32Panel"
    panel.Fill.ForeColor.RGB = RGB(r, g, b)
    panel.Line.ForeColor.RGB = RGB(60, 60, 60)
    panel.TextFrame2.TextRange.Text = "Foreground window" & vbCrLf & parts(0) & vbCrLf & parts(1) & vbCrLf & parts(2) & " x " & parts(3)
    panel.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    panel.TextFrame2.TextRange.Font.Size = 12
    
    ws.Tab.Color = RGB(r, g, b)
    ws.Columns("A:B").AutoFit
End Sub
