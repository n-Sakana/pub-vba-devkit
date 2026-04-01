Attribute VB_Name = "PathBreakDemo"
Option Explicit

Public Sub RunHardcodedPathDemo()
    Dim ws As Worksheet
    Dim hardCodedPath As String
    Dim workbookRelativePath As String
    Dim actualDataPath As String
    
    Set ws = ThisWorkbook.Worksheets("Demo")
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("Check", "Result")
    ws.Range("A2").Value = "Workbook path"
    ws.Range("B2").Value = ThisWorkbook.Path
    
    hardCodedPath = "C:\SharePoint\TeamSite\Shared Documents\案件データ"
    workbookRelativePath = ThisWorkbook.Path & "\案件データ"
    actualDataPath = Environ$("OneDriveCommercial")
    If Len(actualDataPath) = 0 Then actualDataPath = Environ$("OneDrive")
    actualDataPath = actualDataPath & "\_vba_devkit_samples\SharePointDemo\Shared Documents\案件データ"
    
    ws.Range("A4").Value = "Hard-coded path"
    ws.Range("B4").Value = DescribeFolder(hardCodedPath)
    ws.Range("A5").Value = "ThisWorkbook.Path + relative folder"
    ws.Range("B5").Value = DescribeFolder(workbookRelativePath)
    ws.Range("A6").Value = "Actual synced data path"
    ws.Range("B6").Value = DescribeFolder(actualDataPath)
    
    ws.Range("A8").Value = "Expected"
    ws.Range("B8").Value = "Only the actual synced path exists. Old path assumptions fail."
    
    FormatStatus ws.Range("B4")
    FormatStatus ws.Range("B5")
    FormatStatus ws.Range("B6")
    ws.Columns("A:B").AutoFit
End Sub

Private Function DescribeFolder(ByVal targetPath As String) As String
    Dim fileName As String
    If Len(Dir$(targetPath, vbDirectory)) = 0 Then
        DescribeFolder = "Missing -> " & targetPath
        Exit Function
    End If
    fileName = Dir$(targetPath & "\*.*")
    If Len(fileName) = 0 Then
        DescribeFolder = "Exists but empty -> " & targetPath
    Else
        DescribeFolder = "Exists / first file: " & fileName & " -> " & targetPath
    End If
End Function

Private Sub FormatStatus(ByVal cell As Range)
    If InStr(1, cell.Value2, "Missing", vbTextCompare) > 0 Then
        cell.Interior.Color = RGB(255, 220, 220)
    Else
        cell.Interior.Color = RGB(220, 255, 220)
    End If
End Sub
