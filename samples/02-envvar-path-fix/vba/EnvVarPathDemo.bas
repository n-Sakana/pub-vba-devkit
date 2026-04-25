Attribute VB_Name = "EnvVarPathDemo"
Option Explicit

Public Sub RunEnvVarPathDemo()
    Dim ws As Worksheet
    Dim syncRoot As String
    Dim targetPath As String
    Dim item As String
    Dim rowIndex As Long
    
    Set ws = ThisWorkbook.Worksheets("Demo")
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("Item", "Value")
    
    syncRoot = GetPreferredSyncRoot()
    targetPath = syncRoot & "\_vba_devkit_samples\SharePointDemo\Shared Documents\案件データ"
    
    ws.Range("A2").Value = "Resolved sync root"
    ws.Range("B2").Value = syncRoot
    ws.Range("A3").Value = "Resolved target path"
    ws.Range("B3").Value = targetPath
    ws.Range("A5:B5").Value = Array("Files", "Found")
    rowIndex = 6
    item = Dir$(targetPath & "\*.*")
    Do While Len(item) > 0
        ws.Cells(rowIndex, 1).Value = item
        ws.Cells(rowIndex, 2).Value = "OK"
        ws.Cells(rowIndex, 2).Interior.Color = RGB(220, 255, 220)
        rowIndex = rowIndex + 1
        item = Dir$()
    Loop
    
    If rowIndex = 6 Then
        ws.Cells(6, 1).Value = "No files found"
        ws.Cells(6, 2).Value = "Check sample data"
        ws.Cells(6, 2).Interior.Color = RGB(255, 220, 220)
    End If
    
    ws.Columns("A:B").AutoFit
End Sub

Private Function GetPreferredSyncRoot() As String
    GetPreferredSyncRoot = Environ$("OneDriveCommercial")
    If Len(GetPreferredSyncRoot) = 0 Then
        GetPreferredSyncRoot = Environ$("OneDrive")
    End If
    If Len(GetPreferredSyncRoot) = 0 Then
        Err.Raise vbObjectError + 1000, "EnvVarPathDemo", "OneDriveCommercial / OneDrive が取得できません。"
    End If
End Function
