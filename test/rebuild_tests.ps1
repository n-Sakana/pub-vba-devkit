$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false

try {
    $testDir = $PSScriptRoot

    # === test_sample.xlsm (EDR: Declare + call sites) ===
    $wb = $excel.Workbooks.Add()
    $mod = $wb.VBProject.VBComponents.Add(1)
    $mod.Name = 'TestModule'
    $mod.CodeModule.AddFromString(@'
Option Explicit
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) As Long
Public Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Sub TestAPI()
    Dim t As Long
    t = GetTickCount()
    Sleep 100
    MsgBox "Elapsed: " & (GetTickCount() - t) & " ms"
End Sub
'@)
    $cls = $wb.VBProject.VBComponents.Add(2)
    $cls.Name = 'TestClass'
    $cls.CodeModule.AddFromString("Option Explicit`r`nPrivate m_name As String")
    $wb.SaveAs("$testDir\test_sample.xlsm", 52)
    $wb.Close($false)
    Write-Host 'test_sample.xlsm' -ForegroundColor Green

    # === test_large.xlsm (multiple modules with many Declare) ===
    $wb = $excel.Workbooks.Add()
    $m1 = $wb.VBProject.VBComponents.Add(1)
    $m1.Name = 'WindowUtils'
    $m1.CodeModule.AddFromString(@'
Option Explicit
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Public Sub Test(): End Sub
'@)
    $m2 = $wb.VBProject.VBComponents.Add(1)
    $m2.Name = 'TimerUtils'
    $m2.CodeModule.AddFromString(@'
Option Explicit
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) As Long
Public Sub Test(): End Sub
'@)
    $m3 = $wb.VBProject.VBComponents.Add(1)
    $m3.Name = 'SystemInfo'
    $m3.CodeModule.AddFromString(@'
Option Explicit
Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Sub Test(): End Sub
'@)
    $m4 = $wb.VBProject.VBComponents.Add(1)
    $m4.Name = 'AppController'
    $m4.CodeModule.AddFromString("Option Explicit`r`nPublic Sub Test(): End Sub")
    $wb.SaveAs("$testDir\test_large.xlsm", 52)
    $wb.Close($false)
    Write-Host 'test_large.xlsm' -ForegroundColor Green

    # === test_envbiz.xlsm (path + compat + biz patterns, no Declare) ===
    $wb = $excel.Workbooks.Add()
    $mod = $wb.VBProject.VBComponents.Add(1)
    $mod.Name = 'EnvDependentModule'
    $mod.CodeModule.AddFromString(@'
Option Explicit

Public Sub OpenReport()
    Dim path As String
    path = "C:\Reports\monthly_report.xlsx"
    Workbooks.Open "\\server\shared\templates\base.xlsm"

    Dim userPath As String
    userPath = "C:\Users\example\Desktop\output.csv"

    Dim appDataPath As String
    appDataPath = "C:\Users\example\AppData\Local\MyApp\config.ini"

    Application.ActivePrinter = "PrinterName on Ne00:"
    ActiveWorkbook.PrintOut
End Sub

Public Sub ConnectDB()
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Data\master.accdb"
    cn.Open

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM customers", cn

    cn.Close
End Sub
'@)
    $mod2 = $wb.VBProject.VBComponents.Add(1)
    $mod2.Name = 'BizIntegrationModule'
    $mod2.CodeModule.AddFromString(@'
Option Explicit

Public Sub SendReport()
    Dim olApp As Object
    Set olApp = CreateObject("Outlook.Application")
    Dim mail As Object
    Set mail = olApp.CreateItem(0)
    mail.To = "manager@company.com"
    mail.Subject = "Monthly Report"
    mail.Body = "Please see attached."
    mail.Display
    Set mail = Nothing
    Set olApp = Nothing
End Sub

Public Sub GenerateWord()
    Dim wdApp As Object
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Dim doc As Object
    Set doc = wdApp.Documents.Add
    doc.Content.Text = "Generated Report"
    doc.SaveAs2 ThisWorkbook.Path & "\report.docx"
    doc.Close
    wdApp.Quit
End Sub

Public Sub ExportPDF()
    ActiveSheet.ExportAsFixedFormat xlTypePDF, ThisWorkbook.Path & "\output.pdf"
End Sub

Public Sub RunExternal()
    Shell "notepad.exe C:\temp\log.txt", vbNormalFocus
End Sub

Public Sub AccessDB()
    Dim db As Object
    Set db = CreateObject("DAO.DBEngine.36")
End Sub
'@)
    $mod3 = $wb.VBProject.VBComponents.Add(1)
    $mod3.Name = 'CompatModule'
    $mod3.CodeModule.AddFromString(@'
Option Explicit

DefLng A-Z

Public Sub LegacyCode()
    GoSub DoWork
    Exit Sub
DoWork:
    Dim i As Long
    i = 0
    While i < 10
        i = i + 1
    Wend
    Return
End Sub

Public Sub TestIE()
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Navigate "http://localhost:8080/api"
    Set ie = Nothing
End Sub

Public Sub TestDDE()
    On Error Resume Next
    Dim ch As Long
    ch = DDEInitiate("Excel", "Sheet1")
    On Error GoTo 0
End Sub
'@)
    $wb.SaveAs("$testDir\test_envbiz.xlsm", 52)
    $wb.Close($false)
    Write-Host 'test_envbiz.xlsm' -ForegroundColor Green

    # === test_diff_old.xlsm ===
    $wb = $excel.Workbooks.Add()
    $mod = $wb.VBProject.VBComponents.Add(1)
    $mod.Name = 'DataProcessor'
    $mod.CodeModule.AddFromString(@'
Option Explicit

Public Sub ProcessData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        ws.Cells(i, 3).Value = ws.Cells(i, 1).Value * ws.Cells(i, 2).Value
    Next i

    MsgBox "Done: " & (lastRow - 1) & " rows processed"
End Sub

Public Sub ExportCSV()
    Dim path As String
    path = "C:\Reports\output.csv"

    Dim f As Long: f = FreeFile
    Open path For Output As #f

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 1 To lastRow
        Print #f, ws.Cells(i, 1).Value & "," & ws.Cells(i, 2).Value & "," & ws.Cells(i, 3).Value
    Next i
    Close #f
End Sub
'@)
    $cls = $wb.VBProject.VBComponents.Add(2)
    $cls.Name = 'Logger'
    $cls.CodeModule.AddFromString(@'
Option Explicit
Private m_log As String

Public Sub Log(msg As String)
    m_log = m_log & Format(Now, "hh:nn:ss") & " " & msg & vbCrLf
    Debug.Print msg
End Sub

Public Function GetLog() As String
    GetLog = m_log
End Function
'@)
    $wb.SaveAs("$testDir\test_diff_old.xlsm", 52)
    $wb.Close($false)
    Write-Host 'test_diff_old.xlsm' -ForegroundColor Green

    # === test_diff_new.xlsm ===
    $wb = $excel.Workbooks.Add()
    $mod = $wb.VBProject.VBComponents.Add(1)
    $mod.Name = 'DataProcessor'
    $mod.CodeModule.AddFromString(@'
Option Explicit

Public Sub ProcessData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No data to process"
        Exit Sub
    End If

    Dim i As Long
    For i = 2 To lastRow
        ws.Cells(i, 3).Value = ws.Cells(i, 1).Value * ws.Cells(i, 2).Value
        ws.Cells(i, 4).Value = Format(Now, "yyyy/mm/dd")
    Next i

    MsgBox "Done: " & (lastRow - 1) & " rows processed"
End Sub

Public Sub ExportCSV()
    Dim path As String
    path = Environ$("TEMP") & "\output.csv"

    Dim f As Long: f = FreeFile
    Open path For Output As #f

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 1 To lastRow
        Print #f, ws.Cells(i, 1).Value & "," & ws.Cells(i, 2).Value & "," & ws.Cells(i, 3).Value & "," & ws.Cells(i, 4).Value
    Next i
    Close #f
End Sub

Public Sub ValidateData()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim errCount As Long
    Dim i As Long
    For i = 2 To lastRow
        If IsEmpty(ws.Cells(i, 1)) Or IsEmpty(ws.Cells(i, 2)) Then errCount = errCount + 1
    Next i
    If errCount > 0 Then MsgBox errCount & " empty cells found"
End Sub
'@)
    $cls = $wb.VBProject.VBComponents.Add(2)
    $cls.Name = 'Logger'
    $cls.CodeModule.AddFromString(@'
Option Explicit
Private m_log As String
Private m_level As String

Public Sub Init(Optional level As String = "INFO")
    m_level = level
    m_log = ""
End Sub

Public Sub Log(msg As String)
    m_log = m_log & Format(Now, "yyyy/mm/dd hh:nn:ss") & " [" & m_level & "] " & msg & vbCrLf
    Debug.Print "[" & m_level & "]" & msg
End Sub

Public Function GetLog() As String
    GetLog = m_log
End Function

Public Sub Clear()
    m_log = ""
End Sub
'@)
    $val = $wb.VBProject.VBComponents.Add(1)
    $val.Name = 'Validator'
    $val.CodeModule.AddFromString(@'
Option Explicit

Public Function IsValid(val As Variant) As Boolean
    IsValid = Not IsEmpty(val) And Not IsNull(val)
End Function
'@)
    $wb.SaveAs("$testDir\test_diff_new.xlsm", 52)
    $wb.Close($false)
    Write-Host 'test_diff_new.xlsm' -ForegroundColor Green

    # === test_protected.xlsm (password-protected VBA project) ===
    # VBA project password cannot be set programmatically (Protection is read-only).
    # This generates the file without protection; to add VBA password:
    #   1. Open in Excel → Alt+F11 → Tools → VBAProject Properties → Protection
    #   2. Check "Lock project for viewing", set password, save
    $wb = $excel.Workbooks.Add()
    $mod = $wb.VBProject.VBComponents.Add(1)
    $mod.Name = 'TestModule'
    $mod.CodeModule.AddFromString(@'
Option Explicit

Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) As Long
Public Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Sub TestAPI()
    Dim t As Long
    t = GetTickCount()
    Sleep 100
    MsgBox "Elapsed: " & (GetTickCount() - t) & " ms"
End Sub

Public Sub TestNoAPI()
    MsgBox "Hello, no API here"
End Sub

Public Sub TestUserName()
    Dim buf As String: buf = Space(256)
    Dim sz As Long: sz = 256
    GetUserName buf, sz
    MsgBox "User: " & Left(buf, sz - 1)
End Sub
'@)
    $cls = $wb.VBProject.VBComponents.Add(2)
    $cls.Name = 'TestClass'
    $cls.CodeModule.AddFromString(@'
Option Explicit
Private m_name As String
Public Property Let Name(val As String): m_name = val: End Property
Public Property Get Name() As String: Name = m_name: End Property
'@)
    $wb.SaveAs("$testDir\test_protected.xlsm", 52)
    $wb.Close($false)
    Write-Host 'test_protected.xlsm (unprotected -- set VBA password manually)' -ForegroundColor Yellow

    Write-Host "`nAll test files rebuilt." -ForegroundColor Green
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
