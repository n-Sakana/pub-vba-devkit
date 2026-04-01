$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false

try {
    $wb = $excel.Workbooks.Add()
    $mod = $wb.VBProject.VBComponents.Add(1)
    $mod.Name = 'TestReplacements'
    $mod.CodeModule.AddFromString(@'
Option Explicit

' ============================================================
' Win32 API Migration Replacement Tests
' Each test: API version first, then VBA alternative, compare results
' ============================================================

Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function Sleep Lib "kernel32" (ByVal ms As Long) As Long
Private Declare PtrSafe Function GetUserNameA Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare PtrSafe Function GetComputerNameA Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare PtrSafe Function GetTempPathA Lib "kernel32" Alias "GetTempPathA" (ByVal nBufLen As Long, ByVal lpBuffer As String) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private m_results As String

Private Sub Log(test As String, apiVal As String, altVal As String, ok As Boolean)
    Dim status As String: status = IIf(ok, "PASS", "FAIL")
    Dim line As String
    line = status & " | " & test & " | API=" & apiVal & " | Alt=" & altVal
    m_results = m_results & line & vbCrLf
    Debug.Print line
End Sub

Public Function RunAllTests() As String
    m_results = "# Win32 API Replacement Test Results" & vbCrLf
    m_results = m_results & "# Date: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & vbCrLf

    TestGetTickCount
    TestSleep
    TestGetUserName
    TestGetComputerName
    TestGetTempPath
    TestGetSystemMetrics
    TestShowWindow

    RunAllTests = m_results
End Function

' --- GetTickCount -> Timer ---
Private Sub TestGetTickCount()
    ' API version
    Dim t1 As Long: t1 = GetTickCount()
    Dim dummy As Double: dummy = 0
    Dim i As Long
    For i = 1 To 100000: dummy = dummy + Sqr(i): Next i
    Dim apiElapsed As Long: apiElapsed = GetTickCount() - t1

    ' VBA alternative: Timer (Single)
    Dim t2 As Single: t2 = Timer
    dummy = 0
    For i = 1 To 100000: dummy = dummy + Sqr(i): Next i
    Dim altElapsed As Single: altElapsed = Timer - t2
    If altElapsed < 0 Then altElapsed = altElapsed + 86400  ' midnight fix
    Dim altMs As Long: altMs = CLng(altElapsed * 1000)

    ' Both should be similar (within 50% tolerance)
    Dim ratio As Double
    If apiElapsed > 0 Then ratio = Abs(altMs - apiElapsed) / apiElapsed Else ratio = 0
    Log "GetTickCount->Timer", CStr(apiElapsed) & "ms", CStr(altMs) & "ms", ratio < 0.5
End Sub

' --- Sleep -> Application.Wait ---
Private Sub TestSleep()
    ' API version
    Dim t1 As Single: t1 = Timer
    Sleep 500  ' 500ms
    Dim apiElapsed As Single: apiElapsed = (Timer - t1) * 1000

    ' VBA alternative: Application.Wait (1sec min resolution)
    Dim t2 As Single: t2 = Timer
    Application.Wait Now + TimeSerial(0, 0, 1)  ' 1 second
    Dim altElapsed As Single: altElapsed = (Timer - t2) * 1000

    ' API should be ~500ms, Alt should be ~1000ms
    Dim apiOk As Boolean: apiOk = (apiElapsed > 400 And apiElapsed < 700)
    Dim altOk As Boolean: altOk = (altElapsed > 900 And altElapsed < 1200)
    Log "Sleep->Application.Wait", Format(apiElapsed, "0") & "ms", Format(altElapsed, "0") & "ms (1sec min)", apiOk And altOk

    ' DoEvents loop alternative
    Dim t3 As Single: t3 = Timer
    Dim endTime As Single: endTime = Timer + 0.5  ' 500ms
    Do While Timer < endTime: DoEvents: Loop
    Dim loopElapsed As Single: loopElapsed = (Timer - t3) * 1000
    Dim loopOk As Boolean: loopOk = (loopElapsed > 400 And loopElapsed < 700)
    Log "Sleep->DoEvents loop", Format(apiElapsed, "0") & "ms", Format(loopElapsed, "0") & "ms", loopOk
End Sub

' --- GetUserName -> Environ$("USERNAME") ---
Private Sub TestGetUserName()
    ' API version
    Dim buf As String: buf = Space(256)
    Dim sz As Long: sz = 256
    GetUserNameA buf, sz
    Dim apiUser As String: apiUser = Left$(buf, sz - 1)

    ' VBA alternative
    Dim altUser As String: altUser = Environ$("USERNAME")

    Log "GetUserName->Environ$(USERNAME)", apiUser, altUser, (apiUser = altUser)

    ' Also test Application.UserName (may differ)
    Dim appUser As String: appUser = Application.UserName
    Dim same As Boolean: same = (apiUser = appUser)
    Log "GetUserName->Application.UserName", apiUser, appUser, same
    If Not same Then
        m_results = m_results & "  NOTE: Application.UserName is Office display name, not Windows login" & vbCrLf
    End If
End Sub

' --- GetComputerName -> Environ$("COMPUTERNAME") ---
Private Sub TestGetComputerName()
    ' API version
    Dim buf As String: buf = Space(256)
    Dim sz As Long: sz = 256
    GetComputerNameA buf, sz
    Dim apiComp As String: apiComp = Left$(buf, sz - 1)

    ' VBA alternative
    Dim altComp As String: altComp = Environ$("COMPUTERNAME")

    Log "GetComputerName->Environ$(COMPUTERNAME)", apiComp, altComp, (apiComp = altComp)
End Sub

' --- GetTempPath -> Environ$("TEMP") ---
Private Sub TestGetTempPath()
    ' API version
    Dim buf As String: buf = Space(260)
    Dim ret As Long: ret = GetTempPathA(260, buf)
    Dim apiPath As String: apiPath = Left$(buf, ret)

    ' VBA alternative (add trailing backslash)
    Dim altPath As String: altPath = Environ$("TEMP") & "\"

    Log "GetTempPath->Environ$(TEMP)", apiPath, altPath, (LCase(apiPath) = LCase(altPath))
End Sub

' --- GetSystemMetrics -> Application.UsableWidth/Height ---
Private Sub TestGetSystemMetrics()
    ' API version (screen size in pixels)
    Dim apiW As Long: apiW = GetSystemMetrics(0)  ' SM_CXSCREEN
    Dim apiH As Long: apiH = GetSystemMetrics(1)  ' SM_CYSCREEN

    ' VBA alternative (workspace in points - different unit and scope)
    Dim altW As Double: altW = Application.UsableWidth
    Dim altH As Double: altH = Application.UsableHeight

    ' These won't match (pixels vs points, screen vs workspace)
    ' Just verify both return reasonable values > 0
    Dim apiOk As Boolean: apiOk = (apiW > 100 And apiH > 100)
    Dim altOk As Boolean: altOk = (altW > 100 And altH > 100)
    Log "GetSystemMetrics->UsableWidth/Height", apiW & "x" & apiH & "px", Format(altW, "0") & "x" & Format(altH, "0") & "pt", apiOk And altOk
    m_results = m_results & "  NOTE: Different units (pixels vs points) and scope (screen vs workspace)" & vbCrLf
End Sub

' --- ShowWindow -> Application.WindowState ---
Private Sub TestShowWindow()
    ' Just test the VBA WindowState property works
    Dim origState As Long: origState = Application.WindowState

    Application.WindowState = xlMinimized
    Dim isMin As Boolean: isMin = (Application.WindowState = xlMinimized)

    Application.WindowState = xlMaximized
    Dim isMax As Boolean: isMax = (Application.WindowState = xlMaximized)

    Application.WindowState = xlNormal
    Dim isNorm As Boolean: isNorm = (Application.WindowState = xlNormal)

    ' Restore original
    Application.WindowState = origState

    Log "ShowWindow->WindowState", "minimize/maximize/normal", "xlMinimized/xlMaximized/xlNormal", isMin And isMax And isNorm
End Sub
'@)

    $wb.SaveAs("$PSScriptRoot\test_replacements.xlsm", 52)

    # Run the tests
    Write-Host "Running replacement tests..."
    $result = $excel.Run("TestReplacements.RunAllTests")
    $wb.Close($false)

    Write-Host ""
    Write-Host $result

    # Save results
    [IO.File]::WriteAllText("$PSScriptRoot\test_replacements_result.txt", $result, [System.Text.Encoding]::UTF8)
    Write-Host "Results saved to: test_replacements_result.txt"
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
