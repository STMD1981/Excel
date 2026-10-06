Attribute VB_Name = "Module2"
'windows api declarations
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr


Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long


'Constants for ShowWindow
Private Const SW_MINIMIZE = 6
Private Const SW_RESTORE = 9

Option Explicit

Dim NextWakeTime As Date
Dim dtStartTime As Date
Sub StartKeepAwake()
dtStartTime = Now()
    If NextWakeTime > 0 Then
        StopKeepAwake
    End If
    Application.StatusBar = "Keep Awake Macro Running" & Format(dtStartTime, "hh:mm:ss")
    KeepAwake
End Sub

Sub KeepAwake()
SendKeys "{F15}"

Application.StatusBar = "Keep Awake Macro Running " & _
        Format(dtStartTime, "hh:mm:ss") & _
        " | Elapsed since start: " & Format(Now - dtStartTime, "hh:mm:ss")

    
    NextWakeTime = Now() + TimeValue("00:00:30")
    Application.OnTime NextWakeTime, "'" & ThisWorkbook.Name & "'!KeepAwake"
    

End Sub

Sub StopKeepAwake()
    If NextWakeTime > 0 Then
         Application.OnTime NextWakeTime, "'" & ThisWorkbook.Name & "'!KeepAwake", , False
         NextWakeTime = 0
    End If
   
    Application.StatusBar = False
End Sub





Sub LoopWithVisualSectionMinimized()
Dim i As Long
Dim hwnd As LongPtr
Dim xlappname As String


'get excel windows handle

xlappname = Application.Caption
hwnd = FindWindowA("XLMAIN", xlappname)

'minimize excel window without stopping macro
ShowWindow hwnd, SW_MINIMIZE

For i = 1 To 10000
    SelectAndWait "A1"
    SelectAndWait "D1"
    SelectAndWait "D4"
    SelectAndWait "A4"
    
    If TeamsToastVisible() Then
        MsgBox "Teams Notification"
    End If
Next i

'restore excel when done
ShowWindow hwnd, SW_RESTORE

End Sub

Private Sub SelectAndWait(cellAddress As String)
ThisWorkbook.Sheets(1).Range(cellAddress).Select
Sleep 5000

End Sub

Sub working()
Attribute working.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
Dim i As Long
i = 0
Application.StatusBar = i
For i = 1 To 15000
    Application.StatusBar = i
    
    Range("A1").Value = i
    Application.Wait Now + TimeValue("0:00:05")
    SendKeys "~"
    
    Range("D1").Value = i
    Application.Wait Now + TimeValue("0:00:05")
    
    
    SendKeys "~"
    
    Range("D4").Value = i
    Application.Wait Now + TimeValue("0:00:05")
    SendKeys "~"
    
    Range("A4").Value = i
    Application.Wait Now + TimeValue("0:00:05")
    
    
    
    SendKeys "~"
Next i

End Sub

Sub working2()
'
' Macro2 Macro
Dim i As Long
i = 0
Application.StatusBar = i
For i = 1 To 15000
    Application.StatusBar = i
    
    Range("A1").Value = i
    Application.Wait Now + TimeValue("0:00:05")
    'SendKeys "~"
    
    Range("D1").Value = i
    Application.Wait Now + TimeValue("0:00:05")
    'SendKeys "~"
    
    Range("D4").Value = i
    Application.Wait Now + TimeValue("0:00:05")
    'SendKeys "~"
    
    Range("A4").Value = i
    Application.Wait Now + TimeValue("0:00:05")
    'SendKeys "~"
Next i

End Sub

Function HasTeamsNotification() As Boolean
Dim hwnd As LongPtr
Dim windowTitle As String * 255
Dim titleLen As Long

    hwnd = FindWindowA("ApplicationFrameWindow", vbNullString)
    
    If hwnd <> 0 Then
        'Get window title
        titleLen = GetWindowText(hwnd, windowTitle, 255)
        windowTitle = Left(windowTitle, titleLen)
        
        'check for number in parens
        If InStr(windowTitle, "(") > 0 And InStr(windowTitle, ")") > 0 Then
            HasTeamsNotification = True
        Else
            HasTeamsNotification = False
        End If
    Else
        'Teams not running
        HasTeamsNotification = False
    End If
    

End Function

Function TeamsToastVisible() As Boolean

Dim hwnd As LongPtr
hwnd = FindWindowEx(0, 0, "Windows.UI.Core.CoreWindow", "Notification Center")
TeamsToastVisible = (hwnd <> 0)

End Function

