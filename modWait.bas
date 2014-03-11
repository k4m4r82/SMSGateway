Attribute VB_Name = "modWait"
Option Explicit

Private mCancel As Boolean

Type MSG
   hwnd As Long
   message As Long
   wParam As Long
   lParam As Long
   time As Long
   ptX As Long
   ptY As Long
End Type

Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private Sub TimerProc()
    mCancel = True
End Sub

Public Sub Wait(mSecs As Long)
    Dim MyMsg As MSG
    Dim TimerID As Long
    
    TimerID = SetTimer(0, 0, mSecs, AddressOf TimerProc)
    mCancel = False

    Do While Not mCancel
        GetMessage MyMsg, 0, 0, 0
        TranslateMessage MyMsg
        DispatchMessage MyMsg
    Loop

    KillTimer 0, TimerID
End Sub


