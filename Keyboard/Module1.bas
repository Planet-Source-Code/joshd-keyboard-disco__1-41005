Attribute VB_Name = "Module1"
'**************************
'* Module made by A.K.San *
'* Date: 05 NOV 2002      *
'*                        *
'* Modified by Josh Duck  *
'* NOV 2002               *
'**************************

'all the API declarations needed for this program
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public File_Path As String

'some constants for later use
Public Const KEYEVENTF_KEYUP = &H2
Public Const num = &H90
Public Const scr = &H91
Public Const cap = &H14

Public Function KeyState(key As Long) As Boolean
   KeyState = (GetKeyState(key) = 1)
End Function
Public Sub keycap()
    Dim ret As Long
    ret = MapVirtualKey(cap, 0)
    keybd_event cap, ret, 0, 0
    keybd_event cap, ret, KEYEVENTF_KEYUP, 0
End Sub
Public Sub keynum()
    Dim ret As Long
    ret = MapVirtualKey(num, 0)
    keybd_event num, ret, 0, 0
    keybd_event num, ret, KEYEVENTF_KEYUP, 0
End Sub
Public Sub keyscr()
    Dim ret As Long
    ret = MapVirtualKey(scr, 0)
    keybd_event scr, ret, 0, 0
    keybd_event scr, ret, KEYEVENTF_KEYUP, 0
End Sub
