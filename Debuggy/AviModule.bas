Attribute VB_Name = "AviModule"
Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" _
(ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function mciGetErrorString Lib "winmm" Alias _
"mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, _
ByVal uLength As Long) As Long
Public MMFile As String
Private Const WS_CHILD = &H40000000

Public Sub GetSize(ByRef Width As Long, ByRef Height As Long)
Dim OutString As String
Dim SpStrings() As String
OutString = Space(128)
ret = mciSendString("Where " & MMFile & " destination", OutString, 128, 0&)
SpStrings = Split(OutString, " ")
Width = CLng(SpStrings(2))
Height = CLng(SpStrings(3))
End Sub

Public Function OpenMedia(Window As Variant) As Long
Dim CmdStr As String
CmdStr = "Open " & MMFile & " type MPEGVideo alias " & MMFile & " parent " _
& CStr(Window.hWnd) & " style " & CStr(WS_CHILD)
OpenMedia = mciSendString(CmdStr, 0&, 0&, 0&)
End Function
Public Sub SetInWindow(Window As Variant)
Dim CmdStr As String
CmdStr = "put " & MMFile & " window at 0 0 " & CStr(Window.ScaleWidth / _
15) & " " & CStr(Window.ScaleHeight / _
15)
mciSendString CmdStr, 0&, 0&, 0&
End Sub
Public Sub PlayMedia()
mciSendString "Play " & MMFile, 0&, 0&, 0&
End Sub

Public Sub PauseMedia()
mciSendString "Pause " & MMFile, 0&, 0&, 0&
End Sub

Public Sub StopMedia()
mciSendString "Stop " & MMFile, 0&, 0&, 0&
End Sub

Public Sub CloseMedia()
'mciSendString "Close all", 0&, 0&, 0&
mciSendString "Close " & MMFile, 0&, 0&, 0&
End Sub
