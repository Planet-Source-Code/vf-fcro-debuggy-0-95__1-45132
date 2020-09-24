Attribute VB_Name = "WndBreaks"
Option Explicit
Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public TWProc1 As Long
Public TWProc2 As Long

Dim FObjS As String


Public CheckUNI As New Collection

Public XPARAMS() As Long
Public XTEXTS() As String
Public XFNAME As String
Public XFNAME2 As String
'Public XFNAME3 As String
'Public XFNAME4 As String


Public NOTA1 As Long 'Api notifycation Address
Public NOTA2 As Long ' To Address
Public IsENOT As Byte 'IsNotify Enabled



Public BRKW As New Collection
Public BRKWMCMD As New Collection 'On WM_COMMAND
Public Function AddBreakWND(ByRef COL As Collection, ByVal hWnd As Long, ByVal BrkMsgValue As Long, ByVal AfterOrBefore As Long) As Byte
On Error Resume Next
Dim C As Collection
Dim isEx As Byte
Set C = GetRemoveCol(COL, hWnd, isEx)
If isEx = 0 Then Set C = New Collection

Dim Rk(2) As Long
Rk(0) = hWnd
Rk(1) = BrkMsgValue
Rk(2) = AfterOrBefore
C.Add Rk, "X" & BrkMsgValue
If Err <> 0 Then
On Error GoTo 0
Else
AddBreakWND = 1
End If
COL.Add C, "X" & hWnd
End Function
Public Function GetRemoveCol(ByRef COL As Collection, ByVal hWnd As Long, ByRef IsExist As Byte) As Collection
On Error GoTo Dalje
Set GetRemoveCol = COL.Item("X" & hWnd)
COL.Remove "X" & hWnd
IsExist = 1
Exit Function
Dalje:
On Error GoTo 0
End Function

Public Sub RemoveBreakWND(ByRef COL As Collection, ByVal hWnd As Long, ByVal BrkMsgValue As Long, ByVal AfterOrBefore As Long)
On Error GoTo Dalje
Dim C As Collection
Set C = COL.Item("X" & hWnd)
C.Remove "X" & BrkMsgValue
Exit Sub
Dalje:
On Error GoTo 0
End Sub
Public Function GetBreakWND(ByRef COL As Collection, ByVal hWnd As Long, ByVal BrkMsgValue As Long, ByVal AfterOrBefore As Long, ByRef IsValidWNDBP As Byte) As Long()
On Error GoTo Dalje:
Dim C As Collection
Set C = COL.Item("X" & hWnd)
GetBreakWND = C.Item("X" & BrkMsgValue)
IsValidWNDBP = 1
Exit Function
Dalje:
On Error GoTo 0
End Function

Public Sub RemoveEntireWND(ByRef COL As Collection, ByVal hWnd As Long)
On Error GoTo Dalje
COL.Remove "X" & hWnd
Exit Sub
Dalje:
On Error GoTo 0
End Sub


Public Function TextProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case uMsg
Case &H7B
Exit Function
End Select
TextProc = CallWindowProc(TWProc1, hWnd, uMsg, wParam, lParam)
End Function
Public Function TextProc2(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case uMsg
Case &H7B
Exit Function
End Select
TextProc2 = CallWindowProc(TWProc2, hWnd, uMsg, wParam, lParam)
End Function



Public Sub TranslateSpy(ByRef Address As Long, ByRef CONS As CONTEXT, ByRef ThreadId As Long, ByRef PROCRF As Long)
'On Error GoTo Dalje

Dim WWver As Byte
Dim Pstr As String
Dim ret As Long
Dim Aad As Long
Dim PRM() As Long
ret = ReadProcessMemory(ProcessHandle, ByVal CONS.Esp, Aad, 4, ByVal 0&)

If ret = 0 Then Exit Sub


If IsENOT <> 0 Then: If Not (Aad >= NOTA1 And Aad <= NOTA2) Then Exit Sub


If Aad >= DEBUGGYFA And Aad <= DEBUGGYLA Then Exit Sub


Dim FncCall As String
FncCall = GetFromExportsSearch(FindInModules(Address), Address)

Select Case FncCall


'Case "VirtualProtectEx"
'Remove protection
'Dim NPP As Long
'NPP = 4
'WriteProcessMemory ProcessHandle, ByVal AddBy8(CTX.Esp, &H10&), NPP, 4, ByVal 0&

'Case "VirtualAllocEx"
'Dim NPP2 As Long
'NPP2 = 4
'WriteProcessMemory ProcessHandle, ByVal AddBy8(CTX.Esp, &H14&), NPP, 4, ByVal 0&


'Case "SendMessageA"
'GoTo SMW

'Case "SendMessageW"
'SMW:
'ReDim Prm(3)
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(ESPptr, 4), Prm(0), 16, ByVal 0&)
'AddLine "API " & FncCall, Form30.rt1, &HA2&
'AddLine "Call-Retun Address At:" & Hex(Aad) & " ,To hWnd:" & Hex(Prm(0)) & " ,Msg:" & Hex(Prm(1)) _
'& " ,wParam:" & Hex(Prm(2)) & " ,lParam:" & Hex(Prm(3)), Form30.rt1
'AddLine "", Form30.rt1

'Case "SendDlgItemMessageA"
'GoTo SDIM

'Case "SendDlgItemMessageW"
'SDIM:
'ReDim Prm(4)
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(ESPptr, 4), Prm(0), 20, ByVal 0&)
'AddLine "API " & FncCall, Form30.rt1, &HA2&
'AddLine "Call-Retun Address At:" & Hex(Aad) & " ,To hDlg:" & Hex(Prm(0)) & " ,To Dlg Item:" & Hex(Prm(1)) _
'& " ,Msg:" & Hex(Prm(2)) & " ,wParam:" & Hex(Prm(3)) & " ,lParam:" & Hex(Prm(4)), Form30.rt1
'AddLine "", Form30.rt1












'USER32

'***************
'Za Timer!
'Case "SetTimer"
'If SPConfig(0) = 0 Then Exit Sub
'TimerRef FncCall, PRM, Aad, CONS.Esp, ThreadId, PROCRF
'****************


'Case "KillTimer"
'TimerKl FncCall, Prm, Aad, CONS.Esp

'Case "LoadStringA"
'If USERCONFIG(1) <> 0 Then
'LdString FncCall, Prm, Aad, CONS.Esp, 0
'End If

'Case "LoadStringW"
'If USERCONFIG(1) <> 0 Then
'LdString FncCall, Prm, Aad, CONS.Esp, 1
'End If


'KERNEL32

'Case "CreateThread"
'ThreadCr FncCall, Prm, Aad, CONS.Esp

'Case "TerminateThread"
'ThreadKl FncCall, Prm, Aad, CONS.Esp

'Case "ExitThread"
'ThreadEx FncCall, Aad, CONS.Esp


'Case "GetLocalTime"
'ReTime FncCall, Prm, Aad, CONS.Esp






'DOBRO!!!!!!!*******

Case "CreateFileW"
If SPConfig(0) = 1 Then
If CheckInUni(ThreadId, 0) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo InCrF
End If

Case "CreateFileA"
If SPConfig(0) = 1 Then
AddInUni ThreadId, 0
WWver = 0
InCrF:
ReDim XTEXTS(8)
ReDim XPARAMS(6)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 28, ByVal 0&)
XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XFNAME = GetStringFPTR(XPARAMS(0), 512, WWver)
XTEXTS(2) = "Filename: " & XFNAME
XTEXTS(3) = "Access: " & Hex(XPARAMS(1))
XTEXTS(4) = "Share: " & Hex(XPARAMS(2))
XTEXTS(5) = "Security Attributes At: " & Hex(XPARAMS(3))
XTEXTS(6) = "Distribution: " & Hex(XPARAMS(4))
XTEXTS(7) = "File Attributes: " & Hex(XPARAMS(5))
XTEXTS(8) = "Template File Handle: " & Hex(XPARAMS(6))
SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 0, _
"API " & FncCall & Chr(0) & "Filename: " & XFNAME & Chr(0) & "File Handle: "
SPBREAK.Show 1
End If

'Zbog 16-bitne kompatibilnosti
Case "OpenFile"
If SPConfig(0) = 1 Then
ReDim XTEXTS(4)
ReDim XPARAMS(2)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 12, ByVal 0&)
XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XFNAME = GetStringFPTR(XPARAMS(0), 512, 0)
XTEXTS(2) = "Filename: " & XFNAME
XTEXTS(3) = "Open File Stucture At: " & Hex(XPARAMS(1))
XTEXTS(4) = "Style: " & Hex(XPARAMS(2))
SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 1, _
"API " & FncCall & Chr(0) & "Filename: " & XFNAME & Chr(0) & "File Handle: "
SPBREAK.Show 1
End If

Case "CreateFileMappingW"
If SPConfig(0) = 1 Then
If CheckInUni(ThreadId, 2) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inCFM
End If

Case "CreateFileMappingA"
If SPConfig(0) = 1 Then
AddInUni ThreadId, 2
WWver = 0
inCFM:
ReDim XTEXTS(7)  '-
ReDim XPARAMS(5)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 24, ByVal 0&)
XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "File Handle: " & Hex(XPARAMS(0))
XTEXTS(3) = "Security Attributes At: " & Hex(XPARAMS(1))
XTEXTS(4) = "Protection:" & Hex(XPARAMS(2))
XTEXTS(5) = "Length HI (64 bit): " & Hex(XPARAMS(3))
XTEXTS(6) = "Length LO (64 bit): " & Hex(XPARAMS(4))

If XPARAMS(5) = 0 Then
FObjS = "<No Name>"
Else
FObjS = GetStringFPTR(XPARAMS(5), 512, WWver)
End If

XTEXTS(7) = "Mapping Object Name: " & FObjS
SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 2, _
"API " & FncCall & Chr(0) & "File Handle: " & Hex(XPARAMS(0)) & Chr(0) & "Mapping Object Name: " & FObjS & Chr(0) & "File Mapping Handle: "
SPBREAK.Show 1
End If

Case "OpenFileMappingW"
If SPConfig(0) = 1 Then
If CheckInUni(ThreadId, 3) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inOFM
End If

Case "OpenFileMappingA"
If SPConfig(0) = 1 Then
AddInUni ThreadId, 3
WWver = 0
inOFM:
ReDim XTEXTS(4)  '-
ReDim XPARAMS(2)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 12, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "Access: " & Hex(XPARAMS(0))
XTEXTS(3) = "Inherit Handle: " & Hex(XPARAMS(1))

If XPARAMS(2) = 0 Then
FObjS = "<No Name>"
Else
FObjS = GetStringFPTR(XPARAMS(2), 512, WWver)
End If

XTEXTS(4) = "Mapping Object Name: " & FObjS
SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 3, _
"API " & FncCall & Chr(0) & "Mapping Object Name: " & FObjS & Chr(0) & "File Mapping Handle: "
SPBREAK.Show 1
End If


Case "ReadFileEx"
If SPConfig(1) = 1 Then
WWver = 1
GoTo inRFI
End If

Case "ReadFile"

If SPConfig(1) = 1 Then
WWver = 0
inRFI:

Dim ExtDat As String

If WWver = 1 Then
ReDim XTEXTS(7)
ReDim XPARAMS(5)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 24, ByVal 0&)
XTEXTS(7) = "Async Read Procedure At: " & Hex(XPARAMS(5))
ExtDat = XTEXTS(7)
Aad = XPARAMS(5) 'BREAK ON ASYNC READ
Else
ReDim XTEXTS(6)
ReDim XPARAMS(4)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 20, ByVal 0&)
ExtDat = "Result: "
End If

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "File Handle: " & Hex(XPARAMS(0))
XTEXTS(3) = "Buffer At: " & Hex(XPARAMS(1))
XTEXTS(4) = "Buffer Length: " & Hex(XPARAMS(2))
XTEXTS(5) = "Bytes Read Reference At: " & Hex(XPARAMS(3))
XTEXTS(6) = "Overlapped Structure At:" & Hex(XPARAMS(4))

SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 4, _
"API " & FncCall & Chr(0) & "File Handle: " & Hex(XPARAMS(0)) & Chr(0) & "Buffer At :[" & Hex(XPARAMS(1)) & _
" To " & Hex(AddBy8(XPARAMS(1), XPARAMS(2))) & "]" & Chr(0) & ExtDat
SPBREAK.Show 1

End If


Case "WriteFileEx"
If SPConfig(1) = 1 Then
WWver = 1
GoTo inWFI
End If

Case "WriteFile"

If SPConfig(1) = 1 Then
WWver = 0
inWFI:

Dim ExtDat2 As String

If WWver = 1 Then
ReDim XTEXTS(6)
ReDim XPARAMS(4)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 20, ByVal 0&)
XTEXTS(6) = "Async Read Procedure At: " & Hex(XPARAMS(4))
ExtDat2 = XTEXTS(6)
Aad = XPARAMS(4) 'BREAK ON ASYNC READ
Else
ReDim XTEXTS(5)
ReDim XPARAMS(3)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 16, ByVal 0&)
ExtDat2 = "Result: "
End If

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "File Handle: " & Hex(XPARAMS(0))
XTEXTS(3) = "Buffer At: " & Hex(XPARAMS(1))
XTEXTS(4) = "Buffer Length: " & Hex(XPARAMS(2))
XTEXTS(5) = "Overlapped Structure At:" & Hex(XPARAMS(3))

SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 5, _
"API " & FncCall & Chr(0) & "File Handle: " & Hex(XPARAMS(0)) & Chr(0) & "Buffer At :[" & Hex(XPARAMS(1)) & _
" To " & Hex(AddBy8(XPARAMS(1), XPARAMS(2))) & "]" & Chr(0) & ExtDat2
SPBREAK.Show 1
End If



Case "MapViewOfFileEx"
If SPConfig(1) = 1 Then
If CheckInUni(ThreadId, 6) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inMVOF
End If

Case "MapViewOfFile"

If SPConfig(1) = 1 Then
AddInUni ThreadId, 6
WWver = 0
inMVOF:

Dim ExtDat3 As String
If WWver = 1 Then
ReDim XTEXTS(7)
ReDim XPARAMS(5)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 24, ByVal 0&)
XTEXTS(7) = "Suggested Base Address: " & Hex(XPARAMS(5)) & Chr(0) & "Map Begin At: "
ExtDat3 = XTEXTS(7)
Else
ReDim XTEXTS(6)
ReDim XPARAMS(4)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 20, ByVal 0&)
ExtDat3 = "Map Begin At: "
End If

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "File Mapping Handle: " & Hex(XPARAMS(0))
XTEXTS(3) = "Desired Access:" & Hex(XPARAMS(1))
XTEXTS(4) = "Length HI (64 bit):" & Hex(XPARAMS(2))
XTEXTS(5) = "Length LO (64 bit):" & Hex(XPARAMS(3))

If XPARAMS(4) = 0 Then
XTEXTS(6) = "Number Of Bytes: <ENTIRE FILE>"
Else
XTEXTS(6) = "Number Of Bytes:" & Hex(XPARAMS(4))
End If

SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 6, _
"API " & FncCall & Chr(0) & "File Mapping Handle: " & Hex(XPARAMS(0)) & Chr(0) & ExtDat3
SPBREAK.Show 1

End If



Case "CreateDirectoryW"
If SPConfig(2) = 1 Then
If CheckInUni(ThreadId, 7) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inCDWW
End If

Case "CreateDirectoryA"
If SPConfig(2) = 1 Then
AddInUni ThreadId, 7
WWver = 0
inCDWW:

ReDim XTEXTS(3)  '-
ReDim XPARAMS(1)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 8, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XFNAME = GetStringFPTR(XPARAMS(0), 512, WWver)
XTEXTS(2) = "Path: " & XFNAME
XTEXTS(3) = "Security Attributes At: " & Hex(XPARAMS(1))

SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 7, _
"API " & FncCall & Chr(0) & "Path: " & XFNAME & Chr(0) & "Result: "
SPBREAK.Show 1

End If



Case "CreateDirectoryExW"
If SPConfig(2) = 1 Then
If CheckInUni(ThreadId, 8) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inCDWWE
End If

Case "CreateDirectoryExA"
If SPConfig(2) = 1 Then
AddInUni ThreadId, 8
WWver = 0
inCDWWE:

ReDim XTEXTS(4)  '-
ReDim XPARAMS(2)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 12, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XFNAME = GetStringFPTR(XPARAMS(0), 512, WWver)
XFNAME2 = GetStringFPTR(XPARAMS(1), 512, WWver)
XTEXTS(2) = "Template Path: " & XFNAME
XTEXTS(3) = "Path: " & XFNAME2
XTEXTS(4) = "Security Attributes At: " & Hex(XPARAMS(2))

SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 8, _
"API " & FncCall & Chr(0) & "Template Path: " & XFNAME & Chr(0) & "Path: " & XFNAME2 & Chr(0) & "Result: "
SPBREAK.Show 1

End If




Case "RemoveDirectoryW"
If SPConfig(2) = 1 Then
If CheckInUni(ThreadId, 9) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inRDE
End If

Case "RemoveDirectoryA"
If SPConfig(2) = 1 Then
AddInUni ThreadId, 9
WWver = 0
inRDE:
ReDim XTEXTS(2)  '-
ReDim XPARAMS(0)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 4, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XFNAME = GetStringFPTR(XPARAMS(0), 512, WWver)
XTEXTS(2) = "Path: " & XFNAME
SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 9, _
"API " & FncCall & Chr(0) & "Path: " & XFNAME & Chr(0) & "Result: "
SPBREAK.Show 1
End If



Case "DeleteFileW"
If SPConfig(1) = 1 Then
If CheckInUni(ThreadId, 10) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inDLF
End If

Case "DeleteFileA"
If SPConfig(1) = 1 Then
AddInUni ThreadId, 10
WWver = 0
inDLF:
ReDim XTEXTS(2)  '-
ReDim XPARAMS(0)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 4, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XFNAME = GetStringFPTR(XPARAMS(0), 512, WWver)
XTEXTS(2) = "Filename: " & XFNAME
SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 10, _
"API " & FncCall & Chr(0) & "Filename: " & XFNAME & Chr(0) & "Result: "
SPBREAK.Show 1
End If



Case "CopyFileW"
If SPConfig(1) = 1 Then
If CheckInUni(ThreadId, 11) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inCPY
End If

Case "CopyFileA"
If SPConfig(1) = 1 Then
AddInUni ThreadId, 11
WWver = 0
inCPY:
ReDim XTEXTS(4)  '-
ReDim XPARAMS(2)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 12, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XFNAME = GetStringFPTR(XPARAMS(0), 512, WWver)
XFNAME2 = GetStringFPTR(XPARAMS(1), 512, WWver)
XTEXTS(2) = "Filename: " & XFNAME
XTEXTS(3) = "Copy: " & XFNAME2
XTEXTS(4) = "Fail if Exist: " & Hex(XPARAMS(2))

SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 11, _
"API " & FncCall & Chr(0) & "Filename: " & XFNAME & Chr(0) & "Copy: " & XFNAME2 & Chr(0) & "Result: "
SPBREAK.Show 1
End If



Case "CopyFileExW"
If SPConfig(1) = 1 Then
If CheckInUni(ThreadId, 12) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inCPYX
End If

Case "CopyFileExA"
If SPConfig(1) = 1 Then
AddInUni ThreadId, 12
WWver = 0
inCPYX:
ReDim XTEXTS(7)  '-
ReDim XPARAMS(5)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 24, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XFNAME = GetStringFPTR(XPARAMS(0), 512, WWver)
XFNAME2 = GetStringFPTR(XPARAMS(1), 512, WWver)
XTEXTS(2) = "Filename: " & XFNAME
XTEXTS(3) = "Copy: " & XFNAME2
XTEXTS(4) = "Progress Procedure At: " & Hex(XPARAMS(2))
XTEXTS(5) = "lpData At: " & Hex(XPARAMS(3))
XTEXTS(6) = "lpCancel At: " & Hex(XPARAMS(4))
XTEXTS(7) = "Copy Flag: " & Hex(XPARAMS(5))
SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 12, _
"API " & FncCall & Chr(0) & "Filename: " & XFNAME & Chr(0) & "Copy: " & XFNAME2 & Chr(0) & XTEXTS(4) & Chr(0) & "Result: "
SPBREAK.Show 1
End If


Case "MoveFileW"
If SPConfig(3) = 1 Then
If CheckInUni(ThreadId, 13) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inMVF
End If

Case "MoveFileA"
If SPConfig(3) = 1 Then
AddInUni ThreadId, 13
WWver = 0
inMVF:
ReDim XTEXTS(3)  '-
ReDim XPARAMS(1)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 8, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XFNAME = GetStringFPTR(XPARAMS(0), 512, WWver)
XFNAME2 = GetStringFPTR(XPARAMS(1), 512, WWver)
XTEXTS(2) = "Filename or Directory: " & XFNAME
XTEXTS(3) = "Destination Filename or Directory: " & XFNAME2


SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 13, _
"API " & FncCall & Chr(0) & XTEXTS(2) & Chr(0) & XTEXTS(3) & Chr(0) & "Result: "
SPBREAK.Show 1
End If





Case "MoveFileExW"
If SPConfig(3) = 1 Then
If CheckInUni(ThreadId, 14) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inMVFX
End If

Case "MoveFileExA"
If SPConfig(3) = 1 Then
AddInUni ThreadId, 14
WWver = 0
inMVFX:
ReDim XTEXTS(4)  '-
ReDim XPARAMS(2)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 12, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XFNAME = GetStringFPTR(XPARAMS(0), 512, WWver)
XFNAME2 = GetStringFPTR(XPARAMS(1), 512, WWver)
XTEXTS(2) = "Filename or Directory: " & XFNAME
XTEXTS(3) = "Destination Filename or Directory: " & XFNAME2

XTEXTS(4) = "Flags: " & Hex(XPARAMS(2))
If XPARAMS(2) = 4 Then XTEXTS(4) = XTEXTS(4) & " <AFTER REBOOT>"

SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 14, _
"API " & FncCall & Chr(0) & XTEXTS(2) & Chr(0) & XTEXTS(3) & Chr(0) & "Result: "
SPBREAK.Show 1
End If

Case "UnmapViewOfFile"
If SPConfig(1) = 1 Then
ReDim XTEXTS(2)  '-
ReDim XPARAMS(0)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 4, ByVal 0&)
XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "Base Address: " & Hex(XPARAMS(0))
SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 15, _
"API " & FncCall & Chr(0) & XTEXTS(2) & Chr(0) & "Result: "
SPBREAK.Show 1
End If




Case "GetVolumeInformationW"
If SPConfig(4) = 1 Then
If CheckInUni(ThreadId, 16) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inGVI
End If

Case "GetVolumeInformationA"
If SPConfig(4) = 1 Then
AddInUni ThreadId, 16
WWver = 0
inGVI:

ReDim XTEXTS(9)  '-
ReDim XPARAMS(7)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 32, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XFNAME = GetStringFPTR(XPARAMS(0), 512, WWver)
XTEXTS(2) = "Volume Root Path: " & XFNAME
XTEXTS(3) = "Buffer For Name At: " & Hex(XPARAMS(1))
XTEXTS(4) = "Buffer Length: " & Hex(XPARAMS(2))
XTEXTS(5) = "Serial Buffer (DWORD-out) At: " & Hex(XPARAMS(3))
XTEXTS(6) = "Max Filename Length (DWORD-out) At: " & Hex(XPARAMS(4))
XTEXTS(7) = "File System Flag (DWORD-out) At: " & Hex(XPARAMS(5))
XTEXTS(8) = "File System Name Buffer At: " & Hex(XPARAMS(6))
XTEXTS(9) = "Buffer Length: " & Hex(XPARAMS(7))


SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 16, _
"API " & FncCall & Chr(0) & "Volume Root Path: " & XFNAME & Chr(0) & _
"Name Buffer At: [" & Hex(XPARAMS(1)) & " To " & Hex(AddBy8(XPARAMS(1), XPARAMS(2))) & "]" & Chr(0) & _
XTEXTS(5) & Chr(0) & XTEXTS(6) & Chr(0) & XTEXTS(7) & Chr(0) & _
"File System Name Buffer At: [" & Hex(XPARAMS(6)) & " To " & Hex(AddBy8(XPARAMS(6), XPARAMS(7))) & "]" & Chr(0) & "Result: "
SPBREAK.Show 1

End If



Case "GetDriveTypeW"
If SPConfig(4) = 1 Then
If CheckInUni(ThreadId, 17) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inGTD
End If

Case "GetDriveTypeA"
If SPConfig(4) = 1 Then
AddInUni ThreadId, 17
WWver = 0
inGTD:
ReDim XTEXTS(2)  '-
ReDim XPARAMS(0)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 4, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XFNAME = GetStringFPTR(XPARAMS(0), 512, WWver)

XTEXTS(2) = "Volume Root Path: " & XFNAME

SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 17, _
"API " & FncCall & Chr(0) & XTEXTS(2) & Chr(0) & "Volume Type: "
SPBREAK.Show 1
End If



Case "GetLogicalDrives"
If SPConfig(4) = 1 Then
ReDim XTEXTS(1)  '-
XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 18, _
"API " & FncCall & Chr(0) & "Logical Drives Bit Pattern: "
SPBREAK.Show 1
End If



Case "GetLogicalDriveStringsW"
If SPConfig(4) = 1 Then
If CheckInUni(ThreadId, 19) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inGLDS
End If

Case "GetLogicalDriveStringsA"
If SPConfig(4) = 1 Then
AddInUni ThreadId, 19
WWver = 0
inGLDS:
ReDim XTEXTS(3)  '-
ReDim XPARAMS(1)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 8, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "Buffer Length: " & Hex(XPARAMS(0))
XTEXTS(3) = "Buffer At: " & Hex(XPARAMS(1))

SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 19, _
"API " & FncCall & Chr(0) & _
"Logical Drives Buffer At: [" & Hex(XPARAMS(1)) & " To " & Hex(AddBy8(XPARAMS(1), XPARAMS(0))) & "]" & Chr(0) & "Result: "
SPBREAK.Show 1
End If


'*REGISTRY

Case "RegOpenKeyW"
If SPConfig(5) = 1 Then
If CheckInUni(ThreadId, 20) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inROK
End If

Case "RegOpenKeyA"
If SPConfig(5) = 1 Then
AddInUni ThreadId, 20
WWver = 0
inROK:
ReDim XTEXTS(4)  '-
ReDim XPARAMS(2)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 12, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "Hkey: " & KeyToName(XPARAMS(0))

If XPARAMS(1) = 0 Then
XFNAME = "<NULL>"
Else
XFNAME = GetStringFPTR(XPARAMS(1), 512, WWver)
End If

XTEXTS(3) = "SubKey: " & XFNAME
XTEXTS(4) = "HKey Result At: " & Hex(XPARAMS(2))

SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 20, _
"API " & FncCall & Chr(0) & _
XTEXTS(2) & Chr(0) & XTEXTS(3) & Chr(0) & XTEXTS(4) & Chr(0) & "Result: "
SPBREAK.Show 1
End If


Case "RegOpenKeyExW"
If SPConfig(5) = 1 Then
If CheckInUni(ThreadId, 21) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inROKX
End If

Case "RegOpenKeyExA"
If SPConfig(5) = 1 Then
AddInUni ThreadId, 21
WWver = 0
inROKX:
ReDim XTEXTS(6)  '-
ReDim XPARAMS(4)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 20, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "Hkey: " & KeyToName(XPARAMS(0))


If XPARAMS(1) = 0 Then
XFNAME = "<NULL>"
Else
XFNAME = GetStringFPTR(XPARAMS(1), 512, WWver)
End If

XTEXTS(3) = "SubKey: " & XFNAME
XTEXTS(4) = "Options: " & Hex(XPARAMS(2))
XTEXTS(5) = "Access: " & Hex(XPARAMS(3))
XTEXTS(6) = "HKey Result At: " & Hex(XPARAMS(4))

SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 21, _
"API " & FncCall & Chr(0) & _
XTEXTS(2) & Chr(0) & XTEXTS(3) & Chr(0) & XTEXTS(5) & Chr(0) & XTEXTS(6) & Chr(0) & "Result: "
SPBREAK.Show 1
End If




Case "RegCloseKey"
If SPConfig(5) = 1 Then

ReDim XTEXTS(3)  '-
ReDim XPARAMS(0)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 4, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "Hkey: " & KeyToName(XPARAMS(0))

SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 22, _
"API " & FncCall & Chr(0) & _
XTEXTS(2) & Chr(0) & "Result: "
SPBREAK.Show 1
End If




Case "RegCreateKeyW"
If SPConfig(5) = 1 Then
If CheckInUni(ThreadId, 23) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inRCK
End If

Case "RegCreateKeyA"
If SPConfig(5) = 1 Then
AddInUni ThreadId, 23
WWver = 0
inRCK:
ReDim XTEXTS(4)  '-
ReDim XPARAMS(2)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 12, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)

XTEXTS(2) = "Hkey: " & KeyToName(XPARAMS(0))

If XPARAMS(1) = 0 Then
XFNAME = "<NULL>"
Else
XFNAME = GetStringFPTR(XPARAMS(1), 512, WWver)
End If

XTEXTS(3) = "SubKey: " & XFNAME
XTEXTS(4) = "HKey Result At: " & Hex(XPARAMS(2))

SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 23, _
"API " & FncCall & Chr(0) & _
XTEXTS(2) & Chr(0) & XTEXTS(3) & Chr(0) & XTEXTS(4) & Chr(0) & "Result: "
SPBREAK.Show 1
End If




Case "RegCreateKeyExW"
If SPConfig(5) = 1 Then
If CheckInUni(ThreadId, 24) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inCKEE
End If

Case "RegCreateKeyExA"
If SPConfig(5) = 1 Then
AddInUni ThreadId, 24
WWver = 0
inCKEE:
ReDim XTEXTS(10)  '-
ReDim XPARAMS(8)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 36, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "Hkey: " & KeyToName(XPARAMS(0))


If XPARAMS(1) = 0 Then
XFNAME = "<NULL>"
Else
XFNAME = GetStringFPTR(XPARAMS(1), 512, WWver)
End If

XTEXTS(3) = "SubKey: " & XFNAME

XTEXTS(4) = "Reserved: " & Hex(XPARAMS(2))

If XPARAMS(3) = 0 Then
XFNAME2 = "<NULL>"
Else
XFNAME2 = GetStringFPTR(XPARAMS(3), 512, WWver)
End If

XTEXTS(5) = "Class Type: " & XFNAME2

XTEXTS(6) = "Options: " & Hex(XPARAMS(4))
XTEXTS(7) = "Access: " & Hex(XPARAMS(5))
XTEXTS(8) = "Security Attributes At: " & Hex(XPARAMS(6))

XTEXTS(9) = "HKey Result At: " & Hex(XPARAMS(7))
XTEXTS(10) = "Disposition At: " & Hex(XPARAMS(8))

SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 24, _
"API " & FncCall & Chr(0) & _
XTEXTS(2) & Chr(0) & XTEXTS(3) & Chr(0) & XTEXTS(5) & Chr(0) & XTEXTS(6) & _
Chr(0) & XTEXTS(9) & Chr(0) & XTEXTS(10) & Chr(0) & "Result: "
SPBREAK.Show 1
End If





Case "RegDeleteKeyW"
If SPConfig(5) = 1 Then
If CheckInUni(ThreadId, 25) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inRDKK
End If

Case "RegDeleteKeyA"
If SPConfig(5) = 1 Then
AddInUni ThreadId, 25
WWver = 0
inRDKK:
ReDim XTEXTS(3)  '-
ReDim XPARAMS(1)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 8, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "Hkey: " & KeyToName(XPARAMS(0))

If XPARAMS(1) = 0 Then
XFNAME = "<NULL>"
Else
XFNAME = GetStringFPTR(XPARAMS(1), 512, WWver)
End If

XTEXTS(3) = "SubKey: " & XFNAME


SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 25, _
"API " & FncCall & Chr(0) & _
XTEXTS(2) & Chr(0) & XTEXTS(3) & Chr(0) & "Result: "
SPBREAK.Show 1
End If





Case "RegQueryValueW"
If SPConfig(5) = 1 Then
If CheckInUni(ThreadId, 26) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inRQVV
End If

Case "RegQueryValueA"
If SPConfig(5) = 1 Then
AddInUni ThreadId, 26
WWver = 0
inRQVV:
ReDim XTEXTS(5)  '-
ReDim XPARAMS(3)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 16, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "Hkey: " & KeyToName(XPARAMS(0))

If XPARAMS(1) = 0 Then
XFNAME = "<NULL>"
Else
XFNAME = GetStringFPTR(XPARAMS(1), 512, WWver)
End If

XTEXTS(3) = "SubKey: " & XFNAME

XTEXTS(4) = "Value Buffer At: " & Hex(XPARAMS(2))
XTEXTS(5) = "Buffer Length At [DWORD-out]: " & Hex(XPARAMS(3))


SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 26, _
"API " & FncCall & Chr(0) & _
XTEXTS(2) & Chr(0) & XTEXTS(3) & Chr(0) & _
XTEXTS(4) & Chr(0) & XTEXTS(5) & Chr(0) & "Result: "
SPBREAK.Show 1
End If



Case "RegQueryValueExW"
If SPConfig(5) = 1 Then
If CheckInUni(ThreadId, 27) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inRQVVX
End If

Case "RegQueryValueExA"
If SPConfig(5) = 1 Then
AddInUni ThreadId, 27
WWver = 0
inRQVVX:
ReDim XTEXTS(7)  '-
ReDim XPARAMS(5)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 24, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "Hkey: " & KeyToName(XPARAMS(0))

If XPARAMS(1) = 0 Then
XFNAME = "<NULL>"
Else
XFNAME = GetStringFPTR(XPARAMS(1), 512, WWver)
End If

XTEXTS(3) = "Value Name: " & XFNAME
XTEXTS(4) = "Reserved: " & Hex(XPARAMS(2))
XTEXTS(5) = "Type At [DWORD-out]: " & Hex(XPARAMS(3))


XTEXTS(6) = "Value Buffer At: " & Hex(XPARAMS(4))
XTEXTS(7) = "Buffer Length At [DWORD-out]: " & Hex(XPARAMS(5))


SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 27, _
"API " & FncCall & Chr(0) & _
XTEXTS(2) & Chr(0) & XTEXTS(3) & Chr(0) & XTEXTS(5) & Chr(0) & _
XTEXTS(6) & Chr(0) & XTEXTS(7) & Chr(0) & "Result: "
SPBREAK.Show 1
End If




Case "RegSetValueW"
If SPConfig(5) = 1 Then
If CheckInUni(ThreadId, 28) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inRSV
End If

Case "RegSetValueA"
If SPConfig(5) = 1 Then
AddInUni ThreadId, 28
WWver = 0
inRSV:
ReDim XTEXTS(6)  '-
ReDim XPARAMS(4)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 20, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "Hkey: " & KeyToName(XPARAMS(0))

If XPARAMS(1) = 0 Then
XFNAME = "<NULL>"
Else
XFNAME = GetStringFPTR(XPARAMS(1), 512, WWver)
End If

XTEXTS(3) = "SubKey: " & XFNAME
XTEXTS(4) = "Type: " & XPARAMS(2)
XTEXTS(5) = "Value Buffer At: " & Hex(XPARAMS(3))
XTEXTS(6) = "Buffer Length : " & Hex(XPARAMS(4))


SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 28, _
"API " & FncCall & Chr(0) & _
XTEXTS(2) & Chr(0) & XTEXTS(3) & Chr(0) & XTEXTS(4) & Chr(0) & _
"Value Buffer At: [" & Hex(XPARAMS(3)) & " To " & Hex(AddBy8(XPARAMS(3), XPARAMS(4))) & "]" & Chr(0) & "Result: "
SPBREAK.Show 1
End If



Case "RegSetValueExW"
If SPConfig(5) = 1 Then
If CheckInUni(ThreadId, 29) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inRSVX
End If

Case "RegSetValueExA"
If SPConfig(5) = 1 Then
AddInUni ThreadId, 29
WWver = 0
inRSVX:
ReDim XTEXTS(7)  '-
ReDim XPARAMS(5)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 24, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "Hkey: " & KeyToName(XPARAMS(0))

If XPARAMS(1) = 0 Then
XFNAME = "<NULL>"
Else
XFNAME = GetStringFPTR(XPARAMS(1), 512, WWver)
End If

XTEXTS(3) = "Value Name: " & XFNAME
XTEXTS(4) = "Reserved: " & XPARAMS(2)
XTEXTS(5) = "Type: " & XPARAMS(3)
XTEXTS(6) = "Value Buffer At: " & Hex(XPARAMS(4))
XTEXTS(7) = "Buffer Length : " & Hex(XPARAMS(5))


SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 29, _
"API " & FncCall & Chr(0) & _
XTEXTS(2) & Chr(0) & XTEXTS(3) & Chr(0) & XTEXTS(5) & Chr(0) & _
"Value Buffer At: [" & Hex(XPARAMS(4)) & " To " & Hex(AddBy8(XPARAMS(4), XPARAMS(5))) & "]" & Chr(0) & "Result: "
SPBREAK.Show 1
End If




Case "RegDeleteValueW"
If SPConfig(5) = 1 Then
If CheckInUni(ThreadId, 30) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inDVV
End If

Case "RegDeleteValueA"
If SPConfig(5) = 1 Then
AddInUni ThreadId, 30
WWver = 0
inDVV:
ReDim XTEXTS(3)  '-
ReDim XPARAMS(1)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 8, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "Hkey: " & KeyToName(XPARAMS(0))

If XPARAMS(1) = 0 Then
XFNAME = "<NULL>"
Else
XFNAME = GetStringFPTR(XPARAMS(1), 512, WWver)
End If

XTEXTS(3) = "Value Name: " & XFNAME



SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 30, _
"API " & FncCall & Chr(0) & _
XTEXTS(2) & Chr(0) & XTEXTS(3) & Chr(0) & "Result: "
SPBREAK.Show 1
End If



Case "RegLoadKeyW"
If SPConfig(5) = 1 Then
If CheckInUni(ThreadId, 31) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inLKY
End If

Case "RegLoadKeyA"
If SPConfig(5) = 1 Then
AddInUni ThreadId, 31
WWver = 0
inLKY:
ReDim XTEXTS(4)  '-
ReDim XPARAMS(2)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 12, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "Hkey: " & KeyToName(XPARAMS(0))

If XPARAMS(1) = 0 Then
XFNAME = "<NULL>"
Else
XFNAME = GetStringFPTR(XPARAMS(1), 512, WWver)
End If

XTEXTS(3) = "SubKey: " & XFNAME
XFNAME2 = GetStringFPTR(XPARAMS(2), 512, WWver)

XTEXTS(4) = "Filename: " & XFNAME2


SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 31, _
"API " & FncCall & Chr(0) & _
XTEXTS(2) & Chr(0) & XTEXTS(3) & Chr(0) & _
XTEXTS(4) & Chr(0) & "Result: "
SPBREAK.Show 1
End If





Case "RegSaveKeyW"
If SPConfig(5) = 1 Then
If CheckInUni(ThreadId, 32) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inLSY
End If

Case "RegSaveKeyA"
If SPConfig(5) = 1 Then
AddInUni ThreadId, 32
WWver = 0
inLSY:
ReDim XTEXTS(4)  '-
ReDim XPARAMS(2)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 12, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)
XTEXTS(2) = "Hkey: " & KeyToName(XPARAMS(0))

If XPARAMS(1) = 0 Then
XFNAME = "<NULL>"
Else
XFNAME = GetStringFPTR(XPARAMS(1), 512, WWver)
End If

XTEXTS(3) = "Filename: " & XFNAME


XTEXTS(4) = "Security Attributes At: " & Hex(XPARAMS(2))


SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 32, _
"API " & FncCall & Chr(0) & _
XTEXTS(2) & Chr(0) & XTEXTS(3) & Chr(0) & _
XTEXTS(4) & Chr(0) & "Result: "
SPBREAK.Show 1
End If





Case "RegConnectRegistryW"
If SPConfig(5) = 1 Then
If CheckInUni(ThreadId, 33) <> -1 Then Exit Sub 'Surely duplicate call from AnsiVerison
WWver = 1
GoTo inRCR
End If

Case "RegConnectRegistryA"
If SPConfig(5) = 1 Then
AddInUni ThreadId, 33
WWver = 0
inRCR:
ReDim XTEXTS(4)  '-
ReDim XPARAMS(2)
Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(CONS.Esp, 4), XPARAMS(0), 12, ByVal 0&)

XTEXTS(0) = "API " & FncCall
XTEXTS(1) = "Ret Address At: " & Hex(Aad)


If XPARAMS(1) = 0 Then
XFNAME = "<NULL>"
Else
XFNAME = GetStringFPTR(XPARAMS(0), 512, WWver)
End If

XTEXTS(2) = "Machine: " & XFNAME
XTEXTS(3) = "Hkey: " & KeyToName(XPARAMS(1))
XTEXTS(4) = "HKey Result At: " & Hex(XPARAMS(2))

SPBREAK.ReReadF XTEXTS, XPARAMS, ThreadId, PROCRF, Aad, 33, _
"API " & FncCall & Chr(0) & _
XTEXTS(2) & Chr(0) & XTEXTS(3) & Chr(0) & XTEXTS(4) & Chr(0) & "Result: "
SPBREAK.Show 1
End If




'CFMData(0) = "API " & FncCall
'CFMData(1) = "Return Address At: " & Hex(RetAdr)
'CFMData(2) = "On Hwnd:" & Hex(Params(0))
'CFMData(3) = "Id:" & Hex(Params(1))
'CFMData(4) = "Tick Timer:" & Hex(Params(2))
'If Params(3) = 0 Then
'CFMData(5) = "Using WM_TIMER"
'Else
'CFMData(5) = "Timer Proc At:" & Hex(Params(3))
'End If
'SPBREAK.ReReadF CFMData, ThreadId, PROCRF
'SPBREAK.Show 1
'AddBPX RetAdr, 7




'If KERNELCONFIG(4) <> 0 Then
'CreateF FncCall, Prm, Aad, CONS.Esp
'End If

'Case Is = "ReadFile", "WriteFile"
'If KERNELCONFIG(4) <> 0 Then
'ReadF FncCall, Prm, Aad, CONS.Esp, 0
'End If

'Case Is = "ReadFileEx", "WriteFileEx"
'If KERNELCONFIG(4) <> 0 Then
'ReadF FncCall, Prm, Aad, CONS.Esp, 1
'End If


'Case "CreateFileMappingA"
'If KERNELCONFIG(4) <> 0 Then
'CreateFMP FncCall, Prm, Aad, CONS.Esp
'End If


'Case "MapViewOfFile"
'If KERNELCONFIG(4) <> 0 Then
'MapVF FncCall, Prm, Aad, CONS.Esp, 0
'End If

'Case "MapViewOfFileEx"
'If KERNELCONFIG(4) <> 0 Then
'MapVF FncCall, Prm, Aad, CONS.Esp, 1
'End If


'Case "RegOpenKeyA"
'RegOpK FncCall, Prm, Aad, CONS.Esp, 0


'Case "RegOpenKeyExA"
'RegOpK FncCall, Prm, Aad, CONS.Esp, 1

'Case "RegQueryValueA"
'RegQuV FncCall, Prm, Aad, CONS.Esp


End Select


Exit Sub
'Dalje:
'On Error GoTo 0
End Sub

'Private Sub Alloc(FncCall As String, Params() As Long, ByRef RetAdr As Long, ByRef StackPtr As Long)
'ReDim Params(3) As Long
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(StackPtr, 4), Params(0), 16, ByVal 0&)
'APIS.KERNELF.AddInAPI "API " & FncCall
'APIS.KERNELF.AddInAPI "Return Address At: " & Hex(RetAdr)
'APIS.KERNELF.AddInAPI "Type Of Allocation:" & Hex(Params(2)) & " ,Protection:" & Hex(Params(3))
'APIS.KERNELF.AddInAPI "Size Of Allocation:" & Hex(Params(1))
'AddBPX RetAdr, 2
'End Sub







'Private Sub ReTime(FncCall As String, Params() As Long, ByRef RetAdr As Long, ByRef StackPtr As Long)
'Dim SYSTIME As Long
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(StackPtr, 4), SYSTIME, 4, ByVal 0&)
'APIS.KERNELF.AddInAPI "API " & FncCall
'APIS.KERNELF.AddInAPI "Return Address At: " & Hex(RetAdr)
'APIS.KERNELF.AddInAPI "System Time Struct At:" & Hex(SYSTIME)
'APIS.KERNELF.AddInAPI ""
'End Sub








'Za Timer
'Public Sub TimerRef(FncCall As String, Params() As Long, ByRef RetAdr As Long, ByRef StackPtr As Long, ByRef ThreadId As Long, ByRef PROCRF As Long)
'Dim CFMData(7) As String
'ReDim Params(3)
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(StackPtr, 4), Params(0), 16, ByVal 0&)
'CFMData(0) = "API " & FncCall
'CFMData(1) = "Return Address At: " & Hex(RetAdr)
'CFMData(2) = "On Hwnd:" & Hex(Params(0))
'CFMData(3) = "Id:" & Hex(Params(1))
'CFMData(4) = "Tick Timer:" & Hex(Params(2))
'If Params(3) = 0 Then
'CFMData(5) = "Using WM_TIMER"
'Else
'CFMData(5) = "Timer Proc At:" & Hex(Params(3))
'End If
'SPBREAK.ReReadF CFMData, ThreadId, PROCRF
'SPBREAK.Show 1
'AddBPX RetAdr, 7
'End Sub

'Public Sub TimerKl(FncCall As String, Params() As Long, ByRef RetAdr As Long, ByRef StackPtr As Long)
'ReDim Params(1)
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(StackPtr, 4), Params(0), 8, ByVal 0&)
'APIS.GENERALF.AddInAPI "API " & FncCall
'APIS.GENERALF.AddInAPI "Return Address At: " & Hex(RetAdr)
'APIS.GENERALF.AddInAPI "On Hwnd:" & Hex(Params(0))
'APIS.GENERALF.AddInAPI "Id:" & Hex(Params(1))
'APIS.GENERALF.AddInAPI ""
'End Sub

'Public Sub ThreadCr(FncCall As String, Params() As Long, ByRef RetAdr As Long, ByRef StackPtr As Long)
'ReDim Params(5)
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(StackPtr, 4), Params(0), 24, ByVal 0&)
'APIS.GENERALF.AddInAPI "API " & FncCall
'APIS.GENERALF.AddInAPI "Return Address At: " & Hex(RetAdr)
'APIS.GENERALF.AddInAPI "Thread Attributes At:" & Hex(Params(0))
'APIS.GENERALF.AddInAPI "Stack Size:" & Hex(Params(1))
'APIS.GENERALF.AddInAPI "Start Address:" & Hex(Params(2))
'APIS.GENERALF.AddInAPI "Parameter:" & Hex(Params(3))
'APIS.GENERALF.AddInAPI "Creation Flag:" & Hex(Params(4))
'APIS.GENERALF.AddInAPI "Thread Id Reference:" & Hex(Params(5))
'THRREF = Params(5)
'AddBPX RetAdr, 8
'End Sub

'Public Sub ThreadKl(FncCall As String, Params() As Long, ByRef RetAdr As Long, ByRef StackPtr As Long)
'ReDim Params(1)
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(StackPtr, 4), Params(0), 8, ByVal 0&)
'APIS.GENERALF.AddInAPI "API " & FncCall
'APIS.GENERALF.AddInAPI "Return Address At: " & Hex(RetAdr)
'APIS.GENERALF.AddInAPI "Thread Handle:" & Hex(Params(0))
'APIS.GENERALF.AddInAPI "Exit Code:" & Hex(Params(1))
'APIS.GENERALF.AddInAPI ""
'End Sub

'Public Sub ThreadEx(FncCall As String, ByRef RetAdr As Long, ByRef StackPtr As Long)
'Dim OEx As Long
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(StackPtr, 4), OEx, 4, ByVal 0&)
'APIS.GENERALF.AddInAPI "API " & FncCall
'APIS.GENERALF.AddInAPI "Return Address At: " & Hex(RetAdr)
'APIS.GENERALF.AddInAPI "Exit Code:" & Hex(OEx)
'APIS.GENERALF.AddInAPI ""
'End Sub

'Public Sub LdString(FncCall As String, Params() As Long, ByRef RetAdr As Long, ByRef StackPtr As Long, ByRef IsUW As Byte)
'ReDim Params(3)
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(StackPtr, 4), Params(0), 16, ByVal 0&)
'APIS.USERF.AddInAPI "API " & FncCall
'APIS.USERF.AddInAPI "Return Address At: " & Hex(RetAdr)

'APIS.USERF.AddInAPI "Hinstance:" & Hex(Params(0))
'APIS.USERF.AddInAPI "Resource Id:" & Hex(Params(1))
'APIS.USERF.AddInAPI "String Buffer:" & Hex(Params(2))
'APIS.USERF.AddInAPI "Buffer length:" & Hex(Params(3))

'LDRStr = Params(2)
'LDRLen = Params(3)

'If IsUW = 0 Then
'AddBPX RetAdr, 9
'Else
'AddBPX RetAdr, 10
'End If

'End Sub


'Public Sub OpenF(FncCall As String, Params() As Long, ByRef RetAdr As Long, ByRef StackPtr As Long)
'ReDim Params(2)
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(StackPtr, 4), Params(0), 12, ByVal 0&)
'APIS.KERNELF.AddInAPI "API " & FncCall
'APIS.KERNELF.AddInAPI "Return Address At: " & Hex(RetAdr)

'Dim Fname As String
'Fname = GetStringFPTR(Params(0), 256, 0)
'APIS.KERNELF.AddInAPI "Filename:" & Fname
'APIS.KERNELF.AddInAPI "OFStuct At:" & Hex(Params(1))
'APIS.KERNELF.AddInAPI "Style:" & Hex(Params(2))
'AddBPX RetAdr, 12
'End Sub

'Public Sub CreateF(FncCall As String, Params() As Long, ByRef RetAdr As Long, ByRef StackPtr As Long)
'ReDim Params(6)
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(StackPtr, 4), Params(0), 28, ByVal 0&)
'APIS.KERNELF.AddInAPI "API " & FncCall
'APIS.KERNELF.AddInAPI "Return Address At: " & Hex(RetAdr)

'Dim Fname As String
'Fname = GetStringFPTR(Params(0), 512, 0)
'APIS.KERNELF.AddInAPI "Filename:" & Fname
'APIS.KERNELF.AddInAPI "Access:" & Hex(Params(1))
'APIS.KERNELF.AddInAPI "Share:" & Hex(Params(2))
'APIS.KERNELF.AddInAPI "Security Attributes At:" & Hex(Params(3))
'APIS.KERNELF.AddInAPI "Distribution:" & Hex(Params(4))
'APIS.KERNELF.AddInAPI "File Attributes:" & Hex(Params(5))
'APIS.KERNELF.AddInAPI "Template File Handle:" & Hex(Params(6))
'AddBPX RetAdr, 12
'End Sub

'Public Sub ReadF(FncCall As String, Params() As Long, ByRef RetAdr As Long, ByRef StackPtr As Long, ByRef ExtFil As Byte)
'Dim Nms As Long
'ReDim Params(5)
'If ExtFil = 1 Then
'Nms = 24
'Else
'Nms = 20
'End If

'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(StackPtr, 4), Params(0), Nms, ByVal 0&)
'APIS.KERNELF.AddInAPI "API " & FncCall
'APIS.KERNELF.AddInAPI "Return Address At: " & Hex(RetAdr)


'APIS.KERNELF.AddInAPI "File Handle:" & Hex(Params(0))
'APIS.KERNELF.AddInAPI "Buffer At:" & Hex(Params(1))
'APIS.KERNELF.AddInAPI "Buffer Length:" & Hex(Params(2))
'APIS.KERNELF.AddInAPI "Result Ref:" & Hex(Params(3))
'APIS.KERNELF.AddInAPI "Overlapped Struct At:" & Hex(Params(4))
'If ExtFil = 1 Then
'APIS.KERNELF.AddInAPI "Completion Routine At:" & Hex(Params(5))
'End If

'BFRStr = Params(1)
'BFRLen = Params(2)
'AddBPX RetAdr, 13
'APIS.KERNELF.AddInAPI ""
'End Sub


'Public Sub CreateFMP(FncCall As String, Params() As Long, ByRef RetAdr As Long, ByRef StackPtr As Long)
'ReDim Params(5)
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(StackPtr, 4), Params(0), 24, ByVal 0&)
'APIS.KERNELF.AddInAPI "API " & FncCall
'APIS.KERNELF.AddInAPI "Return Address At: " & Hex(RetAdr)

'APIS.KERNELF.AddInAPI "File Handle:" & Hex(Params(0))
'APIS.KERNELF.AddInAPI "Security Attributes At:" & Hex(Params(1))
'APIS.KERNELF.AddInAPI "Protection:" & Hex(Params(2))
'APIS.KERNELF.AddInAPI "Length HI (64 bit):" & Hex(Params(3))
'APIS.KERNELF.AddInAPI "Length LO (64 bit):" & Hex(Params(4))
'Dim FObjS As String
'If Params(5) = 0 Then
'FObjS = "<No Name>"
'Else
'FObjS = GetStringFPTR(Params(5), 512)
'End If

'APIS.KERNELF.AddInAPI "Mapping Object Name:" & FObjS
'AddBPX RetAdr, 14
'End Sub


'Public Sub MapVF(FncCall As String, Params() As Long, ByRef RetAdr As Long, ByRef StackPtr As Long, ByRef TypeOfMP As Byte)
'If TypeOfMP = 1 Then
'ReDim Params(5)
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(StackPtr, 4), Params(0), 24, ByVal 0&)
'Else
'ReDim Params(4)
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(StackPtr, 4), Params(0), 20, ByVal 0&)
'End If

'APIS.KERNELF.AddInAPI "API " & FncCall
'APIS.KERNELF.AddInAPI "Return Address At: " & Hex(RetAdr)

'APIS.KERNELF.AddInAPI "File Mapping Handle:" & Hex(Params(0))
'APIS.KERNELF.AddInAPI "Desired Access:" & Hex(Params(1))
'APIS.KERNELF.AddInAPI "Length HI (64 bit):" & Hex(Params(2))
'APIS.KERNELF.AddInAPI "Length LO (64 bit):" & Hex(Params(3))
'APIS.KERNELF.AddInAPI "Number Of Bytes:" & Hex(Params(4))

'If TypeOfMP = 1 Then
'APIS.KERNELF.AddInAPI "Base Address:" & Hex(Params(5))
'End If

'AddBPX RetAdr, 15
'End Sub


Private Function KeyToName(ByRef Value As Long) As String
Select Case Value
Case &H80000000
KeyToName = "HKEY_CLASSES_ROOT"
Case &H80000005
KeyToName = "HKEY_CURRENT_CONFIG"
Case &H80000001
KeyToName = "HKEY_CURRENT_USER"
Case &H80000002
KeyToName = "HKEY_LOCAL_MACHINE"
Case &H80000004
KeyToName = "HKEY_PERFORMANCE_DATA"
Case &H80000003
KeyToName = "HKEY_USERS"
Case &H80000006
KeyToName = "HKEY_DYN_DATA"
Case Else
KeyToName = Hex(Value)
End Select
End Function


'Public Sub RegOpK(FncCall As String, Params() As Long, ByRef RetAdr As Long, ByRef StackPtr As Long, ByRef TypeOfEx As Byte)
'Dim SKN As String
'If TypeOfEx = 1 Then
'ReDim Params(4)
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(StackPtr, 4), Params(0), 20, ByVal 0&)
'Else
'ReDim Params(2)
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(StackPtr, 4), Params(0), 12, ByVal 0&)
'End If

'APIS.GENERALF.AddInAPI "API " & FncCall
'APIS.GENERALF.AddInAPI "Return Address At: " & Hex(RetAdr)

'APIS.GENERALF.AddInAPI "Key:" & KeyToName(Params(0))

'If Params(1) = 0 Then
'SKN = "<NULL>"
'Else
'SKN = GetStringFPTR(Params(1), 256)
'End If

'APIS.GENERALF.AddInAPI "Sub Key:" & SKN

'If TypeOfEx = 1 Then
'APIS.GENERALF.AddInAPI "Access:" & Hex(Params(3))
'HHk = Params(4)
'Else
'HHk = Params(2)
'End If
'APIS.GENERALF.AddInAPI ""
'AddBPX RetAdr, 16
'End Sub

'Public Sub RegQuV(FncCall As String, Params() As Long, ByRef RetAdr As Long, ByRef StackPtr As Long)
'ReDim Params(3)
'Dim SKN As String
'Call ReadProcessMemory(ProcessHandle, ByVal AddBy8(StackPtr, 4), Params(0), 16, ByVal 0&)
'APIS.GENERALF.AddInAPI "API " & FncCall
'APIS.GENERALF.AddInAPI "Return Address At: " & Hex(RetAdr)

'APIS.GENERALF.AddInAPI "Key:" & KeyToName(Params(0))

'If Params(1) = 0 Then
'SKN = "<NULL>"
'Else
'SKN = GetStringFPTR(Params(1), 256)
'End If
'APIS.GENERALF.AddInAPI "Sub Key:" & SKN
'DQBuff = Params(2)
'DQLen = Params(3)
'AddBPX RetAdr, 17
'End Sub






Public Sub AddBPX(ByRef RetAdr As Long, ByRef TypeOfR As Long, ByRef CustomD As String, ByRef WouldB As Byte)
Dim IsValidBP As Byte
Call GetBreakPoint(ACTIVEBREAKPOINTS, RetAdr, IsValidBP)
If IsValidBP = 0 Then
AddBreakPoint SPYBREAKPOINTS, RetAdr, 0
End If
AddInRet RetAdr, TypeOfR, CustomD, WouldB
End Sub


Public Sub AddInRet(ByRef RetAdr As Long, ByRef TypeOfR As Long, ByRef CustomD As String, ByRef WouldB As Byte)
On Error GoTo Dalje
Dim Vent(2) As String
Vent(0) = CStr(TypeOfR) '*TIP RET BP-a
Vent(1) = CStr(CustomD)
Vent(2) = CStr(WouldB)
PROCRETBRK.Add Vent, "X" & RetAdr
Exit Sub
Dalje:
On Error GoTo 0
End Sub
Public Sub RemoveInRet(ByRef RetAdr As Long)
On Error GoTo Dalje
PROCRETBRK.Remove "X" & RetAdr
Exit Sub
Dalje:
On Error GoTo 0
End Sub
Public Function GetInRet(ByRef RetAdr As Long, ByRef Vent() As String) As Byte
On Error GoTo Dalje
Vent = PROCRETBRK.Item("X" & RetAdr)
GetInRet = 1
Exit Function
Dalje:
On Error GoTo 0
End Function


Public Sub TransRet(ByRef Address As Long, ByRef Retns As Long, ByRef TransData() As String, ByRef BrkRequest As Byte, ByRef ThreadId As Long)
Dim retTp As Long
retTp = CLng(TransData(0))

AddLine Replace(TransData(1), Chr(0), " ,") & Hex(Retns) & " ,EIP: " & Hex(Address) & " ,In Thread: " & ThreadId, Form12.rt1, &HFF&

RemoveInUNI ThreadId, CLng(TransData(0))

If CLng(TransData(2)) <> 0 Then
If vbYes = MsgBox(Replace(TransData(1), Chr(0), vbCrLf) & Hex(Retns) & vbCrLf & "EIP:" & Hex(Address) & vbCrLf & "In Thread: " & ThreadId & vbCrLf & vbCrLf & "Break Execution?", vbYesNo, "Confirm") Then BrkRequest = 1
End If



End Sub

Public Sub AddInUni(ByRef ThreadId As Long, ByRef APIType As Long)
On Error GoTo Dalje
CheckUNI.Add APIType, "X" & ThreadId & APIType
Exit Sub
Dalje:
On Error GoTo 0
End Sub
Public Function CheckInUni(ByRef ThreadId As Long, ByRef APIType As Long) As Long
On Error GoTo Dalje
CheckInUni = CheckUNI.Item("X" & ThreadId & APIType)
Exit Function
Dalje:
On Error GoTo 0
CheckInUni = -1
End Function
Public Sub RemoveInUNI(ByRef ThreadId As Long, ByRef APIType As Long)
On Error GoTo Dalje
CheckUNI.Remove "X" & ThreadId & APIType
Exit Sub
Dalje:
On Error GoTo 0
End Sub
