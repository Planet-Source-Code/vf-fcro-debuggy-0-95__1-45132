Attribute VB_Name = "Module2"
Option Explicit
Public Unhandled As Byte

Public SkipEns As Byte 'skip Load Dll!

Public IsF11 As Boolean 'Is form11 visible


Public NameOfRunned As String

Public ISBPDisabled As Byte
Public IsLoadedProcess As Byte 'Provjera na Create Process
Public TerminateId As Long

Public GLOBALCOUNT As Long
Public GLOBALAFTERCOUNT As Long

Public NOTIFYJMPCALL As Byte 'JMP/CALL
Public NOTIFYVALG As Byte 'LEA,MOV,PUSH

Public DASM As New DisAsm

Public Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hWnd As Long
End Type


Public Type WNDCLASS
    Style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

Public Const CS_CLASSDC = &H40
Public Const CS_OWNDC = &H20
Public Const CS_GLOBALCLASS = &H4000
Public Const CS_HREDRAW = &H2
Public Const CS_PARENTDC = &H80
Public Const CS_VREDRAW = &H1


Const DIFFERENCE = 11
Const RT_ACCELERATOR = 9&
Const RT_ANICURSOR = (21)
Const RT_ANIICON = (22)
Const RT_BITMAP = 2&
Const RT_CURSOR = 1&
Const RT_DIALOG = 5&
Const RT_DLGINCLUDE = (17)
Const RT_FONT = 8&
Const RT_FONTDIR = 7&
Const RT_ICON = 3&
Const RT_GROUP_CURSOR = (RT_CURSOR + DIFFERENCE)
Const RT_GROUP_ICON = (RT_ICON + DIFFERENCE)
Const RT_HTML = (23)
Const RT_MENU = 4&
Const RT_MESSAGETABLE = (11)
Const RT_PLUGPLAY = (19)
Const RT_RCDATA = 10&
Const RT_STRING = 6&
Const RT_VERSION = (16)
Const RT_VXD = (20)

Public LastActiveIMP As Byte
Public LastActiveEXP As Byte

Public DataPW() As Byte '??? Data for Search

Public ActiveH As Long '??
Public ActiveLength As Long  '??
'Public ActiveMName As String '??

Public ActiveProcess As Long 'trenutni process koji se debugira!
Public ActiveStackPosition As Long
Public ActiveBasePosition As Long
Public DISCOUNT As Long 'trenutna adresa gdje smo skrolali.
Public Forward As Byte
Public LAST As Long
Public NextF As Byte
Public NextB As Byte

Public ActiveThread As Long 'izabrani thread
Public ShowCaption As String

Public CTX As CONTEXT
Public ProcessHandle As Long
Public SPYBREAKPOINTS As New Collection
Public ACTIVEBREAKPOINTS As New Collection
Public PROCESSESTHREADS As New Collection
Public ACTMODULESBYPROCESS As New Collection
'Public LASTTHREADEIP As New Collection
'Public RTRIGGER As New Collection
'Zamjena za RTRIGGER
Public TRIGGERADDRESS As Long
Public TRIGGERFLAG As Long


Declare Sub ArrayDescriptor Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc() As Any, ByVal ByteLen As Long)

Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal WMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
 Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
 Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
 Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

 Declare Function GetSystemMenu Lib "user32" _
(ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hWnd As Long, ByVal WMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long

'SREDIO
Public Sub RemoveBreakPoint(COL As Collection, ByVal Address As Long)
On Error GoTo Dalje
Dim BPData() As Long
Dim ORGBYTE As Byte
BPData = COL.Item("X" & Address)
ORGBYTE = CByte(BPData(1))
WriteProcessMemory ProcessHandle, ByVal Address, ORGBYTE, 1, ByVal 0&
FlushInstructionCache ProcessHandle, ByVal Address, 1
COL.Remove "X" & Address
Exit Sub
Dalje:
On Error GoTo 0
End Sub
'SREDIO
Public Sub AddBreakPoint(COL As Collection, ByVal Address As Long, ByVal SkipWrite As Byte)
On Error Resume Next
Dim ORGBYTE As Byte
Dim HARDBP As Byte
Dim BPData(1) As Long
HARDBP = &HCC
Dim PHandle As Long
COL.Remove "X" & Address
If Err <> 0 Then On Error GoTo 0
ReadProcessMemory ProcessHandle, ByVal Address, ORGBYTE, 1, ByVal 0&
If SkipWrite = 0 Then
WriteProcessMemory ProcessHandle, ByVal Address, HARDBP, 1, ByVal 0&
FlushInstructionCache ProcessHandle, ByVal Address, 1
End If
BPData(0) = Address
BPData(1) = CLng(ORGBYTE)
COL.Add BPData, "X" & Address
End Sub
'SREDIO

Public Function GetBreakPoint(COL As Collection, ByVal Address As Long, ByRef IsFounded As Byte) As Byte
On Error GoTo Dalje
Dim BPData() As Long
BPData = COL.Item("X" & Address)
GetBreakPoint = CByte(BPData(1))
IsFounded = 1
Exit Function
Dalje:
On Error GoTo 0
IsFounded = 0
End Function
Public Sub RestoreAllBreakPoints(COL As Collection)
Dim u As Long
Dim BPData() As Long
Dim HBP As Byte
HBP = &HCC
For u = 1 To COL.count
BPData = COL.Item(u)
WriteProcessMemory ProcessHandle, ByVal BPData(0), HBP, 1, ByVal 0&
Next u
End Sub
Public Sub RestoreAllOriginalBytes(COL As Collection)
Dim u As Long
Dim BPData() As Long
Dim ORGBYTE As Byte
For u = 1 To COL.count
BPData = COL.Item(u)
ORGBYTE = CByte(BPData(1))
WriteProcessMemory ProcessHandle, ByVal BPData(0), ORGBYTE, 1, ByVal 0&
Next u
End Sub
Public Sub RestoreOriginalBytes(COL As Collection, ByVal Address As Long)
Dim IsValidBP As Byte
Dim ORGBYTE As Byte
ORGBYTE = GetBreakPoint(COL, Address, IsValidBP)
If IsValidBP = 0 Then Exit Sub
WriteProcessMemory ProcessHandle, ByVal Address, ORGBYTE, 1, ByVal 0&
End Sub
Public Sub RestoreBreakPoint(COL As Collection, ByVal Address As Long)
Dim IsValidBP As Byte
Dim HBP As Byte
Call GetBreakPoint(COL, Address, IsValidBP)
If IsValidBP = 0 Then Exit Sub
HBP = &HCC
WriteProcessMemory ProcessHandle, ByVal Address, HBP, 1, ByVal 0&
End Sub

'SREDIO
Public Function GetModulePath(ByVal BaseAdr As Long) As String
On Error GoTo Dalje
Dim Dat() As String
Dat = ACTMODULESBYPROCESS("X" & BaseAdr)
GetModulePath = Dat(4)

Exit Function
Dalje:
On Error GoTo 0
End Function


Public Function FindInModules(ByVal ActAddress As Long, Optional ByRef BaseAddress As Long, Optional ByRef Length As Long) As String
On Error GoTo Dalje
Dim u As Long
For u = 1 To ACTMODULESBYPROCESS.count
Dim Dat() As String
Dat = ACTMODULESBYPROCESS.Item(u)
If ActAddress >= CLng(Dat(1)) And ActAddress <= AddBy8(CLng(Dat(1)), CLng(Dat(2))) Then FindInModules = Dat(0): BaseAddress = CLng(Dat(1)): Length = CLng(Dat(2)): Exit Function
Next u
Exit Function
Dalje:
On Error GoTo 0
End Function
'DOBRO
Public Sub EnumActiveModules()
On Error GoTo Kraj
'Dim PHandle As Long
Dim N() As Long
Dim ret As Long
Dim PCSLength As Long
ReDim N(999)
ret = EnumProcessModules(ProcessHandle, N(0), 1000, PCSLength)
ReDim Preserve N(PCSLength / 4 - 1)
Dim u As Long
Set ACTMODULESBYPROCESS = Nothing
Dim S(4) As String
Dim INF1 As LPMODULEINFO
Dim MnM As String
Dim Nlen As Long
For u = 0 To UBound(N)
MnM = Space(260)
Nlen = GetModuleFileNameExA(ProcessHandle, N(u), MnM, 260)
S(4) = MnM
MnM = Left(MnM, Nlen)

Call GetModuleInformation(ProcessHandle, N(u), INF1, Len(INF1))
S(0) = NameFromPath(MnM)
S(1) = CStr(INF1.lpBaseOfDll)
S(2) = CStr(INF1.SizeOfImage)
S(3) = CStr(INF1.EntryPoint)

ACTMODULESBYPROCESS.Add S, "X" & CStr(INF1.lpBaseOfDll) 'Stavi kao base adresu

AddInExportsSearch S(0), INF1.lpBaseOfDll
Next u
Exit Sub
Kraj:
On Error GoTo 0
End Sub
Public Sub AddInActiveModules(ByVal Address As Long, Optional MName As String)
On Error GoTo Dalje
Dim S(4) As String
Dim EPPnt As Long
ExPs.ModuleName = ""
ReadPE2 Address, LastActiveIMP, LastActiveEXP

ExPs.ModuleName = Replace(ExPs.ModuleName, vbTab, "")
MName = Replace(MName, vbTab, "")

If Len(MName) = 0 Then
S(0) = ExPs.ModuleName
S(4) = ExPs.ModuleName
Else
S(0) = NameFromPath(MName)
S(4) = MName
End If

S(1) = Address
S(2) = NTHEADER.OptionalHeader.SizeOfImage
If NTHEADER.OptionalHeader.AddressOfEntryPoint <> 0 Then
EPPnt = Address + NTHEADER.OptionalHeader.AddressOfEntryPoint
End If
S(3) = EPPnt
ACTMODULESBYPROCESS.Add S, "X" & Address
Exit Sub
Dalje:
On Error GoTo 0
End Sub
'SREDIO
Public Function RemoveInActiveModules(ByVal Address As Long) As String
On Error GoTo Dalje
Dim S() As String
S = ACTMODULESBYPROCESS.Item("X" & CStr(Address))
RemoveInActiveModules = S(0)
ACTMODULESBYPROCESS.Remove "X" & CStr(Address)
Exit Function
Dalje:
On Error GoTo 0
End Function



Public Sub ReadModules(ByRef LB As ListBox)
On Error GoTo Dalje
LB.Clear
Dim Dat() As String
Dim u As Long
For u = 1 To ACTMODULESBYPROCESS.count
Dat = ACTMODULESBYPROCESS.Item(u)
LB.AddItem Dat(0) & vbTab & Hex(Dat(1)) & vbTab & Hex(Dat(2)) & vbTab & Hex(Dat(3))
Next u
Exit Sub
Dalje:
On Error GoTo 0
End Sub

'SREDIO
'1-parametar ThreadId
'2-parametar Handle
Public Sub AddToThreadsList(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
On Error GoTo Dalje
Dim InTh(2) As Long
InTh(0) = ThreadId
InTh(1) = ThreadHandle
InTh(2) = 1
PROCESSESTHREADS.Add InTh, "X" & ThreadId
Exit Sub
Dalje:
On Error GoTo 0
End Sub
'SREDIO
Public Sub RemoveFromThreadList(ByVal ThreadId As Long)
On Error GoTo Dalje
Dim InTh() As Long
InTh = PROCESSESTHREADS.Item("X" & ThreadId)
CloseHandle InTh(1)
PROCESSESTHREADS.Remove "X" & ThreadId
Exit Sub
Dalje:
On Error GoTo 0
End Sub
Public Function GetHandleOfThread(ByVal ThreadId As Long, Optional ByRef IsRunning As Long) As Long
On Error GoTo Dalje
Dim InTh() As Long
InTh = PROCESSESTHREADS.Item("X" & ThreadId)
GetHandleOfThread = InTh(1)
IsRunning = InTh(2)
Exit Function
Dalje:
On Error GoTo 0
End Function
Public Function IsRunningThread(ByVal ThreadId As Long) As Long
On Error GoTo Dalje
Dim InTh() As Long
InTh = PROCESSESTHREADS.Item("X" & ThreadId)
IsRunningThread = InTh(2)
Exit Function
Dalje:
On Error GoTo 0
End Function
Public Sub ChangeStateThread(ByVal ThreadId As Long, ByVal StateTh As Long)
On Error GoTo Dalje
Dim InTh() As Long
InTh = PROCESSESTHREADS.Item("X" & ThreadId)
PROCESSESTHREADS.Remove "X" & ThreadId
InTh(2) = StateTh
PROCESSESTHREADS.Add InTh, "X" & ThreadId
Exit Sub
Dalje:
On Error GoTo 0
End Sub

'SREDIO
Public Sub ReadThreadsFromProcess(ByRef CB As ListBox)
CB.Clear
On Error GoTo Dalje
Dim i As String
Dim u As Long
Dim InTh() As Long
For u = 1 To PROCESSESTHREADS.count
InTh = PROCESSESTHREADS.Item(u)
If InTh(2) = 0 Then
i = "Suspend"
ElseIf InTh(2) = 1 Then
i = "Running"
ElseIf InTh(2) = 2 Then
i = "Waiting"
End If

If InTh(0) = MainPThread Then i = "(Main) " & i

CB.AddItem InTh(0) & "," & i
Next u
Exit Sub
Dalje:
On Error GoTo 0
End Sub


Public Function CanEnum() As Byte
On Error GoTo Dalje
Dim i As String
Dim u As Long
Dim InTh() As Long
For u = 1 To PROCESSESTHREADS.count
InTh = PROCESSESTHREADS.Item(u)
If InTh(2) = 2 Then CanEnum = 0: Exit Function
Next u
CanEnum = 1
Exit Function
Dalje:
On Error GoTo 0
CanEnum = 0
End Function


' upravljanje hardware-skim breakpointima
'Public Sub SetFlagInTrigger(ByVal ThreadId As Long, ByVal Address As Long, ByVal TRIGGERFLAG As Long)
'On Error GoTo Dalje
'Dim TrDt() As Long
'TrDt = RTRIGGER.Item("X" & ThreadId)
'TrDt(1) = TRIGGERFLAG
'RTRIGGER.Remove "X" & ThreadId
'RTRIGGER.Add TrDt, "X" & ThreadId
'Exit Sub
'Dalje:
'On Error GoTo 0
'End Sub
'Public Sub SetInTrigger(ByVal ThreadId As Long, ByVal Address As Long, Optional ByVal TRIGGERFLAG As Long)
'On Error Resume Next
'RTRIGGER.Remove "X" & ThreadId
'If Err <> 0 Then On Error GoTo 0
'Dim TrDt(1) As Long
'TrDt(0) = Address
'TrDt(1) = TRIGGERFLAG
'RTRIGGER.Add TrDt, "X" & ThreadId
'End Sub
'Public Sub RemoveInTrigger(ByVal ThreadId As Long)
'On Error GoTo Dalje
'RTRIGGER.Remove "X" & ThreadId
'Exit Sub
'Dalje:
'On Error GoTo 0
'End Sub
'Public Function GetFromTrigger(ByVal ThreadId As Long, ByRef IsValid As Byte, ByRef TRIGGERFLAG As Long) As Long
'On Error GoTo Dalje
'Dim TrDt() As Long
'TrDt = RTRIGGER.Item("X" & ThreadId)
'GetFromTrigger = TrDt(0)
'TRIGGERFLAG = TrDt(1)
'IsValid = 1
'Exit Function
'Dalje:
'On Error GoTo 0
'End Function
Public Sub RegClass(ByVal Classname As String, ByVal hInstance As Long, ByVal AddressProc As Long)
Dim CLASSX As WNDCLASS
CLASSX.Style = CS_GLOBALCLASS
CLASSX.lpfnwndproc = AddressProc
CLASSX.hInstance = hInstance
CLASSX.lpszClassName = Classname
Call RegisterClass(CLASSX)
End Sub
Public Sub AddLine(ByRef StrX As String, ByRef RTB As RichTextBox, Optional ByRef ColorX As Long)
If ColorX = 0 Then
RTB.SelColor = &HE7DFD6
Else
RTB.SelColor = ColorX
End If
RTB.SelStart = Len(RTB.Text)
RTB.SelLength = 0
RTB.SelText = StrX & vbCrLf
End Sub

'Communicator Messages Loop
Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Connection Beetween Debugger Thread And our Project!
'The Best way for a harmless conversation with async thread out of our VB thread!


Select Case uMsg

Case 901
MsgBox "Cannot attach on Process:" & wParam & " ,I'm out...!", vbCritical, "Information"
StopAndClear: Unload Form16


Case 902
Form16.Show
Form4.Enabled = False
Form4.Visible = False
DoEvents

Case 910
'If UseTrace = 1 Then
'Form28.ProcessException wParam
'Else
ProcessException wParam
'End If


Case 920 'Hooks Windows (Creating/Destroying Windows)

ProcessHooks wParam, lParam

Case 921 'Peek Message in Debugged process (This debugger process only WM_COMMAND)
'WH_CALLWNDPROC
ProcessMSG wParam, lParam, 0 'wparam=1 , lParam=PTR on struct


Case 922
'WH_CALLWNDPROCRET ***currently under progress
'ProcessMSG wParam, lParam, 1


Case 931
'Window proc -CHECKED FROM REMOTE THREAD CREATED BY THIS DEBUGGER!!!!!
Form14.Label8 = "Window Proc At Address:" & Hex(lParam)


Case 938
Dim VRes(4) As Long
Call ReadProcessMemory(ProcessHandle, ByVal lParam, VRes(0), 20, ByVal 0&)
'ResRC (ResCounter)
ResRC(ResCounter).ResType = GetStringFromPointer(VRes(0))
ResRC(ResCounter).ResName = GetStringFromPointer(VRes(1))
ResRC(ResCounter).LangId = VRes(2)
ResRC(ResCounter).ResAddress = VRes(3)
ResRC(ResCounter).ResLength = VRes(4)
ResCounter = ResCounter + 1


Case 939
'End of enumeration!
If ResCounter = 0 Then
Unload Form1
MsgBox "Module doesn't contain resources", vbInformation, "Information"
Else
ReDim Preserve ResRC(ResCounter - 1)
Form1.ShowIt
End If


Case 940
DataArea = lParam


Case 850 'Single Step
If UseTrace = 1 Then
Form28.ProcessException lParam
Else
ProcessException wParam
End If

Case 851 'Breakpoint
If UseTrace = 1 Then
Form28.ProcessException lParam
Else
ProcessException wParam
End If



End Select


WndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function
Public Sub ProcessMSG(ByVal Param1 As Long, ByVal Param2 As Long, ByVal Param3 As Byte)
Dim CheckThreadId As Long
Dim CheckProcessId As Long

'Process Hooks for WH_CALLWNDPROC
'Message Chain designed by Vanja Fuckar @2002




Dim MSGS As CWPSTRUCT
Dim ret As Long
ret = ReadProcessMemory(ProcessHandle, ByVal Param2, MSGS, Len(MSGS), ByVal 0&)
If ret = 0 Then Exit Sub

Dim IsValidBRK As Byte
Dim WndSd() As Long
'0--HWND
'1--BRKMSG
'2--not used yet
If MSGS.message = WM_COMMAND And ConfigData(6) <> 0 Then
Dim NMSG As Long
Dim CTID As Long
NMSG = GetHI(MSGS.wParam)
CTID = GetLO(MSGS.wParam)


WndSd = GetBreakWND(BRKWMCMD, MSGS.lParam, NMSG, 0, IsValidBRK)
If IsValidBRK = 1 Then
CheckThreadId = GetWindowThreadProcessId(MSGS.lParam, CheckProcessId)
PlayS
ActiveThread = CheckThreadId
Form16.InSuspend
AddLine "WM_COMMAND Breakpoint on value:" & NMSG & ",Class Name:" & ClassNameEx(MSGS.lParam) & ",Hwnd:" & Hex(MSGS.lParam) _
& ",On Parent Class Name:" & ClassNameEx(MSGS.hWnd) & ",Hwnd:" & Hex(MSGS.hWnd), Form12.rt1
If IsF11 Then Unload Form11
End If


Else
'Others WM_
If ConfigData(6) <> 0 Then

WndSd = GetBreakWND(BRKW, MSGS.hWnd, MSGS.message, 0, IsValidBRK)
If IsValidBRK = 1 Then
CheckThreadId = GetWindowThreadProcessId(MSGS.hWnd, CheckProcessId)
PlayS
ActiveThread = CheckThreadId
If IsF11 Then Unload Form11
Form16.InSuspend
AddLine "WM_Value:" & Hex(MSGS.message) & " Breakpoint,Class Name:" & ClassNameEx(MSGS.hWnd) & ",Hwnd:" & Hex(MSGS.hWnd), Form12.rt1
End If
End If
End If

'Menu Click
If ConfigData(8) <> 0 Then
If MSGS.message = &H1ED Then
If StrComp(ClassNameEx(MSGS.hWnd), "#32768") = 0 Then
CheckThreadId = GetWindowThreadProcessId(MSGS.hWnd, CheckProcessId)
PlayS
ActiveThread = CheckThreadId

'Change to ordinary non topmost window
SetWindowPos MSGS.hWnd, -2, 0, 0, 0, 0, &H53&

If IsF11 Then Unload Form11
Form16.InSuspend
AddLine "Menu Click Id:" & Hex(MSGS.wParam), Form12.rt1
End If
End If
End If



'destroying windows
If MSGS.message = WM_DESTROY Then

CheckThreadId = GetWindowThreadProcessId(MSGS.hWnd, CheckProcessId)
AddLine "Destroy Window:" & Hex(MSGS.hWnd) & ",Class Name:" & ClassNameEx(MSGS.hWnd) & ",In Thread:" & CheckThreadId, Form12.rt1, &HFFFF&
RemoveProp MSGS.hWnd, "GOFORDEBUG" 'Remove property from window
RemoveWins MSGS.hWnd
RemoveEntireWND BRKW, MSGS.hWnd
RemoveEntireWND BRKWMCMD, MSGS.hWnd

If ConfigData(5) = 1 Then
ActiveThread = CheckThreadId
Form16.InSuspend
PlayS
End If


End If


End Sub

Public Function ClassNameEx(ByVal hWnd As Long) As String
Dim ClLen As Long
ClassNameEx = Space(260)
ClLen = GetClassName(hWnd, ClassNameEx, 260)
ClassNameEx = Left(ClassNameEx, ClLen)
End Function

Public Sub ProcessHooks(ByVal iMSG As Long, ByVal wParam As Long)
Dim CheckThreadId As Long
Dim CheckProcessId As Long

CheckThreadId = GetWindowThreadProcessId(wParam, CheckProcessId)
If CheckProcessId <> ActiveProcess Or ActiveProcess = 0 Then Exit Sub
Dim ClassNm As String

ClassNm = ClassNameEx(wParam)

'Window Hook message chain.
'Designed by Vanja Fuckar @2002
Select Case iMSG

Case HCBT_CREATEWND
AddLine "Create Window:" & Hex(wParam) & ",Class Name:" & ClassNm & ",In Thread:" & CheckThreadId, Form12.rt1, &HFFFF&
AddWins ClassNm, wParam, CheckThreadId
Call SetProp(wParam, "GOFORDEBUG", 1) 'Insert notify Flag for CallBack!
'If we don't do this,we will receive all messages from all of Desktop Shell Windows and Child Windows!
'surely that will do overload our project,yeeak!

If ConfigData(4) = 1 Then
ActiveThread = CheckThreadId
Form16.InSuspend
PlayS
End If


End Select

End Sub

Public Function ResolveExpName(ByVal EValue As Long) As String

Select Case EValue

Case EXCEPTION_DATATYPE_MISALIGNMENT
ResolveExpName = "EXCEPTION DATATYPE MISALIGNMENT"
Unhandled = 1

Case EXCEPTION_ACCESS_VIOLATION
ResolveExpName = "EXCEPTION ACCESS VIOLATION"
Unhandled = 1

Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED
ResolveExpName = "EXCEPTION ARRAY BOUNDSEXCEEDED"
Unhandled = 1

Case EXCEPTION_FLT_DIVIDE_BY_ZERO
ResolveExpName = "EXCEPTION FLOAT DIVIDE BY ZERO"
Unhandled = 1

Case EXCEPTION_FLT_INVALID_OPERATION
ResolveExpName = "EXCEPTION FLOAT INVALID OPERATION"
Unhandled = 1

Case EXCEPTION_FLT_OVERFLOW
ResolveExpName = "EXCEPTION FLOAT OVERFLOW"
Unhandled = 1

Case EXCEPTION_FLT_INEXACT_RESULT
ResolveExpName = "EXCEPTION INEXACT RESULT"
Unhandled = 1

Case EXCEPTION_INT_DIVIDE_BY_ZERO
ResolveExpName = "EXCEPTION INTEGER DIVIDE BY ZERO"
Unhandled = 1

Case EXCEPTION_INT_OVERFLOW
ResolveExpName = "EXCEPTION INTEGER OVERFLOW"
Unhandled = 1

Case EXCEPTION_ILLEGAL_INSTRUCTION
ResolveExpName = "EXCEPTION ILLEGAL INSTRUCTION"
Unhandled = 1

Case EXCEPTION_PRIV_INSTRUCTION
ResolveExpName = "EXCEPTION PRIVILEGED INSTRUCTION"
Unhandled = 1

Case Else
Unhandled = 0
End Select
End Function




Public Sub ProcessException(ByVal Param1 As Long)
'Param1=PTR on EVENT STRUCTURE!
Dim BrkRequest As Byte
Dim SHTC As Byte
Dim Vent() As String
'Dim PROCRETT As Long
Dim RETCONFIRMATION As Long
Dim TContext As CONTEXT
Dim MInf As LPMODULEINFO
Dim ModuleN As String
Dim TempThreadH As Long
Dim TempProcessH As Long
Dim IsValidBP As Byte

CopyMemory DBGEVENT, ByVal Param1, &H50&

Select Case DBGEVENT.dwDebugEventCode
Case EXCEPTION_DEBUG_EVENT
CopyMemory EXCEPTIONINFO, DBGEVENT.Data(0), Len(EXCEPTIONINFO)

'--------
Select Case EXCEPTIONINFO.ExceptionRecord.ExceptionCode

     
    Case EXCEPTION_BREAKPOINT

    GetBreakPoint SPYBREAKPOINTS, EXCEPTIONINFO.ExceptionRecord.ExceptionAddress, IsValidBP
    If IsValidBP = 1 Then
    CTX = GetContext(DBGEVENT.dwThreadId)
    CTX.Eip = CTX.Eip - 1
    SetContext DBGEVENT.dwThreadId, CTX
    
    
'Maknuli*******RADIMO
    If GetInRet(EXCEPTIONINFO.ExceptionRecord.ExceptionAddress, Vent) = 1 Then
    RemoveInRet EXCEPTIONINFO.ExceptionRecord.ExceptionAddress
    RemoveBreakPoint SPYBREAKPOINTS, EXCEPTIONINFO.ExceptionRecord.ExceptionAddress
    TRIGGERADDRESS = 0: TRIGGERFLAG = -1
    TransRet EXCEPTIONINFO.ExceptionRecord.ExceptionAddress, CTX.Eax, Vent, BrkRequest, DBGEVENT.dwThreadId
    
    If BrkRequest = 1 Then
    ActiveThread = DBGEVENT.dwThreadId
    GoTo OutMi
    End If
    
    
    Else
    
    TRIGGERADDRESS = EXCEPTIONINFO.ExceptionRecord.ExceptionAddress
    TRIGGERFLAG = 1
    RestoreOriginalBytes SPYBREAKPOINTS, CTX.Eip
    SetSingleStep DBGEVENT.dwThreadId
    TranslateSpy EXCEPTIONINFO.ExceptionRecord.ExceptionAddress, CTX, DBGEVENT.dwThreadId, RETCONFIRMATION
    End If
    
    If RETCONFIRMATION = 1 Then
    ActiveThread = DBGEVENT.dwThreadId
    GoTo OutMi
    End If
    
    
    ContinueDebug
    Exit Sub
    End If
    
    
    'Check if is our HARD BREAKPOINT!?
    GetBreakPoint ACTIVEBREAKPOINTS, EXCEPTIONINFO.ExceptionRecord.ExceptionAddress, IsValidBP
    If IsValidBP = 1 Then
    CTX = GetContext(DBGEVENT.dwThreadId)
    CTX.Eip = CTX.Eip - 1
    SetContext DBGEVENT.dwThreadId, CTX
    TRIGGERADDRESS = EXCEPTIONINFO.ExceptionRecord.ExceptionAddress
    TRIGGERFLAG = 0
    
    RestoreOriginalBytes ACTIVEBREAKPOINTS, CTX.Eip
    End If
    AddLine "Breakpoint encounted At:" & Hex(EXCEPTIONINFO.ExceptionRecord.ExceptionAddress) & " ,In Thread:" & DBGEVENT.dwThreadId, Form12.rt1
    

 'Maknuli****RADIMO
    If GetInRet(EXCEPTIONINFO.ExceptionRecord.ExceptionAddress, Vent) = 1 Then
    RemoveInRet EXCEPTIONINFO.ExceptionRecord.ExceptionAddress
    TransRet EXCEPTIONINFO.ExceptionRecord.ExceptionAddress, CTX.Eax, Vent, BrkRequest, DBGEVENT.dwThreadId

    If BrkRequest = 1 Then
    ActiveThread = DBGEVENT.dwThreadId
    GoTo OutMi
    End If

    End If
    
    
 
    Case EXCEPTION_SINGLE_STEP
    
    'Dodao!
    GetBreakPoint SPYBREAKPOINTS, EXCEPTIONINFO.ExceptionRecord.ExceptionAddress, IsValidBP
    If IsValidBP = 1 Then
    CTX = GetContext(DBGEVENT.dwThreadId)
    TranslateSpy EXCEPTIONINFO.ExceptionRecord.ExceptionAddress, CTX, DBGEVENT.dwThreadId, RETCONFIRMATION
    
    If RETCONFIRMATION = 1 Then
    ActiveThread = DBGEVENT.dwThreadId
    GoTo OutMi
    End If

'Maknuli****RADIMO**********
   ElseIf GetInRet(EXCEPTIONINFO.ExceptionRecord.ExceptionAddress, Vent) = 1 Then
    RemoveInRet EXCEPTIONINFO.ExceptionRecord.ExceptionAddress
'    Ako je taj BRK napravljen od usera!
    TransRet EXCEPTIONINFO.ExceptionRecord.ExceptionAddress, CTX.Eax, Vent, BrkRequest, DBGEVENT.dwThreadId
    
    If BrkRequest = 1 Then
    ActiveThread = DBGEVENT.dwThreadId
    GoTo OutMi
    End If
    
    End If
    
     

    If TRIGGERFLAG <> -1 Then
    RestoreBreakPoint SPYBREAKPOINTS, TRIGGERADDRESS
    RestoreBreakPoint ACTIVEBREAKPOINTS, TRIGGERADDRESS
    TRIGGERADDRESS = 0
    
    If TRIGGERFLAG = 1 Then
    'Ukoliko nije na single stepu nastavi
    TRIGGERFLAG = -1
    ContinueDebug
    Exit Sub
    End If
    TRIGGERFLAG = -1
    End If
           
 
    Case Else
 '   Dim Ladd As String
    If EXCEPTIONINFO.ExceptionRecord.ExceptionFlags <> 1 And EXCEPTIONINFO.dwFirstChance <> EXCEPTIONINFO.ExceptionRecord.ExceptionCode Then
  '   Ladd = "Continuable Exception"

      Else
     ' Ladd = "Continuable Exception"
    Unhandled = 0
    ContinueDebugNotHandle
    Exit Sub
    End If
    MsgBox "Exception Number " & Hex(EXCEPTIONINFO.ExceptionRecord.ExceptionCode) & " At Address:" & Hex(EXCEPTIONINFO.ExceptionRecord.ExceptionAddress) _
    & vbCrLf & ResolveExpName(EXCEPTIONINFO.ExceptionRecord.ExceptionCode), vbCritical, "Error"
    'Unhandled = 1
    Unload Form29
    Unload Form28
    
    
End Select

Case CREATE_THREAD_DEBUG_EVENT
CopyMemory CREATETHREADINFO, DBGEVENT.Data(0), Len(CREATETHREADINFO)



If AccThreadX = DBGEVENT.dwThreadId Then ContinueDebug: Exit Sub


If UseTrace = 1 Then
AddLine "Create Thread:" & DBGEVENT.dwThreadId, Form28.rt1
End If

AddToThreadsList DBGEVENT.dwThreadId, CREATETHREADINFO.hThread
ReadThreadsFromProcess Form16.List1

AddLine "Create Thread:" & DBGEVENT.dwThreadId, Form12.rt1
AddLine "Thread Local Base At:" & Hex(CREATETHREADINFO.lpThreadLocalBase), Form12.rt1

SHTC = 2

If ConfigData(2) = 0 Then CTX = GetContext(DBGEVENT.dwThreadId): ShowThreadInfo SHTC: ContinueDebug: Exit Sub




Case CREATE_PROCESS_DEBUG_EVENT
CopyMemory CREATEPROCESSINFO, DBGEVENT.Data(0), Len(CREATEPROCESSINFO)

If IsLoadedProcess = 1 Then
'Za sada single process debugger
MsgBox "Atempt to Create Another Process!" & vbCrLf & _
"This is a single Process Debugger.New Process will be terminated!", vbCritical, "Info"
TerminateProcess CREATEPROCESSINFO.hProcess, 0
TerminateId = DBGEVENT.dwProcessId
IsLoadedProcess = 2
AddLine "Create Process:" & DBGEVENT.dwProcessId & ",Thread:" & DBGEVENT.dwThreadId & " ***Terminated By This Debugger***", Form12.rt1
ContinueDebug
Exit Sub
End If


'Call GetModuleInformation(CREATEPROCESSINFO.hProcess, CREATEPROCESSINFO.lpBaseOfImage, MInf, Len(MInf))
ProcessHandle = CREATEPROCESSINFO.hProcess
ActiveProcess = DBGEVENT.dwProcessId
MainPThread = DBGEVENT.dwThreadId
EnumActiveModules
AddToThreadsList DBGEVENT.dwThreadId, CREATEPROCESSINFO.hThread
ReadThreadsFromProcess Form16.List1
Form16.Caption = "Disassembling Process:" & ActiveProcess
AddLine "Create Process:" & DBGEVENT.dwProcessId & ",Thread:" & DBGEVENT.dwThreadId, Form12.rt1
AddLine "Thread Local Base At:" & Hex(CREATEPROCESSINFO.lpThreadLocalBase), Form12.rt1


IsLoadedProcess = 1
SHTC = 1

If Len(NameOfRunned) = 0 Then EnumWindows AddressOf EnumW, 0

'Dodan glavni modul...
'*********************

If Len(FindInModules(CREATEPROCESSINFO.lpBaseOfImage)) = 0 Then
'Ako nema tada
Dim S(4) As String
S(0) = NameFromPath(NameOfRunned)
AddInExportsSearch S(0), CREATEPROCESSINFO.lpBaseOfImage
S(1) = CREATEPROCESSINFO.lpBaseOfImage
S(2) = NTHEADER.OptionalHeader.SizeOfImage
S(3) = CREATEPROCESSINFO.lpBaseOfImage + NTHEADER.OptionalHeader.AddressOfEntryPoint
S(4) = NameOfRunned
ACTMODULESBYPROCESS.Add S, "X" & CREATEPROCESSINFO.lpBaseOfImage

End If




Case EXIT_THREAD_DEBUG_EVENT
CopyMemory EXITTHREADINFO, DBGEVENT.Data(0), Len(EXITTHREADINFO)

If AccThreadX = DBGEVENT.dwThreadId Then ContinueDebug: Exit Sub

If UseTrace = 1 And DBGEVENT.dwThreadId = Form28.ContThread Then
MsgBox "Traced Thread Terminated!", vbInformation, "Information!"
Unload Form29: Unload Form28
ElseIf UseTrace = 1 Then
AddLine "Terminate Thread:" & DBGEVENT.dwThreadId, Form28.rt1
End If



ShowDatas DBGEVENT.dwProcessId, DBGEVENT.dwThreadId, CTX, 0

ActiveThread = 0
Form16.Label2 = "Selected:"
RemoveFromThreadList DBGEVENT.dwThreadId
ReadThreadsFromProcess Form16.List1
AddLine "Exit Thread:" & DBGEVENT.dwThreadId, Form12.rt1
ContinueDebug




Exit Sub



Case EXIT_PROCESS_DEBUG_EVENT
If IsLoadedProcess = 2 Then
'Ako je process pokrenut od main process-a
AddLine "Exit Process:" & DBGEVENT.dwProcessId & " ***Terminated By This Debugger***", Form12.rt1
IsLoadedProcess = 1: ContinueDebug: Exit Sub
Else
StopAndClear
Unload Form16: Exit Sub
End If


Case LOAD_DLL_DEBUG_EVENT
CopyMemory LOADDLL, DBGEVENT.Data(0), Len(LOADDLL)

'Preskoci ako je terminirajuci Process
If TerminateId = DBGEVENT.dwProcessId Then
AddLine "Loading ntdll.dll in process under the termination!", Form12.rt1
TerminateId = 0: ContinueDebug: Exit Sub
End If

If SkipEns = 1 Then GoTo SkipEE

Dim APP As Long
Call ReadProcessMemory(ProcessHandle, ByVal LOADDLL.lpImageName, APP, 4, ByVal 0&)
Dim Ostring As String
If LOADDLL.fUnicode = 1 Then
Ostring = GetStringFPTR(APP, 1024, 1)
Else
Ostring = GetStringFPTR(APP, 1024)
End If

AddInActiveModules LOADDLL.lpBaseOfDll, Ostring
AddInExportsSearch Ostring, LOADDLL.lpBaseOfDll


SkipEE:
If Len(Ostring) = 0 Then
AddLine "Load Dll:" & ExPs.ModuleName, Form12.rt1
Else
AddLine "Load Dll:" & Ostring, Form12.rt1
End If

'End If

If UseTrace = 1 And DBGEVENT.dwThreadId = Form28.ContThread Then

If Len(Ostring) = 0 Then
AddLine "Load Dll:" & ExPs.ModuleName, Form28.rt1
Else
AddLine "Load Dll:" & Ostring, Form28.rt1
End If


SetSingleStep DBGEVENT.dwThreadId: ContinueDebug: Exit Sub
End If


If ConfigData(0) = 0 Then ContinueDebug: Exit Sub

Case UNLOAD_DLL_DEBUG_EVENT
CopyMemory UNLOADDLL, DBGEVENT.Data(0), Len(UNLOADDLL)
ModuleN = RemoveInActiveModules(UNLOADDLL.lpBaseOfDll)
DeleteInExportsSearch ModuleN
AddLine "Unload Dll:" & ModuleN, Form12.rt1

If UseTrace = 1 And DBGEVENT.dwThreadId = Form28.ContThread Then
AddLine "UnLoad Dll:" & ExPs.ModuleName, Form28.rt1
SetSingleStep DBGEVENT.dwThreadId: ContinueDebug: Exit Sub
End If

If ConfigData(1) = 0 Then ContinueDebug: Exit Sub


Case Else 'DBG_DIVOVERFLOW
ContinueDebug
Exit Sub

End Select


OutMi:

ShowDatas DBGEVENT.dwProcessId, DBGEVENT.dwThreadId, CTX, SHTC

End Sub
Public Sub ShowDatas(ByVal ProcessId As Long, ByVal ThreadId As Long, ByRef ContX As CONTEXT, ByVal ShowInfoTC As Byte)
ContX = GetContext(ThreadId)
'AddLastEip ThreadId, ContX.Eip
ReadMem ProcessId, ContX.Eip
ActiveThread = ThreadId

PlayS
ChangeStateThread ThreadId, 2
ReadThreadsFromProcess Form16.List1

TouchIt ThreadId
Form16.Label2 = "Selected: " & ThreadId

ShowThreadInfo ShowInfoTC


'dodao ovdje
If Form8.Visible = True Then
PrintDump Form8.TextX, ActiveMemPos
If gBegAdr <> 0 And gLenAdr <> 0 Then
GetDataFromMem gBegAdr, DataPW, gLenAdr
End If
MEMINF = QueryMem(ActiveMemPos, MEMStr)
Form8.Text4 = MEMStr
End If


End Sub

Public Sub ShowThreadInfo(ByVal ShowInfoTC As Byte)
If Len(NameOfRunned) <> 0 Then
If ShowInfoTC = 1 Then
AddLine "Process Start At Address:" & Hex(CTX.Eax), Form12.rt1
ElseIf ShowInfoTC = 2 Then
AddLine "Thread Start At Address:" & Hex(CTX.Eax), Form12.rt1
End If
End If
End Sub

Public Function GetModuleNameFromHandle(ByVal hModule As Long) As String
Dim lLen As Long
GetModuleNameFromHandle = Space(260)
lLen = GetModuleFileNameExA(ProcessHandle, hModule, GetModuleNameFromHandle, 260)
GetModuleNameFromHandle = Left(GetModuleNameFromHandle, lLen)
End Function

Public Sub ReadMem(ByVal ProcessId As Long, ByVal Address As Long)
Dim TName As String
TName = FindInModules(Address, ActiveH, ActiveLength)
If Len(TName) = 0 Then TName = "Unknown Address Or Not Valid"
Form16.Label7 = "Module:" & TName
DISCOUNT = Address
AddForward Form16.rt1, Form16.rt2, 25, ProcessId, Form16.List8
End Sub




Public Function GetContext(ByVal ThreadId As Long) As CONTEXT
Dim ThHandle As Long
ThHandle = GetHandleOfThread(ThreadId)
GetContext.ContextFlags = CONTEXT_i486 Or CONTEXT_CONTROL Or CONTEXT_INTEGER Or CONTEXT_SEGMENTS Or CONTEXT_FLOATING_POINT
GetThreadContext ThHandle, GetContext
End Function
Public Sub SetContext(ByVal ThreadId As Long, NewContext As CONTEXT)
Dim ThHandle As Long
ThHandle = GetHandleOfThread(ThreadId)
SetThreadContext ThHandle, NewContext
End Sub
Public Sub SetSingleStep(ByVal ThreadId As Long)
CTX = GetContext(ThreadId)
CTX.EFlags = CTX.EFlags Or &H100&
SetContext ThreadId, CTX
End Sub
Public Sub ClearSingleStep(ByVal ThreadId As Long)
CTX = GetContext(ThreadId)
If (CTX.EFlags And &H100&) = &H100& Then CTX.EFlags = CTX.EFlags Xor &H100&
SetContext ThreadId, CTX
End Sub

Public Sub GetDataFromMem(ByVal Address As Long, ByRef DataX() As Byte, ByVal NumOfBytes As Long, Optional ByRef IsValidMem As Long)
ReDim DataX(NumOfBytes - 1)
IsValidMem = ReadProcessMemory(ProcessHandle, ByVal Address, DataX(0), NumOfBytes, ByVal 0&)
End Sub
Public Sub AddForward25(ByRef RTB As Variant, ByRef RTB2 As Variant, ByVal Nums As Long, ByVal ProcessId As Long, LB2 As ListBox)

Dim u As Long
Dim ORGBYTE As Byte
Dim IsValidBP As Byte
Dim IsError As Byte
Dim DTX() As Byte
Dim CMDS As String
LAST = DISCOUNT

For u = 0 To Nums
GetDataFromMem LAST, DTX, 16
DASM.BaseAddress = LAST

ORGBYTE = GetBreakPoint(SPYBREAKPOINTS, LAST, IsValidBP)
If IsValidBP = 1 Then DTX(0) = ORGBYTE: GoTo InLea


If IsValidBP = 0 Then
ORGBYTE = GetBreakPoint(ACTIVEBREAKPOINTS, LAST, IsValidBP)
End If

If IsValidBP <> 0 Then
DTX(0) = ORGBYTE
End If

InLea:

CMDS = DASM.DisAssemble(DTX, 0, Forward, 0, 0, IsError)
LAST = LAST + Forward
Next u
DISCOUNT = LAST
NextB = 0


AddForward RTB, RTB2, Nums, ProcessId, LB2
End Sub

Public Sub FillLabelModule(ByRef Address As Long)
Dim TName As String
Dim ICtc As String
TName = FindInModules(Address)
If Len(TName) = 0 Then TName = "Unknown Address Or Not Valid"
If TName = ValidCRef Then ICtc = "(Cached) "
Form16.Label7 = ICtc & "Module:" & TName
End Sub


Public Sub AddBackward25(ByRef RTB As Variant, ByRef RTB2 As Variant, ByVal Nums As Long, ByVal ProcessId As Long, LB2 As ListBox)
Dim u As Long
For u = 0 To Nums - 1
AddBackward RTB, RTB2, Nums, ProcessId, LB2, 1
Next u

AddBackward RTB, RTB2, Nums, ProcessId, LB2
End Sub
Public Sub AddForward(ByRef RTB As Variant, ByRef RTB2 As Variant, ByVal Nums As Long, ByVal ProcessId As Long, LB2 As ListBox)
FillLabelModule DISCOUNT

Dim DTX() As Byte
Dim CMDS As String
Dim AREF As String
Dim Allpr As String
Dim CHKC As String
Dim IsError As Byte
Dim CRef As String
Dim ExpSt As String
Dim IsValidBP As Byte
Dim BinBin As String
Dim ORGBYTE As Byte
Dim GFI As String
Dim IsString As Long
Dim i As Long
Dim u As Long
LAST = DISCOUNT

For u = 0 To Nums

GetDataFromMem LAST, DTX, 16
DASM.BaseAddress = LAST
ORGBYTE = GetBreakPoint(SPYBREAKPOINTS, LAST, IsValidBP)
If IsValidBP = 1 Then DTX(0) = ORGBYTE: GoTo InLea


If IsValidBP = 0 Then
ORGBYTE = GetBreakPoint(ACTIVEBREAKPOINTS, LAST, IsValidBP)
End If

If IsValidBP = 0 Then
InLea:
LB2.List(u) = ""
Else
LB2.List(u) = "*BP*"
DTX(0) = ORGBYTE
End If



CMDS = DASM.DisAssemble(DTX, 0, Forward, 0, 0, IsError)

NotifyData1(u) = NOTIFYJMPCALL
NotifyData2(u) = VALUES1
NotifyData3(u) = VALUES2
NotifyData4(u) = VALUES3
BinBin = ""
RWDump DTX, 0, Forward, BinBin

RTB(u).ToolTipText = "OFFSET: " & BinBin

If NOTIFYVALG = 1 Then
AREF = IsStringOnAdr(IsString)
If IsString = 1 Then AREF = "(Possible) String: " & AREF
End If

ExpSt = GetFromExportsSearch(FindInModules(LAST), LAST)
If Len(ExpSt) <> 0 Then ExpSt = "Export:" & ExpSt


GFI = GetFromIndex(INDEXESR, REFSR, LAST)
If Len(GFI) = 0 Then
GFI = GetFromIndex(EINDEXESR, EREFSR, LAST)
End If


CHKC = CheckCALL(NotifyData2(u), 1)
Allpr = ExpSt



If Len(Allpr) <> 0 And Len(CHKC) <> 0 Then
Allpr = Allpr & " ;" & CHKC
ElseIf Len(CHKC) <> 0 Then
Allpr = CHKC
End If

If Len(Allpr) <> 0 And Len(AREF) <> 0 Then
If Len(CHKC) = 0 Then
Allpr = Allpr & " ;" & AREF
End If
ElseIf Len(AREF) <> 0 Then
Allpr = AREF
End If


If Len(Allpr) <> 0 And Len(GFI) <> 0 Then
Allpr = Allpr & " ;" & GFI
ElseIf Len(GFI) <> 0 Then
Allpr = GFI
End If

RTB(u) = Hex(LAST) & vbTab & CMDS
RTB2(u) = Allpr
'RTB2(u) = ExpSt & AREF & CheckCALL(NotifyData2(u)) & GFI

If u = 0 Then
NextB = Forward
End If
LAST = LAST + Forward
AREF = ""
Next u
NextF = Forward

End Sub
Public Sub AddBackward(ByRef RTB As Variant, ByRef RTB2 As Variant, ByVal Nums As Long, ByVal ProcessId As Long, LB2 As ListBox, Optional ByVal SkipPres As Byte)
Dim IsError As Byte
Dim DTX() As Byte
GetDataFromMem DISCOUNT - 49, DTX, 50
RestoreOrgBytes DTX, DISCOUNT - 49, ProcessId
DASM.DisassembleBack DTX, 49, Forward, IsError
DISCOUNT = DISCOUNT - Forward
NextB = 0
NextF = 0
If SkipPres = 0 Then
AddForward Form16.rt1, Form16.rt2, Nums, ProcessId, LB2
End If
End Sub
Public Sub RestoreOrgBytes(Data() As Byte, ByVal StartAdr As Long, ByVal ProcessId As Long)
Dim u As Long
Dim IsValidBP As Byte
Dim ORGBYTE As Byte
For u = 0 To UBound(Data)
ORGBYTE = GetBreakPoint(SPYBREAKPOINTS, StartAdr + u, IsValidBP)
If IsValidBP = 0 Then
ORGBYTE = GetBreakPoint(ACTIVEBREAKPOINTS, StartAdr + u, IsValidBP)
End If
If IsValidBP = 1 Then
Data(u) = ORGBYTE
End If
IsValidBP = 0
Next u
End Sub



'Public Function GetFromProcess(ByRef COL As Collection, ByVal ProcessId As Long, ByRef IsValid As Byte) As Long
'On Error GoTo Dalje
'GetFromProcess = COL.Item("X" & ProcessId)
'IsValid = 1
'Exit Function
'Dalje:
'On Error GoTo 0
'End Function


'Public Sub AddInProcess(ByRef COL As Collection, ByVal ProcessId As Long)
'On Error GoTo Dalje
'Dim ProcH As Long
'ProcH = OpenProcess(PROCESS_ALL_ACCESS, 0, ProcessId)
'Dim VHandle As Long
'VHandle = VirtualAllocEx(ProcH, ByVal 0&, ByVal 100&, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
'COL.Add VHandle, "X" & ProcessId
'Exit Sub
'Dalje:
'On Error GoTo 0
'End Sub
'Public Sub RemoveFromProcess(ByRef COL As Collection, ByVal ProcessId As Long)
'On Error GoTo Dalje
'Dim VHandle As Long
'VHandle = COL.Item("X" & ProcessId)
'Dim ProcH As Long
'ProcH = OpenProcess(PROCESS_ALL_ACCESS, 0, ProcessId)
'Call VirtualFreeEx(ProcH, ByVal VHandle, ByVal 100&, MEM_DECOMMIT)
'CloseHandle ProcH
'COL.Remove "X" & ProcessId
'Exit Sub
'Dalje:
'On Error GoTo 0
'End Sub
Public Function IsValidK(ByRef S As String) As Byte
On Error GoTo Eend
Dim K As Byte
K = CByte("&H" & S)
IsValidK = 1
Eend:
On Error GoTo 0
End Function
Public Sub StopAndClear()
Set ModulesExports = Nothing
Set ACTIVEBREAKPOINTS = Nothing
Set SPYBREAKPOINTS = Nothing
Set PROCESSESTHREADS = Nothing
Set ACTMODULESBYPROCESS = Nothing
Set PROCRETBRK = Nothing

Set CheckUNI = Nothing
Set WINS = Nothing
Set BRKW = Nothing
Set BRKWMCMD = Nothing
Erase REFSR
Set INDEXESR = Nothing
Erase EREFSR
Set EINDEXESR = Nothing
Erase SREFSR
Set SINDEXESR = Nothing
Set WATCHES = Nothing
ValidCRef = ""
ExPs.ModuleName = ""
Erase DataPW
gSTARTADR = 0
gLASTADR = 0
gSTARTADR2 = 0
gLASTADR2 = 0
gBegAdr = 0
gLenAdr = 0
DEBUGGYFA = 0
DEBUGGYLA = 0
Unhandled = 0
DebuggyOut = 0
AccThreadX = 0
SusThreadX = 0
EnumRSX = 0
DataArea = 0
UseTrace = 0
ISBPDisabled = 0



StopDebug
Unload Form32
Unload HModules
Unload Form30
Unload Form29
Unload Form28
Unload Form26
Unload Form31
Unload Form27
Unload Form25
Unload Form21: Unload Form20: Unload Form19: Unload Form22: Unload Form23
Unload Form1
Unload Form6
Unload Form9
Unload Form10
Unload Form24
End Sub


Public Function TestPTR(ByVal Address As Long, Optional ByRef DataB As Byte) As Byte
TestPTR = ReadProcessMemory(ProcessHandle, ByVal Address, DataB, 1, ByVal 0&)
End Function

Public Sub ReadStackFrame(LB As ListBox, ByVal Eip As Long, ByVal Ebp As Long)
Dim IsErr As Byte
Dim TA As Long
Dim TA2 As Long
Dim TName As String
Dim u As Long
Dim RAdrs() As Long
LB.Clear
TName = FindInModules(Eip, TA, TA2)
LB.AddItem Hex(Eip) & vbTab & TName
RAdrs = CallStack(Ebp, IsErr)
If IsErr = 1 Then Exit Sub
For u = 0 To UBound(RAdrs)
TName = FindInModules(RAdrs(u), TA, TA2)
LB.AddItem Hex(RAdrs(u)) & vbTab & TName
Next u

End Sub


Public Function CallStack(ByVal Ebp As Long, ByRef IsError As Byte) As Long()
On Error GoTo Dalje
Dim StackFr() As Long
Dim count As Long
ReDim StackFr(10000)

Dim PtAddress(1) As Long
Do
Call ReadProcessMemory(ProcessHandle, ByVal Ebp, PtAddress(0), 8, ByVal 0&)
If Ebp = PtAddress(0) Or PtAddress(0) = 0 Then Exit Do
StackFr(count) = PtAddress(1)
count = count + 1
Ebp = PtAddress(0)
Loop
If count = 0 Then GoTo Dalje
ReDim Preserve StackFr(count - 1)
CallStack = StackFr
Exit Function
Dalje:
On Error GoTo 0
IsError = 1
End Function


Public Sub Read9Stack(TB As TextBox, ByVal StackPos As Long, ByVal EXXp As Long, ByVal EXXpS As String)
Dim FFAdr As Long
Dim u As Long
Dim Xret As Long
Dim Buffy As Long
Dim VD As String

TB = ""
FFAdr = SubBy8(StackPos, 12)

For u = 0 To 6

If EXXp < FFAdr Then
VD = "+" & Hex(SubBy8(FFAdr, EXXp))
ElseIf EXXp > FFAdr Then
VD = "-" & Hex(SubBy8(EXXp, FFAdr))
Else
VD = "**"
End If



Xret = ReadProcessMemory(ProcessHandle, ByVal FFAdr, Buffy, 4, ByVal 0&)
If Xret = 0 Then
TB = TB & "[" & EXXpS & VD & "]" & vbTab & "Not Valid" & vbCrLf
Else
TB = TB & "[" & EXXpS & VD & "]" & vbTab & Hex(Buffy) & vbCrLf
End If
FFAdr = FFAdr + 4
Next u

End Sub



Public Function CheckCALL(ByRef NewVal As Long, Optional ByVal NotForCache As Byte) As String
Dim Redr As String
Dim TName As String
Dim BaAdr As Long

If NOTIFYJMPCALL = 2 Or NOTIFYJMPCALL = 1 Then
'CALL ADR,JMP ADR
TName = FindInModules(VALUES1, BaAdr)
If Len(TName) = 0 Then Exit Function
CheckCALL = GetFromExportsSearch(TName, VALUES1)
If Len(CheckCALL) = 0 Then


Dim OIsValid As Byte
Dim OredTemp() As Byte
GetDataFromMem VALUES1, OredTemp, 16
Dim Ofwr As Byte

If NotForCache = 1 Then
Dim ORGBYTE
ORGBYTE = GetBreakPoint(SPYBREAKPOINTS, VALUES1, Ofwr)
If Ofwr = 0 Then
ORGBYTE = GetBreakPoint(ACTIVEBREAKPOINTS, VALUES1, Ofwr)
End If

If Ofwr = 1 Then OredTemp(0) = ORGBYTE
End If


Dim Ored As String
DASM.BaseAddress = VALUES1
NewVal = VALUES1
Call DASM.DisAssemble(OredTemp, 0, Ofwr, 0, 0)
If NOTIFYJMPCALL = 4 Then
'Redir with JMP DWORD PTR[ ]
Redr = "Redirect to "
GoTo InNtf2

End If

Else
CheckCALL = "Import:" & TName & ":" & CheckCALL
End If

ElseIf NOTIFYJMPCALL = 3 Or NOTIFYJMPCALL = 4 Or NOTIFYJMPCALL = 5 Then
'CALL DWORD [ADR],JMP DWORD [ADR],MOV XXX,DWORD PTR[ADR]
InNtf2:
Dim LxAddr As Long
Call ReadProcessMemory(ProcessHandle, ByVal VALUES1, LxAddr, 4, ByVal 0&)


TName = FindInModules(LxAddr, BaAdr)
If Len(TName) = 0 Or LxAddr = 0 Then NewVal = 0: Exit Function
CheckCALL = GetFromExportsSearch(TName, LxAddr)
If Len(CheckCALL) <> 0 Then
CheckCALL = "Import:" & Redr & TName & ":" & CheckCALL
End If

If Len(Redr) = 0 Then NewVal = LxAddr



End If
End Function

Public Function GetFromIndex(INDX As Collection, REFX() As Collection, ByRef ToAdr As Long) As String
Dim STS() As String
Dim u As Long
Dim RZt As Byte
Dim LYX As Long 'Index if exist!
RZt = CheckExs(INDX, ToAdr, LYX)
If RZt = 0 Then Exit Function
ReDim STS(REFX(LYX).count - 1)
For u = 1 To REFX(LYX).count
STS(u - 1) = Hex(REFX(LYX).Item(u))
Next u
GetFromIndex = "Jumps From:" & Join(STS, ",")
End Function


Public Sub AddInIndex(INDX As Collection, REFX() As Collection, ByRef FromAdr As Long, ByRef ToAdr As Long)
On Error GoTo Dalje

Dim IXX(1) As Long
'0-index
'1-Address To JMP
Dim RZt As Byte
Dim LYX As Long 'Index if exist!
RZt = CheckExs(INDX, ToAdr, LYX)
If RZt = 0 Then
IXX(0) = INDX.count
IXX(1) = ToAdr
INDX.Add IXX, "X" & ToAdr
Set REFX(IXX(0)) = New Collection
REFX(IXX(0)).Add FromAdr
Else
REFX(LYX).Add FromAdr
End If
Exit Sub
Dalje:
On Error GoTo 0
End Sub

Public Function CheckExs(INDX As Collection, ByRef ToAdr As Long, ByRef INX As Long) As Byte
On Error GoTo Dalje
Dim VD() As Long
VD = INDX("X" & ToAdr)
INX = VD(0) 'Uzmi index
CheckExs = 1
Exit Function
Dalje:
On Error GoTo 0
End Function

Public Sub AddInStringIndex(INDX As Collection, REFX() As Collection, ByRef FromAdr As Long, ByRef ToAdr As Long, ByRef StringX As String)
On Error GoTo Dalje

Dim IXX(2) As String
'0-index
'1-Address To JMP
'2-string
Dim RZt As Byte
Dim LYX As Long 'Index if exist!
RZt = CheckStringExs(INDX, ToAdr, LYX)
If RZt = 0 Then
IXX(0) = INDX.count
IXX(1) = ToAdr
IXX(2) = StringX
INDX.Add IXX, "X" & ToAdr
Set REFX(INDX.count - 1) = New Collection
REFX(INDX.count - 1).Add FromAdr
Else
REFX(LYX).Add FromAdr
End If
Exit Sub
Dalje:
On Error GoTo 0
End Sub
Public Function CheckStringExs(INDX As Collection, ByRef ToAdr As Long, ByRef INX As Long) As Byte
On Error GoTo Dalje
Dim VD() As String
VD = INDX("X" & ToAdr)
INX = CLng(VD(0)) 'Uzmi index
CheckStringExs = 1
Exit Function
Dalje:
On Error GoTo 0
End Function
Public Function GetFromStringIndex(INDX As Collection, REFX() As Collection, ByRef ToAdr As Long) As String
Dim STS() As String
Dim u As Long
Dim RZt As Byte
Dim LYX As Long 'Index if exist!
RZt = CheckStringExs(INDX, ToAdr, LYX)
If RZt = 0 Then Exit Function
ReDim STS(REFX(LYX).count - 1)
For u = 1 To REFX(LYX).count
STS(u - 1) = Hex(CLng(REFX(LYX).Item(u)))
Next u
GetFromStringIndex = "Refs From:" & Join(STS, ",")
End Function
Public Function GetStringFPTR(ByVal point As Long, Optional ByVal Mlen As Long = 255, Optional ByVal Unicode As Byte) As String
If Mlen <= 0 Then Exit Function
Dim StringsD() As Byte
Dim iret As Long
Dim lLen As Long
Do
ReDim StringsD(Mlen - 1)
iret = ReadProcessMemory(ProcessHandle, ByVal point, StringsD(0), Mlen, ByVal 0&)
Mlen = Mlen / 2
If Mlen = 0 Then Exit Function
Loop While iret = 0

If Unicode = 1 Then
lLen = lstrlenW(StringsD(0))
Else
lLen = lstrlen(StringsD(0))
End If

GetStringFPTR = Space(lLen)

If Unicode = 1 Then
CopyMemory ByVal StrPtr(GetStringFPTR), StringsD(0), lLen * 2
Else
CopyMemory ByVal GetStringFPTR, StringsD(0), lLen
End If

End Function

Public Function GetStringFromPointer(ByRef point As Long) As String
If (point > &HFFFF&) Or (point < 0) Then
GetStringFromPointer = GetStringFPTR(point)
Else
'If IsType = 0 Then
'GetStringFromPointer = CStr(point)
'Else
'GetStringFromPointer = NName(point)
'End If
GetStringFromPointer = CStr(point)
End If
End Function
Public Function EnumIt(ByVal MxHandle As Long) As Byte
If EnumRSX = 0 Then MsgBox "Cannot Enumerate Resources!", vbExclamation, "Information": Exit Function
ReDim ResRC(80000)
ResCounter = 0
Call CreateRemoteThread(ProcessHandle, ByVal 0&, 10, ByVal EnumRSX, ByVal MxHandle, 0, AccThreadX)
EnumIt = 1
End Function
Public Function NName(ByVal TypeNM As Long) As String
Select Case TypeNM
Case RT_ACCELERATOR
NName = "Accelerator Table"
Case RT_ANICURSOR
NName = "Animated Cursor"
Case RT_ANIICON
NName = "Animated Icon"
Case RT_BITMAP
NName = "Bitmap"
Case RT_CURSOR
NName = "Single Cursor"
Case RT_DIALOG
NName = "Dialog Box"
Case RT_DLGINCLUDE
NName = "DlgBox definition"
Case RT_FONT
NName = "Font"
Case RT_FONTDIR
NName = "Font directory"
Case RT_ICON
NName = "Single Icon"
Case RT_GROUP_CURSOR
NName = "Group Cursor"
Case RT_GROUP_ICON
NName = "Group Icon"
Case RT_HTML
NName = "HTML document"
Case RT_MENU
NName = "Menu"
Case RT_MESSAGETABLE
NName = "Message Table"
Case RT_PLUGPLAY
NName = "Plug and Play"
Case RT_RCDATA
NName = "RC Data"
Case RT_VERSION
NName = "Version Info"
Case RT_VXD
NName = "VXD"
Case RT_STRING
NName = "String"
Case Else
NName = CStr(TypeNM)
End Select
End Function
Public Sub FreezeAll()
Dim InTh() As Long
Dim u As Long
For u = 1 To PROCESSESTHREADS.count
InTh = PROCESSESTHREADS.Item(u)
SuspendThread InTh(0)
Next u
End Sub
Public Sub UnFreezeAll()
Dim InTh() As Long
Dim u As Long
For u = 1 To PROCESSESTHREADS.count
InTh = PROCESSESTHREADS.Item(u)
ResumeThread InTh(0)
Next u
End Sub
Public Function IndexFromPoint(ByRef Handle As Long, ByRef X As Integer, ByRef y As Long) As Long
Dim XY As Long
CopyMemory ByVal VarPtr(XY), X, 2
CopyMemory ByVal (VarPtr(XY) + 2), y, 2
IndexFromPoint = SendMessage(Handle, &H1A9&, ByVal 0&, ByVal XY)
End Function
