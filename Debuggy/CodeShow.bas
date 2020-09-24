Attribute VB_Name = "CodeShow"
Public DebuggyOut As Long 'Query Window Proc Procedure
Public AccThreadX As Long 'Created Remote Thread
Public SusThreadX As Long 'Terminate Thread Procedure
Public EnumRSX As Long 'Enum Resources!



Public DataArea As Long 'Data Place in process


Public MainPThread As Long 'Which one is main thread?

Public UseCache As Byte 'Signal to Disasm
Public ValidCRef As String 'Reference Valid For Module!

Public INDEXESR As New Collection 'Internals
Public REFSR() As New Collection

Public EINDEXESR As New Collection 'Externals
Public EREFSR() As New Collection

Public SINDEXESR As New Collection 'STRINGS
Public SREFSR() As New Collection


Public DEBUGGYFA As Long
Public DEBUGGYLA As Long


Public UseTrace As Byte
Public TraceConfig(20) As Long




Public isCC As Long

Public TemConfig(20) As Long
Public ConfigData(20) As Long
Public SPConfig(30) As Long 'special Breaks Config

Public Type ResT
ResType As String
TypeFlag As Byte
ResName As String
NameFlag As Byte
LangId As Long
ResAddress As Long
ResLength As Long
End Type

Public ResRC() As ResT
Public ResCounter As Long

Public ChoosedAdr As Long

Public X1 As String
Public X2 As String
Public X3 As String

Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Const HCBT_ACTIVATE = 5
Public Const HCBT_DESTROYWND = 4
Public Const HCBT_CREATEWND = 3


Public Const WM_DESTROY = &H2
Public Const WM_CLOSE = &H10
Public Const WM_INITDIALOG = &H110
Public Const WM_SIZE = &H5
Public Const WM_SETREDRAW = &HB
Public Const WM_SIZING = &H214
Public Const WM_ACTIVATE = &H6
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206

Public Const WM_PAINT = &HF
Public Const WM_NCPAINT = &H85
Public Const WM_ERASEBKGND = &H14
Public Const WM_DRAWITEM = &H2B
Public Const WM_SETTEXT = &HC
Public Const WM_SETICON = &H80
Public Const WM_SETFONT = &H30
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_SETCURSOR = &H20
Public Const WM_MOUSEMOVE = &H200
Public Const WM_CHAR = &H102
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_NOTIFY = &H4E
Public Const WM_COMMAND = &H111
Public Const WM_VSCROLL = &H115
Public Const WM_HSCROLL = &H114
Public Const WM_INITMENU = &H116
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSKEYDOWN = &H104


Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long
Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public ActiveMemPos As Long
Public NotifyData1(29) As Long 'Real Notify flag
Public NotifyData2(29) As Long 'Values1
Public NotifyData3(29) As Long 'Values2
Public NotifyData4(29) As Long 'Values3


Public FileDataPW() As Byte

Public Function GetHexDump(ByVal Position As Long, Optional ByVal FPW As Byte) As String
On Error GoTo Dalje
GetHexDump = Space(76)


Mid(GetHexDump, 1, 1) = Hex((Position) And &HF0000000)
Mid(GetHexDump, 2, 1) = Hex((Position) And &HF000000)
Mid(GetHexDump, 3, 1) = Hex((Position) And &HF00000)
Mid(GetHexDump, 4, 1) = Hex((Position) And &HF0000)
Mid(GetHexDump, 5, 1) = Hex((Position) And &HF000&)
Mid(GetHexDump, 6, 1) = Hex((Position) And &HF00&)
Mid(GetHexDump, 7, 1) = Hex((Position) And &HF0&)
Mid(GetHexDump, 8, 1) = Hex((Position) And &HF&)

Dim plc As Integer
Dim u As Long
Dim BytD As Byte
Dim ResZ As Byte
Dim Cnnt As Byte
Cnnt = 61
plc = 11
mmax = 16

For u = 0 To 15
If FPW = 0 Then
ResZ = TestPTR(AddBy8(Position, u), BytD)
Else
BytD = FileDataPW(AddBy8(Position, u))
ResZ = 1
End If

If ResZ = 0 Then

Mid(GetHexDump, plc, 2) = "??"
Mid(GetHexDump, Cnnt, 1) = "?"

Else

'Ako je ispravna mem lokacija
Mid(GetHexDump, plc, 1) = Hex(BytD And &HF0)
Mid(GetHexDump, plc + 1, 1) = Hex(BytD And &HF)

If BytD > 13 Then
Mid(GetHexDump, Cnnt, 1) = Chr$(BytD)
Else
Mid(GetHexDump, Cnnt, 1) = Chr(Asc("."))
End If

End If

Cnnt = Cnnt + 1
plc = plc + 3
Next u


Exit Function
Dalje:
On Error GoTo 0
End Function






Public Sub PrintDump(ByVal TXT As TextBox, ByVal Position As Long)
Dim u As Long
Dim DaX(19) As String
For u = 0 To 19
DaX(u) = GetHexDump(AddBy8(Position, (u * 16&))) & vbCrLf
Next u
TXT = Join(DaX, "")
End Sub

Public Function QueryMem(ByVal Address As Long, ByRef MemSData As String) As MEMORY_BASIC_INFORMATION
VirtualQueryEx ProcessHandle, ByVal Address, QueryMem, Len(QueryMem)
If QueryMem.AllocationBase = 0 Then
MemSData = "Invalid Memory Region"
Else
Dim MsDt(3) As String



MsDt(0) = "Test At Address:" & Hex(Address) & ",Base Address:" & Hex(QueryMem.BaseAddress) & ",Region Length:" & Hex(QueryMem.RegionSize)


MsDt(1) = "Protection:"
If (QueryMem.AllocationProtect And PAGE_GUARD) = PAGE_GUARD Then MsDt(1) = MsDt(1) & "GUARD "
If (QueryMem.AllocationProtect And PAGE_NOACCESS) = PAGE_NOACCESS Then MsDt(1) = MsDt(1) & "NOACCESS "
If (QueryMem.AllocationProtect And PAGE_NOCACHE) = PAGE_NOCACHE Then MsDt(1) = MsDt(1) & "NOCACHE "
If (QueryMem.AllocationProtect And PAGE_READONLY) = PAGE_READONLY Then MsDt(1) = MsDt(1) & "READONLY "
If (QueryMem.AllocationProtect And PAGE_READWRITE) = PAGE_READWRITE Then MsDt(1) = MsDt(1) & "READWRITE "
If (QueryMem.AllocationProtect And PAGE_WRITECOPY) = PAGE_WRITECOPY Then MsDt(1) = MsDt(1) & "WRITECOPY "
If (QueryMem.AllocationProtect And PAGE_EXECUTE_READWRITE) = PAGE_EXECUTE_READWRITE Then MsDt(1) = MsDt(1) & "EXECUTEREADWRITE "
If (QueryMem.AllocationProtect And PAGE_EXECUTE_READ) = PAGE_EXECUTE_READ Then MsDt(1) = MsDt(1) & "EXECUTEREAD "
If (QueryMem.AllocationProtect And PAGE_EXECUTE_WRITECOPY) = PAGE_EXECUTE_WRITECOPY Then MsDt(1) = MsDt(1) & "EXECUTEWRITECOPY "

MsDt(2) = "State:"
If (QueryMem.State And MEM_COMMIT) = MEM_COMMIT Then MsDt(2) = MsDt(2) & "COMMIT "
If (QueryMem.State And MEM_RESERVE) = MEM_RESERVE Then MsDt(2) = MsDt(2) & "RESERVE "
If (QueryMem.State And MEM_RELEASE) = MEM_RELEASE Then MsDt(2) = MsDt(2) & "RELEASE "

MsDt(3) = "Type:"
If (QueryMem.lType And MEM_MAPPED) = MEM_MAPPED Then MsDt(3) = MsDt(3) & "MAPPED "
If (QueryMem.lType And MEM_IMAGE) = MEM_IMAGE Then MsDt(3) = MsDt(3) & "IMAGE "
If (QueryMem.lType And MEM_PRIVATE) = MEM_PRIVATE Then MsDt(3) = MsDt(3) & "PRIVATE "


MemSData = Join(MsDt, vbCrLf)

End If


End Function


Public Sub PlayS()
Beep 595, 3
Beep 2900, 1
Beep 11, 2
End Sub
Public Function IsStringOnAdr(Optional ByRef IsString As Long, Optional ByVal AddressFrom As Long, Optional ByVal ToCache As Byte) As String
Dim AaX As Long
If VALUES1 <> 0 Then
AaX = VALUES1
ElseIf VALUES3 <> 0 Then
AaX = VALUES3
End If

IsString = 0
If AaX >= 0 And AaX <= 65535 Then Exit Function

Dim IsWp As Long
Dim XYData() As Byte
Dim IsValidPRef As Long
GetDataFromMem SubBy8(AaX, 4), XYData, 260, IsWp


If IsWp = 0 Then Exit Function

Dim ret As Long

Dim TMaxLen As Long
TMaxLen = 256
Dim LenStrX As Long
Dim CCLen As Long
CopyMemory LenStrX, XYData(0), 4
CCLen = lstrlenW(XYData(4))


If LenStrX <= 0 Or LenStrX > 65535 Then GoTo InAnsi
If LenStrX > 256 Then CopyMemory XYData(0), TMaxLen, 4: LenStrX = 256


If (CCLen * 2) = LenStrX Then
'Pravi Unicode
IsStringOnAdr = Space(LenStrX / 2)
CopyMemory ByVal StrPtr(IsStringOnAdr), XYData(4), LenStrX
IsString = 1

If ToCache = 1 Then
AddInStringIndex SINDEXESR, SREFSR, AddressFrom, AaX, IsStringOnAdr
End If


Exit Function
Else
InAnsi:
XYData(259) = 0
Dim AClen As Long

AClen = lstrlen(XYData(4))
If AClen = 0 Then Exit Function
Dim u As Long
For u = 1 To AClen

If IsCharVD(XYData(3 + u)) = 0 Then Exit Function

Next u

'Ansi String!
IsStringOnAdr = Space(AClen)
CopyMemory ByVal IsStringOnAdr, XYData(4), AClen
IsString = 1

If ToCache = 1 Then
AddInStringIndex SINDEXESR, SREFSR, AddressFrom, AaX, IsStringOnAdr
End If

Exit Function
End If


IsString = 0
End Function
Public Function IsCharVD(ByRef BT As Byte) As Byte
Select Case BT
Case Is = 9, 13, 32 To 128
IsCharVD = 1
End Select
End Function


Public Sub RemoveX(ByVal hWnd As Long)
Const SC_CLOSE = &HF060
Dim hMenu As Long
hMenu = GetSystemMenu(hWnd, 0&)
If hMenu Then
Call DeleteMenu(hMenu, SC_CLOSE, 0)
DrawMenuBar (hWnd)
End If
CloseHandle hMenu
End Sub
Public Sub OnScreen(ByVal hWnd As Long)
ReleaseCapture
SendMessage hWnd, &HA1, 2, 0&
End Sub
Public Sub NoResize(ByVal hWnd As Long)
Dim LNG As Long
LNG = -1865809920
SetWindowLong hWnd, -16, LNG
End Sub

Public Sub TouchIt(ByVal ThreadId As Long)
If Form18.ShowingTH = ThreadId Then OnScreen Form18.hWnd: Form18.ReadIt
If Form16.FBASE.ShowingTH = ThreadId Then OnScreen Form16.FBASE.hWnd: Form16.FBASE.ReadIt
If Form16.FSTACK.ShowingTH = ThreadId Then OnScreen Form16.FSTACK.hWnd: Form16.FSTACK.ReadIt
If Form7.ShowTH = ThreadId Then OnScreen Form7.hWnd: Form7.ReadIt
If Form24.Visible = True Then OnScreen Form24.hWnd: Form24.ReadExpresses

End Sub

Public Function AddSlash(ByRef StringX As String) As String
If Right(StringX, 1) <> "\" Then AddSlash = StringX & "\"
End Function
Public Sub SelectRange(ByVal Handle As Long, ByVal FirstIndex As Long, ByVal LastIndex As Long)
Call SendMessage(Handle, &H183, ByVal FirstIndex, ByVal LastIndex)
End Sub
Public Sub ClearSelected(ByVal Handle As Long)
Call SendMessage(Handle, &H185, ByVal 0&, ByVal True)
End Sub
Public Function GetSelectedItems(ByVal Handle As Long, ByRef Items() As Long) As Long
GetSelectedItems = SendMessage(Handle, &H190, ByVal 0&, ByVal 0&)
If GetSelectedItems = 0 Then GetSelectedItems = 0: Exit Function
ReDim Items(GetSelectedItems - 1)
GetSelectedItems = SendMessage(Handle, &H191, ByVal GetSelectedItems, Items(0))
End Function
Public Sub SpeedUpAdding(ByVal Handle As Long, ByVal NumberOfItems As Long, ByVal MemoryReservation As Long)
Call SendMessage(Handle, &H1A8, ByVal NumberOfItems, ByVal MemoryReservation)
End Sub



Public Sub SleepMe(ByVal TM As Long)
Dim TC As Long
Dim TC2 As Long
TC = GetTickCount
Do
TC2 = GetTickCount
DoEvents
Loop While SubBy8(TC2, TC) < TM
End Sub
