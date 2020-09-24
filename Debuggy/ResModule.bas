Attribute VB_Name = "ResModule"
Option Explicit

Public Type GUID
  dwData1 As Long
  wData2 As Integer
  wData3 As Integer
  abData4(7) As Byte
End Type
Type PictureDescription
    cbSizeofStruct As Long
    PicType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function LoadMenuIndirect Lib "user32" Alias "LoadMenuIndirectA" (ByVal lpMenuTemplate As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long

Declare Function CreateDialogIndirectParam Lib "user32" Alias "CreateDialogIndirectParamA" (ByVal hInstance As Long, lpTemplate As Any, ByVal hwndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Long
Declare Function GetDlgCtrlID Lib "user32" (ByVal hWnd As Long) As Long
Declare Function CreateStreamOnHGlobal Lib "ole32" _
                              (ByVal hGlobal As Long, _
                              ByVal fDeleteOnRelease As Long, _
                              ppstm As Any) As Long
Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictureDescription, riid As GUID, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Declare Function CreateIconFromResourceEx Lib "user32" (presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal uFlags As Long) As Long

Declare Function OleLoadPicture Lib "olepro32" _
                              (pStream As Any, _
                              ByVal lSize As Long, _
                              ByVal fRunmode As Long, _
                              riid As GUID, _
                              ppvObj As Any) As Long

Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As GUID) As Long
Const sIID_IPicture = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
Public STD1() As Picture
Public PicWidth() As Long
Public PicHeight() As Long

Public ResResData() As Byte

Public Type NewHDR
 reserved As Integer
 restypeX As Integer
 rescountX As Integer
End Type

Public Type CursorResDir
Width As Integer
Height As Integer
End Type

Public Type CResDir
curresD As CursorResDir
hotXY As Long
bitesinres As Long
cursorID As Integer
End Type

Public Type IconResDir
 Width As Byte
 Height As Byte
 colorCount As Byte
 reserved As Byte
End Type

Public Type ResDIR
 iconresD As IconResDir
 planes As Integer
 bitcount As Integer
 bitesinres As Long
 iconID As Integer
End Type

Public Type LocalHDR
 Xspot As Integer
 Yspor As Integer
End Type

Public Type ResStringEx
id As Long
Data As String
End Type

Public Type MessageResBlock
 lowId As Long
 HighId As Long
 EntryPoint As Long
End Type
Public MRB() As MessageResBlock

Public LDSTRINGS() As ResStringEx

Public CRSDIR() As CResDir
Public RSDIR() As ResDIR

Public NEWH As NewHDR

Public BITCRI As New Collection

Const GMEM_MOVEABLE = &H2
Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByValdwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long



Public Function FixBitmap(ByRef DataP() As Byte, ByVal Address As Long, ByVal Length As Long) As Byte
Dim iSInvalid As Long
ReDim DataP(Length - 1 + 14)
iSInvalid = ReadProcessMemory(ProcessHandle, ByVal Address, DataP(14), Length, ByVal 0&)
If iSInvalid = 0 Then Exit Function
DataP(0) = &H42
DataP(1) = &H4D
CopyMemory DataP(2), AddBy8(Length, 14), 4
Dim OFFSET As Long
Dim CLRUSED As Long
CopyMemory CLRUSED, DataP(50), 4
Dim COMPRSED As Long
CopyMemory COMPRSED, DataP(30), 4
OFFSET = 56 - 2
Select Case DataP(28)
Case 1
OFFSET = OFFSET + 8
Case 4
If Not CBool(CLRUSED) Then
OFFSET = OFFSET + 64
Else
OFFSET = OFFSET + CLRUSED * 4
End If
Case 8
If Not CBool(CLRUSED) Then
OFFSET = OFFSET + 1024
Else
OFFSET = OFFSET + CLRUSED * 4
End If
End Select
CopyMemory DataP(10), OFFSET, 4
FixBitmap = 1
End Function
Public Function FixCursor(ByRef DataP() As Byte, ByVal Address As Long, ByVal Length As Long) As Byte
Dim iSInvalid As Long
ReDim DataP(22 - 1 + Length)
iSInvalid = ReadProcessMemory(ProcessHandle, ByVal Address, DataP(18), Length, ByVal 0&)
If iSInvalid = 0 Then Exit Function
DataP(0) = 0
DataP(2) = 2
DataP(4) = 1
DataP(6) = DataP(26)
DataP(7) = DataP(26)
CopyMemory DataP(10), DataP(18), 4
CopyMemory DataP(14), Length, 4
CopyMemory DataP(18), CLng(&H16), 4
FixCursor = 1
End Function
Public Function FixIcon(ByRef DataP() As Byte, ByVal Address As Long, ByVal Length As Long) As Byte
Dim iSInvalid As Long
ReDim DataP(22 - 1 + Length)
iSInvalid = ReadProcessMemory(ProcessHandle, ByVal Address, DataP(22), Length, ByVal 0&)
If iSInvalid = 0 Then Exit Function
DataP(0) = 0
DataP(2) = 1
DataP(4) = 1
DataP(6) = DataP(22)
DataP(7) = DataP(22)
CopyMemory DataP(10), DataP(30), 2
CopyMemory DataP(12), DataP(32), 2
CopyMemory DataP(14), Length, 4
CopyMemory DataP(18), CLng(&H16), 4
Dim BPTMP As Integer
CopyMemory BPTMP, DataP(12), 2
If BPTMP = 4 Then
DataP(8) = CByte(16)
ElseIf BPTMP = 16 Then
DataP(8) = 0
ElseIf BPTMP = 1 Then
DataP(8) = 2
Else
DataP(8) = 0
End If
FixIcon = 1
End Function
Public Function CursorGroup(ByVal Address As Long, ByVal Length As Long) As Byte

Dim DataPP() As Byte
Dim DStore() As Byte
Dim iSInvalid As Long
ReDim DataPP(Length - 1)
iSInvalid = ReadProcessMemory(ProcessHandle, ByVal Address, DataPP(0), Length, ByVal 0&)
If iSInvalid = 0 Then Exit Function

Dim Nadr As Long
Dim Nlen As Long
'Uzmi broj clanova
Dim u As Long
CopyMemory NEWH, DataPP(0), Len(NEWH)
ReDim STD1(NEWH.rescountX - 1)
ReDim CRSDIR(NEWH.rescountX - 1)
ReDim PicWidth(NEWH.rescountX - 1)
ReDim PicHeight(NEWH.rescountX - 1)

For u = 0 To NEWH.rescountX - 1
Nadr = 0
Nlen = 0
CopyMemory CRSDIR(u), DataPP(6 + (u) * Len(CRSDIR(0))), Len(CRSDIR(0))

If 0 = FindInListRes("1", CStr(CRSDIR(u).cursorID), Nadr, Nlen) Then Exit Function
If 0 = FixCursor(ResResData, Nadr, Nlen) Then Exit Function

Set STD1(u) = GetPicture(ResResData)
If STD1(u) = 0 Then Exit Function

GetDataFromMem Nadr, DStore, Nlen
BITCRI.Add DStore

PicWidth(u) = CLng(STD1(u).Width * (567 / 1000))
PicHeight(u) = CLng(STD1(u).Height * (567 / 1000))

Next u
CursorGroup = 1
Erase DStore
Erase DataPP
End Function

Public Function IconGroup(ByVal Address As Long, ByVal Length As Long) As Byte

Dim DataPP() As Byte
Dim DStore() As Byte
Dim iSInvalid As Long
ReDim DataPP(Length - 1)
iSInvalid = ReadProcessMemory(ProcessHandle, ByVal Address, DataPP(0), Length, ByVal 0&)
If iSInvalid = 0 Then Exit Function

Dim Nadr As Long
Dim Nlen As Long
'Uzmi broj clanova
Dim u As Long
CopyMemory NEWH, DataPP(0), Len(NEWH)
ReDim STD1(NEWH.rescountX - 1)
ReDim RSDIR(NEWH.rescountX - 1)
ReDim PicWidth(NEWH.rescountX - 1)
ReDim PicHeight(NEWH.rescountX - 1)

For u = 0 To NEWH.rescountX - 1
Nadr = 0
Nlen = 0
CopyMemory RSDIR(u), DataPP(6 + (u) * Len(RSDIR(0))), Len(RSDIR(0))
If 0 = FindInListRes("3", CStr(RSDIR(u).iconID), Nadr, Nlen) Then Exit Function
If 0 = FixIcon(ResResData, Nadr, Nlen) Then Exit Function

Set STD1(u) = GetIconToPicture(ResResData)
If STD1(u) = 0 Then Exit Function

GetDataFromMem Nadr, DStore, Nlen
BITCRI.Add DStore

PicWidth(u) = CLng(STD1(u).Width * (567 / 1000))
PicHeight(u) = CLng(STD1(u).Height * (567 / 1000))

Next u
IconGroup = 1
Erase DStore
Erase DataPP
End Function
Public Sub SaveCursor(ByVal FNMM As Long)


Dim CalcEntry As Long
CalcEntry = Len(NEWH) + NEWH.rescountX * (Len(CRSDIR(0)) + 2) 'izracunaj duzinu Headera..

Dim HOTSPOT As LocalHDR
Dim TwlData() As Byte

Put #FNMM, , NEWH
Dim u As Long
For u = 0 To NEWH.rescountX - 1

Put #FNMM, , CByte(CRSDIR(u).curresD.Width)
Put #FNMM, , CByte(CRSDIR(u).curresD.Height)
Put #FNMM, , CByte(0)
Put #FNMM, , CByte(0)
Put #FNMM, , CRSDIR(u).hotXY
Put #FNMM, , CRSDIR(u).bitesinres - 4
Put #FNMM, , CalcEntry
CalcEntry = CalcEntry + CRSDIR(u).bitesinres - 4
TwlData = BITCRI.Item(u + 1)
CopyMemory HOTSPOT, TwlData(0), Len(HOTSPOT)
Seek #FNMM, Loc(FNMM) - 11
Put #FNMM, , HOTSPOT
Seek #FNMM, LOF(FNMM) + 1
Next u


Dim ReData() As Byte
For u = 0 To NEWH.rescountX - 1
TwlData = BITCRI.Item(u + 1)
ReDim ReData(UBound(TwlData) - 4)
CopyMemory ReData(0), TwlData(4), UBound(ReData) + 1
Put #FNMM, , ReData
Next u


Erase TwlData
Erase ReData
End Sub
Public Sub SaveIcon(ByVal FNMM As Long)

Dim CalcEntry As Long
CalcEntry = Len(NEWH) + NEWH.rescountX * (Len(RSDIR(0)) + 2) 'izracunaj duzinu Headera..

Dim u As Long
Dim TwlData() As Byte

Put #FNMM, , NEWH
For u = 0 To NEWH.rescountX - 1
Put #FNMM, , RSDIR(u)
Seek #FNMM, Loc(FNMM) - 1
Put #FNMM, , CalcEntry
CalcEntry = CalcEntry + RSDIR(u).bitesinres
Next u

For u = 0 To NEWH.rescountX - 1
TwlData = BITCRI.Item(u + 1)
Put #FNMM, , TwlData
Next u


Erase TwlData
End Sub

Public Function FindInListRes(ByVal TypeO As String, ByVal NameO As String, ByRef Address As Long, ByRef Length As Long, Optional ByRef LangId As Integer) As Byte
Dim u As Long
For u = 0 To UBound(ResRC)
If StrComp(TypeO, ResRC(u).ResType) = 0 And StrComp(NameO, ResRC(u).ResName) = 0 Then
Address = ResRC(u).ResAddress: Length = ResRC(u).ResLength: FindInListRes = 1: Exit Function
LangId = ResRC(u).LangId
End If
Next u
End Function

Public Function GetIconToPicture(Data() As Byte) As IPicture
Dim hMem  As Long
Dim lpMem  As Long
hMem = GlobalAlloc(GMEM_MOVEABLE, UBound(Data) + 1)
lpMem = GlobalLock(hMem)
CopyMemory ByVal lpMem, Data(0), UBound(Data) + 1
Dim IID_IPicture As GUID
Call CLSIDFromString(StrPtr(sIID_IPicture), IID_IPicture)
Dim hIcon As Long
hIcon = CreateIconFromResourceEx(ByVal (lpMem + &H16), UBound(Data) + 1, 1, &H30000, 0, 0, &H1000)
Call GlobalUnlock(hMem)
Dim tPicConv As PictureDescription
With tPicConv
.cbSizeofStruct = Len(tPicConv)
.PicType = vbPicTypeIcon
.hImage = hIcon
End With
OleCreatePictureIndirect tPicConv, IID_IPicture, True, GetIconToPicture
Call GlobalFree(hMem)
DestroyIcon hIcon
End Function

Public Function GetPicture(ByRef DataP() As Byte) As IPicture
Dim hMem  As Long
Dim lpMem  As Long
Dim IID_IPicture As GUID
Dim istm As stdole.IUnknown
Dim ipic As IPicture
hMem = GlobalAlloc(GMEM_MOVEABLE, UBound(DataP) + 1)
lpMem = GlobalLock(hMem)
CopyMemory ByVal lpMem, DataP(0), UBound(DataP) + 1
Call GlobalUnlock(hMem)
Call CreateStreamOnHGlobal(hMem, 1, istm)
Call CLSIDFromString(StrPtr(sIID_IPicture), IID_IPicture)
Call OleLoadPicture(ByVal ObjPtr(istm), UBound(DataP) + 1, 0, IID_IPicture, GetPicture)
Call GlobalFree(hMem)
End Function
Public Sub LoadStr(ByVal Address As Long, ByVal Length As Long, ByVal entry As Long, ByRef IsNotEmpty As Byte, ByRef tmpBFR() As ResStringEx)
On Error GoTo Dalje
Dim DataP() As Byte
Dim iSInvalid As Long
ReDim DataP(Length - 1)
iSInvalid = ReadProcessMemory(ProcessHandle, ByVal Address, DataP(0), Length, ByVal 0&)
If iSInvalid = 0 Then IsNotEmpty = 0: Exit Sub
Dim u As Long
ReDim tmpBFR(15)
Dim CountY As Integer
Dim CountX As Long
Dim CHECKLNG As Integer
Dim CHECKLNG1 As Long
entry = (entry - 1) * 16
For u = 0 To 15
CopyMemory CHECKLNG, DataP(CountX), 2
CHECKLNG1 = IntToLong(CHECKLNG)
If CHECKLNG1 = 0 Then CountX = CountX + 2: GoTo OpTs:
If CHECKLNG1 > Length Then GoTo Dalje
tmpBFR(CountY).Data = Space(CHECKLNG1)
CopyMemory ByVal StrPtr(tmpBFR(CountY).Data), DataP(CountX + 2), CHECKLNG1 * 2
CountX = CountX + 2 + CHECKLNG1 * 2
tmpBFR(CountY).id = entry
CountY = CountY + 1
OpTs:
entry = entry + 1
Next u
ReDim Preserve tmpBFR(CountY - 1)
IsNotEmpty = 1
Exit Sub
Dalje:
On Error GoTo 0
IsNotEmpty = 0
End Sub

Public Sub LoadMSGTable(ByVal Address As Long, ByVal Length As Long, ByRef IsNotEmpty As Byte, ByRef tmpBFR() As ResStringEx)
On Error GoTo Dalje
Dim DataP() As Byte
Dim iSInvalid As Long
ReDim DataP(Length - 1)
iSInvalid = ReadProcessMemory(ProcessHandle, ByVal Address, DataP(0), Length, ByVal 0&)
If iSInvalid = 0 Then IsNotEmpty = 0: Exit Sub
Dim lenEntries As Long
Dim CountX As Long
Dim totalEntries As Long
Dim Xcnt As Long
Dim lLen As Integer: Dim lLen1 As Long
Dim Llent As Integer: Dim lLent1 As Long
Dim Flag As Integer
Dim u As Long
Dim uu As Long
CopyMemory lenEntries, DataP(0), 4
CountX = 4
ReDim MRB(lenEntries - 1)
For u = 1 To lenEntries
CopyMemory MRB(u - 1), DataP(CountX), Len(MRB(u - 1))
CountX = CountX + Len(MRB(u - 1))
totalEntries = totalEntries + (MRB(u - 1).HighId - MRB(u - 1).lowId) + 1
Next u
ReDim tmpBFR(totalEntries - 1)
For u = 1 To lenEntries
For uu = MRB(u - 1).lowId To MRB(u - 1).HighId
CopyMemory lLen, DataP(CountX), 2
lLen1 = IntToLong(lLen)
CopyMemory Flag, DataP(CountX + 2), 2
CountX = CountX + 4
If Not CBool(Flag) Then
lLent1 = lstrlen(ByVal VarPtr(DataP(CountX)))
If lLent1 > Length Then GoTo Dalje
tmpBFR(Xcnt).Data = Space(lLent1)
CopyMemory ByVal tmpBFR(Xcnt).Data, DataP(CountX), lLent1
Else
lLent1 = lstrlenW(ByVal VarPtr(DataP(CountX)))
If lLent1 > Length Then GoTo Dalje
tmpBFR(Xcnt).Data = Space(lLent1)
CopyMemory ByVal StrPtr(tmpBFR(Xcnt).Data), DataP(CountX), lLent1 * 2
End If
CountX = CountX + lLen1 - 4
tmpBFR(Xcnt).id = uu
Xcnt = Xcnt + 1
Next uu
Next u
IsNotEmpty = 1
Exit Sub
Dalje:
On Error GoTo 0
IsNotEmpty = 0
End Sub

Public Function dProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
End Function
Public Function EnumChildRS(ByVal hWnd As Long, ByVal Param1 As Long) As Long
Dim wlen As Long
Dim TextPrev As String
wlen = GetWindowTextLength(hWnd)
If wlen = 0 Then
TextPrev = " (No Text)"
Else
TextPrev = Space(wlen + 1)
If wlen > 256 Then wlen = 256
GetWindowText hWnd, TextPrev, wlen + 1
End If
EnableWindow hWnd, 1
ShowWindow hWnd, 1

Form22.List1.AddItem "Ctrl Id:" & Hex(GetDlgCtrlID(hWnd)) & " ,Class Name:" & ClassNameEx(hWnd) & " ,Style:" & Hex(GetWindowLong(hWnd, -16)) _
& " ,Ex Style:" & Hex(GetWindowLong(hWnd, -20)) & " ,Text:" & TextPrev
EnumChildRS = 1
End Function

