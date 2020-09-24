Attribute VB_Name = "ResSave"
Dim Fgfile As Long
Public Sub SaveToRes(ByVal spath As String, ByVal IncVarInf As Byte)
On Error GoTo Greska
Fgfile = FreeFile
Open spath For Binary As #Fgfile

Dim u As Long
Dim TypeX1 As Long
Dim TypeX2() As Byte
Dim NameX1 As Long
Dim NameX2() As Byte
Dim TEMPNAME As String
Dim MEMCONT() As Byte
Dim X As Long

PutRESheader


'For u = 0 To listX.ListCount - 1
For u = 0 To UBound(ResRC)
If ResRC(u).ResType = "1" Or ResRC(u).ResType = "3" Then GoTo Dalje

If IncVarInf = 0 And ResRC(u).ResType = "16" Then GoTo Dalje


Dim ret() As Long
'ret = NameType(CStr(ResName.Item(EXPLIST.Item(u + 1))), CStr(ResType.Item(EXPLIST.Item(u + 1))), NameX2, TypeX2)
ret = NameType(ResRC(u).ResName, ResRC(u).ResType, NameX2, TypeX2)
TypeX1 = ret(0)
NameX1 = ret(1)
Erase ret


If TypeX1 = 14 Or TypeX1 = 12 Then
'Ako je ikona ili kursor----Jebem M$ Umjesto da ide prvo GROUP ikona i kursora pa tek SINGLE,,,oni to bilježe obrnuto!!!!

'Zapamti o kojem se tipu radi....
Dim oldTYPE As Long
If TypeX1 = 14 Then
oldTYPE = 3
ElseIf TypeX1 = 12 Then
oldTYPE = 1
End If

Dim MEM1() As Byte

PutHeadMem NameX1, NameX2, TypeX1, TypeX2, MEM1, True, LongToInt(ResRC(u).LangId), , ResRC(u).ResAddress, ResRC(u).ResLength

'Uzmi broj clanova
Dim tmpHDR As NewHDR
Dim tmpRSDIR() As ResDIR
CopyMemory tmpHDR, MEM1(0), Len(tmpHDR)
ReDim tmpRSDIR(tmpHDR.rescountX - 1)
For X = 1 To tmpHDR.rescountX
'Popuni Clanove sa informacijama
CopyMemory tmpRSDIR(X - 1), MEM1(6 + (X - 1) * Len(RSDIR(0))), Len(RSDIR(0))
Next X


Dim LGGi As Integer
Dim Nadr As Long
Dim Nlen As Long

For X = 1 To tmpHDR.rescountX
'Zapisi Single Icone / Cursor
Dim ret2() As Long
ret2 = NameType(tmpRSDIR(X - 1).iconID, CStr(oldTYPE), NameX2, TypeX2)
TypeX1 = ret2(0)
NameX1 = ret2(1)
Erase ret2


Call FindInListRes(CStr(oldTYPE), CStr(tmpRSDIR(X - 1).iconID), Nadr, Nlen, LGGi)

PutHeadMem NameX1, NameX2, TypeX1, TypeX2, MEMCONT, , LGGi, , Nadr, Nlen
Erase TypeX2
Erase NameX2
Erase MEMCONT
Next X


ret2 = NameType(ResRC(u).ResName, ResRC(u).ResType, NameX2, TypeX2)
TypeX1 = ret2(0)
NameX1 = ret2(1)
Erase ret2
PutHeadMem NameX1, NameX2, TypeX1, TypeX2, MEMCONT, , LongToInt(ResRC(u).LangId), , ResRC(u).ResAddress, ResRC(u).ResLength
Erase TypeX2
Erase NameX2
Erase MEMCONT



Else
'Zapisi sve ostalo!...Normalnim putem....
PutHeadMem NameX1, NameX2, TypeX1, TypeX2, MEMCONT, , LongToInt(ResRC(u).LangId), , ResRC(u).ResAddress, ResRC(u).ResLength
Erase TypeX2
Erase NameX2
Erase MEMCONT
End If


Dalje:
Next u



Close #Fgfile
Exit Sub
Greska:
On Error GoTo 0
MsgBox "Error in Res File!", vbCritical, "Error"
Close #Fgfile
End Sub
Public Sub PutRESheader()
'PRE-HEADER
Put #Fgfile, , CLng(0)
Put #Fgfile, , CLng(&H20)
Put #Fgfile, , CLng(&HFFFF&)
Put #Fgfile, , CLng(&HFFFF&)
Put #Fgfile, , CLng(0)
Put #Fgfile, , CLng(0)
Put #Fgfile, , CLng(0)
Put #Fgfile, , CLng(0)
'END OF PRE-HEADER
End Sub
Public Sub PutHeadMem(ByVal NameX1 As Long, NameX2() As Byte, ByVal TypeX1 As Long, TypeX2() As Byte, MEMCONT() As Byte, Optional OnlyLoadMem As Boolean, Optional LANGX As Integer, Optional OnlySaveName As Boolean, Optional AddressFrom As Long, Optional Length As Long)
Dim ResHedLen As Long 'Resource Header length
Dim nameQ As Boolean
Dim typeQ As Boolean

Dim SZ1 As Long
'Dword Alignment dimenzion
'Svaki resource (bez glavnog pre-headera) mora biti djeljiv sa 4?!

Dim Resst2 As Long
Dim Resst As Long

If OnlySaveName Then SZ1 = UBound(MEMCONT) + 1: GoTo rdalje

GetDataFromMem AddressFrom, MEMCONT, Length
SZ1 = Length

If OnlyLoadMem Then Exit Sub

rdalje:
ResHedLen = 24
If (NameX1 < 0) Or (NameX1 > &HFFFF&) Then
'ResHedLen = ResHedLen + (lstrlen(VarPtr(NameX2(0))) + 1) * 2
ResHedLen = ResHedLen + (UBound(NameX2) + 1) * 2
nameQ = True
Else
ResHedLen = ResHedLen + 4
End If
If (TypeX1 < 0) Or (TypeX1 > &HFFFF&) Then
'ResHedLen = ResHedLen + (lstrlen(VarPtr(TypeX2(0))) + 1) * 2
ResHedLen = ResHedLen + (UBound(TypeX2) + 1) * 2
typeQ = True
Else
ResHedLen = ResHedLen + 4
End If
Put #Fgfile, , SZ1
Resst = ResHedLen Mod 4
If Resst <> 0 Then
ResHedLen = ResHedLen + Resst
End If
Put #Fgfile, , ResHedLen
If typeQ Then
Dim UNI1 As String
ReDim Preserve TypeX2(UBound(TypeX2) - 1)
UNI1 = StrConv(TypeX2, vbUnicode)
UNI1 = StrConv(UNI1, vbUnicode)
Put #Fgfile, , UNI1
Put #Fgfile, , CInt(0)
Else
Put #Fgfile, , CInt(&HFFFF)
Put #Fgfile, , LongToInt(TypeX1)
End If
If nameQ Then
Dim UNI2 As String
ReDim Preserve NameX2(UBound(NameX2) - 1)
UNI2 = StrConv(NameX2, vbUnicode)
UNI2 = StrConv(UNI2, vbUnicode)
Put #Fgfile, , UNI2
Put #Fgfile, , CInt(0)
Else
Put #Fgfile, , CInt(&HFFFF)
Put #Fgfile, , LongToInt(NameX1)
End If
If Resst <> 0 Then Put #Fgfile, , CInt(0)
Put #Fgfile, , CLng(0) 'Data Version
Put #Fgfile, , CInt(&H1030) 'Memory Flag
Put #Fgfile, , LANGX
Put #Fgfile, , CLng(0) 'Version
Put #Fgfile, , CLng(0) 'Characteristic
Put #Fgfile, , MEMCONT 'Put Memory Data
'Postavi da HEADER clana bude djeljiv sa 4
If ((ResHedLen + SZ1) Mod 4) <> 0 Then
Resst2 = ResHedLen + SZ1
Do While (Resst2 Mod 4) <> 0
Put #Fgfile, , CByte(0)
Resst2 = Resst2 + 1
Loop
End If
End Sub
Public Function NameType(ByVal TEMPNAME As String, ByVal TEMPTYPE As String, NameX2() As Byte, TypeX2() As Byte) As Long()
Dim tmpLNG() As Long
ReDim tmpLNG(1)
If Not IsNumeric(TEMPTYPE) Then
TypeX2 = StrConv(TEMPTYPE & Chr(CByte(0)), vbFromUnicode)
tmpLNG(0) = VarPtr(TypeX2(0))
Else
tmpLNG(0) = CLng(TEMPTYPE)
End If
If Not IsNumeric(TEMPNAME) Then
NameX2 = StrConv(TEMPNAME & Chr(CByte(0)), vbFromUnicode)
tmpLNG(1) = VarPtr(NameX2(0))
Else
tmpLNG(1) = CLng(TEMPNAME)
End If
NameType = tmpLNG
End Function
