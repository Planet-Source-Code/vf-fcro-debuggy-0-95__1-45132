Attribute VB_Name = "WatchMdl"
Option Explicit

Public WATCHES As New Collection


Public Sub AddInWatches(ByRef Address As Long, ByRef TypeOfW As Long, Optional ByRef IsAccept As Long)
'Type=0 BYTE,1 WORD,2 DWORD,3 DOUBLE,4 ANSI,5 UNICODE
'Type=10 BYTE PTR,11 WORD PTR,12 DWORD PTR,13 DOUBLE PTR,14 ANSI STRING,15 UNICODE STRING
On Error GoTo Dalje
Dim TYP(2) As Long
TYP(0) = Address
TYP(1) = TypeOfW
WATCHES.Add TYP, "X" & Address & "Y" & TypeOfW
IsAccept = 1
Exit Sub
Dalje:
On Error GoTo 0
IsAccept = 0
End Sub
Public Sub RemoveInWatches(ByRef Address As Long, ByRef TypeOfW As Long)
On Error GoTo Dalje
WATCHES.Remove "X" & Address & "Y" & TypeOfW
Exit Sub
Dalje:
On Error GoTo 0
End Sub
'VALUES
Public Function ReadAsByte(ByVal Address As Long, ByRef IsValid As Long) As String
Dim Ep1 As Byte
IsValid = ReadProcessMemory(ProcessHandle, ByVal Address, Ep1, 1, ByVal 0&)
If IsValid = 0 Then
ReadAsByte = "<INVALID ADRESS>"
Else
ReadAsByte = Hex(Ep1)
End If
End Function
Public Function ReadAsWord(ByVal Address As Long, ByRef IsValid As Long) As String
Dim Ep24 As Long
IsValid = ReadProcessMemory(ProcessHandle, ByVal Address, Ep24, 2, ByVal 0&)
If IsValid = 0 Then
ReadAsWord = "<INVALID ADRESS>"
Else
ReadAsWord = Hex(Ep24)
End If
End Function
Public Function ReadAsDword(ByVal Address As Long, ByRef IsValid As Long) As String
Dim Ep24 As Long
IsValid = ReadProcessMemory(ProcessHandle, ByVal Address, Ep24, 4, ByVal 0&)
If IsValid = 0 Then
ReadAsDword = "<INVALID ADRESS>"
Else
ReadAsDword = Hex(Ep24)
End If
End Function
Public Function ReadAsQuad(ByVal Address As Long, ByRef IsValid As Long) As String
On Error GoTo Dalje:
Dim Ep8 As Double
IsValid = ReadProcessMemory(ProcessHandle, ByVal Address, Ep8, 8, ByVal 0&)
If IsValid = 0 Then
ReadAsQuad = "<INVALID ADRESS>"
Else
ReadAsQuad = Ep8
End If
Exit Function
Dalje:
On Error GoTo 0
ReadAsQuad = "<INVALID QUAD EXPRESSION>"
IsValid = 1
End Function



'*PTRS
Public Function ReadAsBytePTR(ByVal Address As Long, ByRef IsValid As Long) As String
Dim PTRS As Long
IsValid = ReadProcessMemory(ProcessHandle, ByVal Address, PTRS, 4, ByVal 0&)
If IsValid = 0 Then ReadAsBytePTR = "<INVALID ADRESS>": Exit Function
Dim Ep1 As Byte
IsValid = ReadProcessMemory(ProcessHandle, ByVal PTRS, Ep1, 1, ByVal 0&)
If IsValid = 0 Then
ReadAsBytePTR = "<INVALID PTR>"
Else
ReadAsBytePTR = Hex(Ep1)
End If
End Function
Public Function ReadAsWordPTR(ByVal Address As Long, ByRef IsValid As Long) As String
Dim PTRS As Long
IsValid = ReadProcessMemory(ProcessHandle, ByVal Address, PTRS, 4, ByVal 0&)
If IsValid = 0 Then ReadAsWordPTR = "<INVALID ADRESS>": Exit Function
Dim Ep24 As Long
IsValid = ReadProcessMemory(ProcessHandle, ByVal PTRS, Ep24, 2, ByVal 0&)
If IsValid = 0 Then
ReadAsWordPTR = "<INVALID PTR>"
Else
ReadAsWordPTR = Hex(Ep24)
End If
End Function
Public Function ReadAsDwordPTR(ByVal Address As Long, ByRef IsValid As Long) As String
Dim PTRS As Long
IsValid = ReadProcessMemory(ProcessHandle, ByVal Address, PTRS, 4, ByVal 0&)
If IsValid = 0 Then ReadAsDwordPTR = "<INVALID ADRESS>": Exit Function
Dim Ep24 As Long
IsValid = ReadProcessMemory(ProcessHandle, ByVal PTRS, Ep24, 4, ByVal 0&)
If IsValid = 0 Then
ReadAsDwordPTR = "<INVALID PTR>"
Else
ReadAsDwordPTR = Hex(Ep24)
End If
End Function
Public Function ReadAsQuadPTR(ByVal Address As Long, ByRef IsValid As Long) As String
On Error GoTo Dalje:
Dim Ep8 As Double
Dim PTRS As Long
IsValid = ReadProcessMemory(ProcessHandle, ByVal Address, PTRS, 4, ByVal 0&)
If IsValid = 0 Then ReadAsQuadPTR = "<INVALID ADRESS>": Exit Function
IsValid = ReadProcessMemory(ProcessHandle, ByVal PTRS, Ep8, 8, ByVal 0&)
If IsValid = 0 Then
ReadAsQuadPTR = "<INVALID PTR>"
Else
ReadAsQuadPTR = Ep8
End If
Exit Function
Dalje:
On Error GoTo 0
IsValid = 1
ReadAsQuadPTR = "<INVALID QUAD EXPRESSION>"
End Function


Public Function GetDBBLock(ByRef Address As Long) As String
GetDBBLock = Space(48)
Dim u As Long
Dim DBy1 As Byte
Dim Tocnt As Long
Dim IsValid As Long
Tocnt = 1
For u = 0 To 15
IsValid = ReadProcessMemory(ProcessHandle, ByVal AddBy8(Address, u), DBy1, 1, ByVal 0&)
If IsValid <> 0 Then
Mid(GetDBBLock, Tocnt, 1) = Hex((DBy1 And &HF0) / &H10)
Mid(GetDBBLock, Tocnt + 1, 1) = Hex(DBy1 And &HF)
Else
Mid(GetDBBLock, Tocnt, 2) = "??"
End If
Tocnt = Tocnt + 3
Next u
End Function


Public Sub FindLastEA(TextX1 As TextBox, TextX2 As TextBox)
If Len(TextX1) = 0 Then MsgBox "From Address isn't set!", vbInformation, "Information": Exit Sub
Dim LPP As Long
Dim IsValid As Byte
Dim TBuff(127) As Byte
LPP = CLng("&H" & TextX1)
TextX2 = ""


Do
IsValid = ReadProcessMemory(ProcessHandle, ByVal LPP, TBuff(0), 128, ByVal 0&)
LPP = LPP + 4
Loop While IsValid <> 0
LPP = LPP - 4

Do
IsValid = ReadProcessMemory(ProcessHandle, ByVal LPP, TBuff(0), 1, ByVal 0&)
LPP = LPP + 1
Loop While IsValid <> 0

LPP = LPP - 1

TextX2 = Hex(LPP)
Erase TBuff


End Sub
