Attribute VB_Name = "Module4"
Public ModulesExports As New Collection
Public WINS As New Collection



Public Sub AddWins(ByVal ClassNm As String, ByVal hWnd As Long, ByVal ThreadId As Long)
On Error GoTo Dalje
Dim WSd(2) As String
WSd(0) = ClassNm
WSd(1) = CStr(hWnd)
WSd(2) = CStr(ThreadId)

WINS.Add WSd, "X" & hWnd
Exit Sub
Dalje:
On Error GoTo 0
End Sub
Public Sub RemoveWins(ByVal hWnd As Long)
On Error GoTo Dalje
WINS.Remove "X" & hWnd
Exit Sub
Dalje:
On Error GoTo 0
End Sub


Public Sub AddInExportsSearch(ByVal MName As String, ByVal BaseAdr As Long)
On Error GoTo Dalje
Dim IMPFF As Byte
Dim EXPFF As Byte

ReadPE2 BaseAdr, IMPFF, EXPFF

If Len(MName) = 0 Then MName = ExPs.ModuleName

If StrComp(UCase(MName), "DEBUGGY.DLL") = 0 Then
DEBUGGYFA = BaseAdr
DEBUGGYLA = AddBy8(BaseAdr, NTHEADER.OptionalHeader.SizeOfImage)
End If


Dim u As Long
Dim CNew As New Collection

If NTHEADER.OptionalHeader.AddressOfEntryPoint <> 0 Then
AddToCols CNew, AddBy8(BaseAdr, NTHEADER.OptionalHeader.AddressOfEntryPoint), "Entry Point"
End If

If EXPFF = 1 Then
For u = 0 To UBound(ExPs.FuncNames)
AddToCols CNew, ExPs.FuncAddress(u), ExPs.FuncNames(u)
If StrComp(UCase(ExPs.ModuleName), "DEBUGGY.DLL") = 0 Then
    If StrComp(ExPs.FuncNames(u), "VFDebuggerThread") = 0 Then DebuggyOut = ExPs.FuncAddress(u)
    If StrComp(ExPs.FuncNames(u), "VFDebuggerTerminate") = 0 Then SusThreadX = ExPs.FuncAddress(u)
    If StrComp(ExPs.FuncNames(u), "VFDebuggerEnumRS") = 0 Then EnumRSX = ExPs.FuncAddress(u)
Else


AddSpiesBP NameFromPath(MName), u



End If
Next u
End If



ModulesExports.Add CNew, NameFromPath(MName)
Exit Sub
Dalje:
On Error GoTo 0
End Sub
Public Sub AddSpiesBP(ByRef MName As String, ByVal CounterLoop As Long)
Dim TAlloc As Long

'If PROCESSALLOC.count = 0 Then Exit Sub

If StrComp(UCase(MName), "KERNEL32.DLL") = 0 Then
   
' Select Case ExPs.FuncNames(CounterLoop)
    
' Case Is = "SendMessageA", "SendDlgItemMessageA", "SendMessageW", "SendDlgItemMessageW"
'   If ConfigData(7) = 0 Then
'   AddBreakPoint SPYBREAKPOINTS, ExPs.FuncAddress(CounterLoop), 1
'   Else
'   AddBreakPoint SPYBREAKPOINTS, ExPs.FuncAddress(CounterLoop), 0
'   End If
   
' End Select

  Select Case ExPs.FuncNames(CounterLoop)
    
 Case Is = "CreateFileA", "CreateFileW", "OpenFile", "CreateFileMappingA", _
 "CreateFileMappingW", "OpenFileMappingA", "OpenFileMappingW", "ReadFile", _
 "ReadFileEx", "WriteFile", "WriteFileEx", "MapViewOfFile", "MapViewOfFileEx", _
 "CreateDirectoryA", "CreateDirectoryW", "CreateDirectoryExA", "CreateDirectoryExW", _
 "RemoveDirectoryA", "RemoveDirectoryW", "DeleteFileA", "DeleteFileW", "CopyFileA", "CopyFileW", _
 "CopyFileExA", "CopyFileExW", "MoveFileA", "MoveFileW", "MoveFileExA", "MoveFileExW", "UnmapViewOfFile", _
 "GetVolumeInformationA", "GetVolumeInformationW", "GetDriveTypeA", "GetDriveTypeW", _
 "GetLogicalDrives", "GetLogicalDriveStringsA", "GetLogicalDriveStringsW"
 
 

    If ConfigData(7) = 0 Then
   AddBreakPoint SPYBREAKPOINTS, ExPs.FuncAddress(CounterLoop), 1
   Else
   AddBreakPoint SPYBREAKPOINTS, ExPs.FuncAddress(CounterLoop), 0
   End If

End Select

ElseIf StrComp(UCase(MName), "ADVAPI32.DLL") = 0 Then

 Select Case ExPs.FuncNames(CounterLoop)

Case Is = "RegOpenKeyA", "RegOpenKeyW", "RegOpenKeyExA", "RegOpenKeyExW", _
"RegCloseKey", "RegCreateKeyA", "RegCreateKeyW", "RegCreateKeyExA", "RegCreateKeyExW", _
"RegDeleteKeyA", "RegDeleteKeyW", "RegQueryValueA", "RegQueryValueW", _
"RegQueryValueExA", "RegQueryValueExW", "RegSetValueA", "RegSetValueW", "RegSetValueExA", "RegSetValueExW", _
"RegDeleteValueA", "RegDeleteValueW", "RegLoadKeyA", "RegLoadKeyW", "RegSaveKeyA", "RegSaveKeyW", _
"RegConnectRegistryA", "RegConnectRegistryW"
   If ConfigData(7) = 0 Then
   AddBreakPoint SPYBREAKPOINTS, ExPs.FuncAddress(CounterLoop), 1
   Else
   AddBreakPoint SPYBREAKPOINTS, ExPs.FuncAddress(CounterLoop), 0
   End If
End Select





    
ElseIf StrComp(UCase(MName), "USER32.DLL") = 0 Then
   
'  Select Case ExPs.FuncNames(CounterLoop)
    
' Case Is = "SetTimer", "KillTimer" _
' , "LoadStringA", "LoadStringW"
 

'    If ConfigData(7) = 0 Then
'   AddBreakPoint SPYBREAKPOINTS, ExPs.FuncAddress(CounterLoop), 1
'   Else
'   AddBreakPoint SPYBREAKPOINTS, ExPs.FuncAddress(CounterLoop), 0
'End If
   
'End Select



End If

End Sub
Public Sub DeleteInExportsSearch(ByVal MName As String)
On Error GoTo Dalje
Dim COld As Collection
Set COld = ModulesExports.Item(MName)
Set COld = Nothing
ModulesExports.Remove (MName)
Exit Sub
Dalje:
On Error GoTo 0
End Sub

Public Function GetFromExportsSearch(ByVal MName As String, ByVal Address As Long) As String
On Error GoTo Dalje
Dim COld As Collection
Set COld = ModulesExports.Item(MName)
GetFromExportsSearch = COld("X" & Address)
Exit Function
Dalje:
On Error GoTo 0
End Function
Public Sub AddToCols(COL As Collection, ByRef DataL As Long, ByRef DataS As String)
On Error GoTo Dalje
COL.Add DataS, "X" & DataL
Exit Sub
Dalje:
On Error GoTo 0
Dim S2 As String
S2 = COL.Item("X" & DataL)
If S2 = DataS Then Exit Sub
S2 = S2 & "," & DataS
COL.Remove "X" & DataL
COL.Add S2, "X" & DataL
End Sub


