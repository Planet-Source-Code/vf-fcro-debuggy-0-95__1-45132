VERSION 5.00
Begin VB.Form Form32 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Type Library Info"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11790
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form32"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   10800
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DACCC2&
      Height          =   7470
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   11775
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00DACCC2&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Information"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   11775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00DACCC2&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Type Library Information"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "Form32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TLINF As TypeLibInfo



Private Sub Combo1_Click()
On Error GoTo Dalje
If Combo1.ListCount <> 0 And Combo1.ListIndex <> -1 Then

Select Case Combo1.ListIndex

Case 0
If TLINF.CoClasses.count <> 0 Then
SpeedUpAdding List1.hWnd, 5000, 100000
Label1 = "Information about CoClasses"
List1.Clear: ReadClass TLINF.CoClasses, List1
End If

Case 1
If TLINF.Interfaces.count <> 0 Then
SpeedUpAdding List1.hWnd, 5000, 100000
Label1 = "Information about Interfaces"
List1.Clear: ReadInterfaces TLINF.Interfaces, List1
End If

Case 2
If TLINF.Interfaces.count <> 0 Then
SpeedUpAdding List1.hWnd, 5000, 100000
Label1 = "Information about Type Definition"
List1.Clear: ReadTypes TLINF.Records, List1
End If

Case 3
If TLINF.Declarations.count <> 0 Then
SpeedUpAdding List1.hWnd, 5000, 100000
Label1 = "Information about Modules"
List1.Clear: ReadDeclarations TLINF.Declarations, List1
End If

Case 4
If TLINF.Constants.count <> 0 Then
SpeedUpAdding List1.hWnd, 5000, 100000
Label1 = "Information about Enumerations"
List1.Clear: ReadConstants TLINF.Constants, List1
End If

Case 5
If TLINF.Unions.count <> 0 Then
SpeedUpAdding List1.hWnd, 5000, 100000
Label1 = "Information about Unions"
List1.Clear: ReadUnions TLINF.Unions, List1
End If

End Select




End If

Exit Sub

Dalje:
On Error GoTo 0
MsgBox "Error while reading Type Library!", vbCritical, "Error"
End Sub

Public Sub StartRead()
On Error GoTo Dalje


List1.Clear
Combo1.Clear


Combo1.AddItem "CoClasses [" & TLINF.CoClasses.count & "]"

Combo1.AddItem "Interfaces [" & TLINF.Interfaces.count & "]"

Combo1.AddItem "DefTypes [" & TLINF.Records.count & "]"

Combo1.AddItem "Modules [" & TLINF.Declarations.count & "]"

Combo1.AddItem "Enums [" & TLINF.Constants.count & "]"

Combo1.AddItem "Unions [" & TLINF.Unions.count & "]"



Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Not an Type Library File", vbExclamation, "Information"
End Sub


Private Sub ReadUnions(UNS As Unions, LB As ListBox)
Dim u As Long
For u = 1 To UNS.count
LB.AddItem UNS.Item(u).Name
ReadMembers2 UNS.Item(u).Members, LB, vbTab
LB.AddItem ""
Next u

End Sub



Private Sub ReadConstants(CNST As tli.Constants, LB As ListBox)
Dim u As Long
For u = 1 To CNST.count
LB.AddItem CNST.Item(u).Name
ReadMembers2 CNST.Item(u).Members, LB, vbTab, 1
LB.AddItem ""
Next u

End Sub


Private Sub ReadDeclarations(DCL As Declarations, LB As ListBox)
Dim u As Long

For u = 1 To DCL.count
LB.AddItem DCL.Item(u).Name

ReadMembers DCL.Item(u).Members, LB, "              "
LB.AddItem ""
Next u

End Sub


Private Sub ReadTypes(RCD As Records, LB As ListBox)
Dim u As Long
Dim i As Long
If RCD.count = 0 Then Exit Sub


For u = 1 To RCD.count
LB.AddItem RCD.Item(u).Name


ReadMembers2 RCD.Item(u).Members, LB, vbTab

LB.AddItem ""

Next u

End Sub

Private Sub ReadClass(CLX As CoClasses, LB As ListBox)
Dim u As Long
Dim i As Long

If CLX.count = 0 Then Exit Sub


Dim OUTS As String
For u = 1 To CLX.count

OUTS = CLX.Item(u).Name & vbTab & CLX.Item(u).GUID
LB.AddItem OUTS

ReadInterfaces CLX.Item(u).Interfaces, List1, , 1, 1, vbTab, 1

LB.AddItem ""
Next u

End Sub
Private Sub ReadInterfaces(INFS As Interfaces, LB As ListBox, Optional ByRef NameExc As Byte, Optional ByRef GuidExc As Byte, Optional ByRef KindExc As Byte, Optional ByRef Prestring As String, Optional ByRef SkipReadMembers As Byte)
On Error GoTo Dalje
Dim u As Long

If INFS.count = 0 Then Exit Sub



Dim OUTS As String
For u = 1 To INFS.count

OUTS = ""
If NameExc = 0 Then OUTS = OUTS & INFS.Item(u).Name & vbTab

If GuidExc = 0 Then OUTS = OUTS & INFS.Item(u).GUID & vbTab

If KindExc = 0 Then OUTS = OUTS & INFS.Item(u).TypeKindString

If Right(OUTS, 1) = vbTab Then
LB.AddItem Prestring & Left(OUTS, Len(OUTS) - 1)
Else
LB.AddItem Prestring & OUTS
End If

If SkipReadMembers = 0 Then
ReadMembers INFS.Item(u).Members, LB, "              "
LB.AddItem ""
End If

Next u

Exit Sub
Dalje:
On Error GoTo 0
End Sub
Private Sub ReadMembers(MEMB As Members, LB As ListBox, ByRef Prestring As String)
On Error GoTo Dalje
Dim u As Long
Dim OUTS As String
Dim OUTSP As String
Dim RTTS As String
For u = 1 To MEMB.count
OUTS = GetKindString(MEMB.Item(u).InvokeKind) & vbTab & MEMB.Item(u).Name
OUTSP = ReadParams(MEMB.Item(u).Parameters)
RTTS = GetTypeString(MEMB.Item(u).ReturnType)

LB.AddItem Prestring & OUTS & vbTab & OUTSP & "  ,Ret:" & RTTS
Next u



Exit Sub
Dalje:
On Error GoTo 0
End Sub
Private Sub ReadMembers2(MEMB As Members, LB As ListBox, ByRef Prestring As String, Optional ByRef SkipReadType As Byte)
On Error GoTo Dalje
Dim u As Long
Dim OUTS As String
For u = 1 To MEMB.count
OUTS = Prestring & MEMB.Item(u).Name
If SkipReadType = 0 Then
OUTS = OUTS & ":" & GetTypeString(MEMB.Item(u).ReturnType)
Else
OUTS = OUTS & "=" & Hex(MEMB.Item(u).Value)
End If

LB.AddItem OUTS
Next u
Exit Sub
Dalje:
On Error GoTo 0
End Sub


Private Function ReadParams(PRM As Parameters) As String
On Error GoTo Dalje
Dim u As Long
ReadParams = "("
Dim OUTS As String
Dim OUTSP As String
For u = 1 To PRM.count
ReadParams = ReadParams & PRM.Item(u).Name
ReadParams = ReadParams & ":" & GetTypeString(PRM.Item(u).VarTypeInfo)
If u <> PRM.count Then ReadParams = ReadParams & " ,"
Next u
ReadParams = ReadParams & ")"
Exit Function
Dalje:
On Error GoTo 0
ReadParams = "[ERROR WHILE READING PARAMETERS]"
End Function

Private Function GetTypeString(TLL As VarTypeInfo) As String

If TLL.IsExternalType Then
If Len(TLL.TypeLibInfoExternal.Name) <> 0 Then
 GetTypeString = GetTypeString & TLL.TypeLibInfoExternal.Name & "."
End If
End If


Dim NMS As Long
NMS = TLL
If NMS >= 8192 Then
NMS = NMS - 8192: GetTypeString = GetTypeString & "SAFEARRAY "
ElseIf NMS >= 16384 Then
NMS = NMS - 16384: GetTypeString = GetTypeString & "REFERENCE "
ElseIf NMS >= 4096 Then
NMS = NMS - 4096: GetTypeString = GetTypeString & "VECTOR "
End If

Select Case NMS
Case 0
GetTypeString = GetTypeString & TLL.TypeInfo.Name
Case tli.VT_ARRAY
GetTypeString = GetTypeString & "VT_ARRAY"
Case tli.VT_I1
GetTypeString = GetTypeString & "VT_I1"
Case tli.VT_I2
GetTypeString = GetTypeString & "VT_I2"
Case tli.VT_I4
GetTypeString = GetTypeString & "VT_I4"
Case tli.VT_BSTR
GetTypeString = GetTypeString & "VT_BSTR"
Case tli.VT_BOOL
GetTypeString = GetTypeString & "VT_BOOL"
Case tli.VT_CLSID
GetTypeString = GetTypeString & "VT_CLSID"
Case tli.VT_CY
GetTypeString = GetTypeString & "VT_CY"
Case tli.VT_I1
GetTypeString = GetTypeString & "VT_I1"
Case tli.VT_I2
GetTypeString = GetTypeString & "VT_I2"
Case tli.VT_I4
GetTypeString = GetTypeString & "VT_I4"
Case tli.VT_I8
GetTypeString = GetTypeString & "VT_I8"
Case tli.VT_LPSTR
GetTypeString = GetTypeString & "VT_LPSTR"
Case tli.VT_LPWSTR
GetTypeString = GetTypeString & "VT_LPWSTR"
Case tli.VT_DATE
GetTypeString = GetTypeString & "VT_DATE"
Case tli.VT_R4
GetTypeString = GetTypeString & "VT_R4"
Case tli.VT_R8
GetTypeString = GetTypeString & "VT_R8"
Case tli.VT_UINT
GetTypeString = GetTypeString & "VT_UINT"
Case tli.VT_UI1
GetTypeString = GetTypeString & "VT_UI1"
Case tli.VT_UI2
GetTypeString = GetTypeString & "VT_UI2"
Case tli.VT_UI4
GetTypeString = GetTypeString & "VT_UI4"
Case tli.VT_UI8
GetTypeString = GetTypeString & "VT_UI8"
Case tli.VT_VOID
GetTypeString = GetTypeString & "VT_VOID"
Case tli.VT_INT
GetTypeString = GetTypeString & "VT_INT"
Case tli.VT_NULL
GetTypeString = GetTypeString & "VT_NULL"
Case tli.VT_PTR
GetTypeString = GetTypeString & "VT_PTR"
Case tli.VT_VARIANT
GetTypeString = GetTypeString & "VT_VARIANT"
Case tli.VT_DISPATCH
GetTypeString = GetTypeString & "VT_DISPATCH"
Case tli.VT_UNKNOWN
GetTypeString = GetTypeString & "VT_UNKNOWN"
Case tli.VT_SAFEARRAY
GetTypeString = GetTypeString & "VT_SAFEARRAY"
Case tli.VT_HRESULT
GetTypeString = GetTypeString & "VT_HRESULT"
Case tli.VT_ERROR
GetTypeString = GetTypeString & "VT_ERROR"
Case tli.VT_DECIMAL
GetTypeString = GetTypeString & "VT_DECIMAL"
Case Else
GetTypeString = "[Unresolved Type]"

End Select


End Function

Private Function GetKindString(TLL As Long) As String

Select Case TLL

Case tli.INVOKE_FUNC
GetKindString = "Method"
Case tli.INVOKE_EVENTFUNC
GetKindString = "Event"
Case tli.INVOKE_PROPERTYGET
GetKindString = "Prop Get"
Case tli.INVOKE_PROPERTYPUT
GetKindString = "Prop Put"
Case tli.INVOKE_PROPERTYPUTREF
GetKindString = "Prop Put Ref"
Case tli.INVOKE_UNKNOWN
GetKindString = "Prop"
Case tli.INVOKE_CONST
GetKindString = "Const"
'Case tli.INVOKE_PROPERTYPUT Or tli.INVOKE_PROPERTYGET
'GetKindString = "Property Get/Put"
'Case tli.INVOKE_PROPERTYPUTREF Or tli.INVOKE_PROPERTYGET
'GetKindString = "Property Get/Put Ref"
'Case tli.INVOKE_PROPERTYPUTREF Or tli.INVOKE_PROPERTYPUT
'GetKindString = "Property Put/Put Ref"
'Case tli.INVOKE_PROPERTYGET Or tli.INVOKE_PROPERTYPUTREF Or tli.INVOKE_PROPERTYPUT
'GetKindString = "Property Get/Put/Put Ref"

End Select


End Function







Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Combo1 = ""
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Dim Tabs() As Long
ReDim Tabs(1)
Tabs(0) = 0
Tabs(1) = 85

Call SendMessage(List1.hWnd, &H192, ByVal UBound(Tabs) + 1, Tabs(0))


Call SendMessage(List1.hWnd, &H194, ByVal 4000, ByVal 0&)



End Sub


