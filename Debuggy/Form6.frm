VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Imports/Exports"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11550
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Ole / Com"
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Set Breakpoints"
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Select All"
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PE Header"
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Resources"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   7800
      Width           =   1095
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
      ForeColor       =   &H00E7DFD6&
      Height          =   5310
      Left            =   0
      TabIndex        =   6
      Top             =   2400
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   9120
      TabIndex        =   5
      Top             =   7800
      Width           =   1095
   End
   Begin VB.ListBox List3 
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
      ForeColor       =   &H00E7DFD6&
      Height          =   7470
      Left            =   5760
      MultiSelect     =   1  'Simple
      TabIndex        =   3
      Top             =   240
      Width           =   5775
   End
   Begin VB.ListBox List2 
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
      ForeColor       =   &H00E7DFD6&
      Height          =   1950
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E7DFD6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exports Functions / Addresses"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00DCB17C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Imports By Module:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   5775
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CA9273&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Imports Modules"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ModuleToShow As Long
Private ImpX As Byte
Private ExpX As Byte
Private Sub Command1_Click()
NextB = 0: Unload Me
End Sub



Private Sub Command2_Click()
If CanEnum = 0 Then MsgBox "Cannot enumerate resources on BREAKPOINT!", vbExclamation, "Information": Exit Sub


Dim retR As Byte
retR = EnumIt(ModuleToShow)
If retR = 0 Then Exit Sub
Form1.ShowHinstance = ModuleToShow
Form1.Show 1
End Sub

Private Sub Command3_Click()
Form25.Show 1
End Sub

Private Sub Command4_Click()
If ExpX = 0 Then MsgBox "No Exports for this module", vbExclamation, "Information": Exit Sub

If List3.ListCount = 0 Then MsgBox "An empty Exports list!", vbExclamation, "Information": Exit Sub
ClearSelected List3.hWnd
End Sub

Private Sub Command5_Click()
If ExpX = 0 Then MsgBox "No Exports for this module", vbExclamation, "Information": Exit Sub

If List3.ListCount = 0 Then MsgBox "An empty Exports list!", vbExclamation, "Information": Exit Sub
SelectRange List3.hWnd, 0, List3.ListCount - 1
End Sub

Private Sub Command6_Click()
If ExpX = 0 Then MsgBox "No Exports for this module", vbExclamation, "Information": Exit Sub


Dim Itms() As Long

Dim PiRef As Byte
Dim ItmsCount As Long
ItmsCount = GetSelectedItems(List3.hWnd, Itms)
If ItmsCount = 0 Then MsgBox "Nothing selected yet.", vbInformation, "Information": Exit Sub

Dim AdrR As Long
Dim Vuz() As String
Dim u As Long
For u = 0 To ItmsCount - 1
Vuz = Split(List3.List(Itms(u)), vbTab)
AdrR = CLng("&H" & Vuz(1))
Dim IsValidBP As Byte
Dim IsValidMemPTR As Byte
Dim BTTX As Byte
IsValidMemPTR = TestPTR(AdrR, BTTX)
If IsValidMemPTR = 0 Then PiRef = 1: GoTo Dalje

Call GetBreakPoint(SPYBREAKPOINTS, AdrR, IsValidBP)
If IsValidBP = 1 Then PiRef = 1: GoTo Dalje

Call GetBreakPoint(ACTIVEBREAKPOINTS, AdrR, IsValidBP)
If IsValidBP = 0 Then
AddBreakPoint ACTIVEBREAKPOINTS, AdrR, ISBPDisabled
Call SendMessage(List3.hWnd, &H185, ByVal 0&, ByVal Itms(u))
Else
PiRef = 1
End If

Dalje:
Next u
Erase Itms
Erase Vuz
NextB = 0
Form16.ReleaseShow 1
PrintDump Form8.TextX, ActiveMemPos
If PiRef = 1 Then
MsgBox "Cannot set Breakpoints on all of selected address!", vbExclamation, "Information"
Else
MsgBox "Set all Breakpoints!", vbInformation, "Information"
End If
End Sub

Private Sub Command7_Click()
On Error GoTo Dalje
Dim PTH As String
PTH = GetModulePath(ModuleToShow)
Set Form32.TLINF = tli.TypeLibInfoFromFile(PTH)

Form32.StartRead
Form32.Show 1
Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Module wasn't OLE/COM or doesn't contain Type Library " & vbCrLf & "(Search in the Resources for TYPELIB).", vbExclamation, "Information"
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Call SendMessage(List1.hWnd, &H194, ByVal 1000, ByVal 0&)
Call SendMessage(List3.hWnd, &H194, ByVal 1000, ByVal 0&)

Dim Tabs() As Long
ReDim Tabs(1)
Tabs(0) = 180
Tabs(1) = 220

Call SendMessage(List1.hWnd, &H192, ByVal UBound(Tabs) + 1, Tabs(0))
Call SendMessage(List3.hWnd, &H192, ByVal UBound(Tabs) + 1, Tabs(0))


ReadPE2 ModuleToShow, ImpX, ExpX
Dim u As Long

List1.Clear
List2.Clear
List3.Clear

If ImpX = 0 Then
List2.AddItem "No Imports!"
Else

For u = 0 To UBound(IMPS)
List2.AddItem IMPS(u).Module

Next u
End If


If ExpX = 0 Then
List3.AddItem "No Exports!"
Else

For u = 0 To UBound(ExPs.FuncNames)
If ExPs.FuncAddress(u) <> 0 Then
'List3.AddItem EXPS.FuncNames(u) & vbTab & "Ord:" & EXPS.Ord(u) & vbTab & Hex(EXPS.FuncAddress(u))
List3.AddItem ExPs.FuncNames(u) & vbTab & Hex(ExPs.FuncAddress(u))
End If
Next u
End If



End Sub

Private Sub List1_dblClick()
On Error GoTo Dalje
If List1.ListCount = 0 Or List1.ListIndex = -1 Then Exit Sub
Dim RLer() As String
RLer = Split(List1.List(List1.ListIndex), vbTab)
Dim GAddress As Long
GAddress = CLng("&H" & RLer(1))
ChoosedAdr = GAddress:  Unload Me
Exit Sub
Dalje:
On Error GoTo 0
End Sub

Private Sub List2_dblClick()
Dim KLI As Long
Dim Xadr As String
KLI = List2.ListIndex
If (List2.ListCount = 1 And List2.List(0) = "No Imports!") Or KLI = -1 Then Exit Sub
Dim i As Long
List1.Clear
Label1 = "Imports By Module:" & IMPS(KLI).Module
For i = 0 To UBound(IMPS(KLI).FuncNames)
'List1.AddItem IMPS(KLI).FuncNames(i) & vbTab & "Ord:" & IMPS(KLI).Ord(i) & vbTab & Hex(IMPS(KLI).CallingAddresses(i))


If TestPTR(IMPS(KLI).CallingAddresses(i)) = 0 Then
Xadr = "Not Loaded Yet By ntdll.dll"
Else
Xadr = Hex(IMPS(KLI).CallingAddresses(i))
End If
List1.AddItem IMPS(KLI).FuncNames(i) & vbTab & Xadr
Next i

End Sub

Private Sub List3_dblClick()
On Error GoTo Dalje
If ExpX = 0 Or List3.ListIndex = -1 Then Exit Sub
Dim RLer() As String
RLer = Split(List3.List(List3.ListIndex), vbTab)
Dim GAddress As Long
GAddress = CLng("&H" & RLer(1))
ChoosedAdr = GAddress:  Unload Me
Exit Sub
Dalje:
On Error GoTo 0
End Sub
