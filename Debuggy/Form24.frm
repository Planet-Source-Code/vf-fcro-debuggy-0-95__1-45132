VERSION 5.00
Begin VB.Form Form24 
   Caption         =   "Watch"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form24.frx":0000
   LinkTopic       =   "Form24"
   MaxButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DACCC2&
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text3 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   960
      MaxLength       =   8
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00E9DDDA&
      Height          =   345
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Address="
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   2880
      Width           =   855
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E7DFD6&
      Height          =   2760
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   2880
      Width           =   855
   End
End
Attribute VB_Name = "Form24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error GoTo Dalje
If Len(Text3) = 0 Then MsgBox "An empty address", vbCritical, "Error": Exit Sub
Dim XAdrr As Long
Dim Xtyp As Long
XAdrr = CLng("&H" & Text3)
If Combo1.ListIndex = -1 Or Len(Combo1.Text) = 0 Then MsgBox "Select Type of Expression!", vbExclamation, "Require": Exit Sub
Xtyp = Combo1.ItemData(Combo1.ListIndex)
Dim IsAccept As Long
Dim SLens As String



AddInWatches XAdrr, Xtyp, IsAccept
If IsAccept = 0 Then MsgBox "Expression already exist!", vbInformation, "Information"

ReadExpresses
Exit Sub

Dalje:
On Error GoTo 0
MsgBox "Error in Value", vbCritical, "Error"
End Sub

Private Sub Command3_Click()
If List2.ListCount = 0 Or List2.ListIndex = -1 Then MsgBox "Select Expression First", vbInformation, "Require": Exit Sub
Dim FrmLB() As String
FrmLB = Split(List2.List(List2.ListIndex), vbTab)
Dim Fadd() As String
Fadd = Split(FrmLB(0), ":")
RemoveInWatches CLng("&H" & Fadd(1)), CLng(List2.ItemData(List2.ListIndex))
Erase FrmLB
Erase Fadd
'ReadExpresses
List2.RemoveItem List2.ListIndex
End Sub

Private Sub Form_Load()
RemoveX hWnd
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
'Call SendMessage(List2.hWnd, &H194, ByVal 2000, ByVal 0&)
Combo1.Clear
Combo1.AddItem "Byte"
Combo1.ItemData(Combo1.ListCount - 1) = 0
Combo1.AddItem "Word"
Combo1.ItemData(Combo1.ListCount - 1) = 1
Combo1.AddItem "Dword"
Combo1.ItemData(Combo1.ListCount - 1) = 2
Combo1.AddItem "Quad"
Combo1.ItemData(Combo1.ListCount - 1) = 3
Combo1.AddItem "Data Block 16"
Combo1.ItemData(Combo1.ListCount - 1) = 4


Combo1.AddItem "Byte PTR"
Combo1.ItemData(Combo1.ListCount - 1) = 10
Combo1.AddItem "Word PTR"
Combo1.ItemData(Combo1.ListCount - 1) = 11
Combo1.AddItem "Dword PTR"
Combo1.ItemData(Combo1.ListCount - 1) = 12
Combo1.AddItem "Quad PTR"
Combo1.ItemData(Combo1.ListCount - 1) = 13



End Sub



Private Sub Text3_KeyPress(KeyAscii As Integer)
On Error GoTo Dalje
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 8 Then Exit Sub
If IsValidK(Chr(KeyAscii)) = 0 Then KeyAscii = 0
Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Unknown Value Type", vbCritical, "Error"
End Sub

Public Sub ReadExpresses()

List2.Clear
Dim u As Long
Dim Typs() As Long
Dim ExpEx1(1) As String
Dim IsValid As Long



'0-adr ,1-type ,2-PTR/Value ,3-Value
For u = 1 To WATCHES.count

Typs = WATCHES.Item(u)

Select Case Typs(1)

Case 0
ExpEx1(0) = "Address:" & Hex(Typs(0))
ExpEx1(1) = "BYTE=" & ReadAsByte(Typs(0), IsValid)

Case 1
ExpEx1(0) = "Address:" & Hex(Typs(0))
ExpEx1(1) = "WORD=" & ReadAsWord(Typs(0), IsValid)


Case 2
ExpEx1(0) = "Address:" & Hex(Typs(0))
ExpEx1(1) = "DWORD=" & ReadAsDword(Typs(0), IsValid)


Case 3
ExpEx1(0) = "Address:" & Hex(Typs(0))
ExpEx1(1) = "QUAD=" & ReadAsQuad(Typs(0), IsValid)

Case 4
ExpEx1(0) = "Address:" & Hex(Typs(0))
ExpEx1(1) = "DATA BLOCK 16=" & GetDBBLock(Typs(0))


Case 10
ExpEx1(0) = "Address:" & Hex(Typs(0))
ExpEx1(1) = "PTR ON BYTE=" & ReadAsBytePTR(Typs(0), IsValid)

Case 11
ExpEx1(0) = "Address:" & Hex(Typs(0))
ExpEx1(1) = "PTR ON WORD=" & ReadAsWordPTR(Typs(0), IsValid)

Case 12
ExpEx1(0) = "Address:" & Hex(Typs(0))
ExpEx1(1) = "PTR ON DWORD=" & ReadAsDwordPTR(Typs(0), IsValid)

Case 13
ExpEx1(0) = "Address:" & Hex(Typs(0))
ExpEx1(1) = "PTR ON QUAD=" & ReadAsQuadPTR(Typs(0), IsValid)


End Select

List2.AddItem Join(ExpEx1, vbTab)
List2.ItemData(List2.ListCount - 1) = Typs(1)
Erase ExpEx1
Next u



End Sub

