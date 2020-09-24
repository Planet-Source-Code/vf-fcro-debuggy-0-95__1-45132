VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modules in Process"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7125
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Find Hidden Module"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Examine"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   4680
      Width           =   1095
   End
   Begin VB.ListBox List7 
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
      Height          =   4350
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00DCB17C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Modules in Process/Virtual Address/Length/Entry Point"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
Unload Me
End Sub



Private Sub Command2_Click()
List7_dblClick
End Sub





Private Sub Command3_Click()
HModules.Show 1
ReadModules List7
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2

Call SendMessage(List7.hWnd, &H194, ByVal 1000, ByVal 0&)
Dim Tabs(3) As Long
Tabs(0) = 150
Tabs(1) = 50
Tabs(2) = 60
Tabs(3) = 70
Call SendMessage(List7.hWnd, &H192, ByVal UBound(Tabs) + 1, Tabs(0))
ReadModules List7
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo dalje
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 8 Then
Exit Sub
ElseIf KeyAscii = 13 Then
If Len(Text1) = 0 Then Text1 = "": Exit Sub
DISCOUNT = CLng("&H" & Text1): NextB = 0: Unload Me
End If

If IsValidK(Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
Exit Sub
dalje:
On Error GoTo 0
MsgBox "Unknown Value Type", vbCritical, "Error"
End Sub


Private Sub List7_dblClick()
If List7.ListCount = 0 Or List7.ListIndex = -1 Then MsgBox "Select Module Fist", vbInformation, "Require": Exit Sub

Dim SxND() As String
SxND = Split(List7.List(List7.ListIndex), vbTab)
Form6.ModuleToShow = CLng("&H" & SxND(1))
Form6.Caption = "Imports/Exports by Module:" & SxND(0)
Form6.Show 1
Form16.ReleaseShow 1

End Sub

Private Sub List7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo dalje
Dim IDX As Long
IDX = IndexFromPoint(List7.hWnd, X / 15, Y / 15)

If IDX <> -1 Then
Dim PTH As String
Dim SxND() As String
SxND = Split(List7.List(IDX), vbTab)
List7.ToolTipText = GetModulePath(CLng("&H" & SxND(1)))
Erase SxND
End If
Exit Sub

dalje:
On Error GoTo 0
End Sub
