VERSION 5.00
Begin VB.Form Form22 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dialog Box Resource"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form22"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
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
      Height          =   2760
      Left            =   0
      TabIndex        =   4
      Top             =   5400
      Width           =   8535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   8280
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   8505
      TabIndex        =   0
      Top             =   240
      Width           =   8535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E7DFD6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dialog Childs Informations"
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
      TabIndex        =   3
      Top             =   5160
      Width           =   8535
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00DCB17C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dialog Box"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WHDL As Long

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Call SendMessage(List1.hWnd, &H194, ByVal 6000, ByVal 0&)
SetParent WHDL, Picture1.hWnd
Dim Wsst As Long
Wsst = GetWindowLong(WHDL, -16)
Wsst = Wsst Or &H80000000
SetWindowLong WHDL, -16, Wsst
SetWindowPos WHDL, 0, 0, 0, 0, 0, 1
ShowWindow WHDL, 1
List1.Clear
LockWindowUpdate List1.hWnd
EnumChildWindows WHDL, AddressOf EnumChildRS, 0
LockWindowUpdate 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
DestroyWindow WHDL
End Sub
