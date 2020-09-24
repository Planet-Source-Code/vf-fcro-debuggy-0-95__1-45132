VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form20 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "AVI Resource"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6390
   LinkTopic       =   "Form20"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   5760
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5760
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   5760
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
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
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5385
      ScaleWidth      =   6345
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00DCB17C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Size:"
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
      TabIndex        =   2
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private IsPSD As Byte
Private IsSign As Byte
Private FNamUse As String


Private Sub Command2_Click()
Dim ret As Long
Dim Widt As Long
Dim Heig As Long

CloseMedia
ret = OpenMedia(Picture1)

If ret <> 0 Then MsgBox "Cannot Display movie!", vbCritical, "Error": Exit Sub

GetSize Widt, Heig
If IsSign = 0 Then Label7 = "Size Width:" & Widt & ",Height:" & Heig
IsSign = 1

If Widt < 425 Then Picture1.Width = Widt * 15
If Heig < 425 Then Picture1.Height = Heig * 15
'SetInWindow Picture1
'PlayMedia


SetInWindow Picture1
PlayMedia

IsPSD = 1
End Sub



Private Sub Command4_Click()
cd1.ShowSave
If Len(cd1.Filename) = 0 Then Exit Sub
If 0 = CopyFile(FNamUse, cd1.Filename, 0) Then MsgBox "Unable to Save File", vbCritical, "Error!"
End Sub

Private Sub Command5_Click()
On Error Resume Next
StopMedia
CloseMedia
Kill FNamUse
If Err <> 0 Then On Error GoTo 0: MsgBox "Cannot erase temporary media file!", vbCritical, "Error"
Unload Me
End Sub

Private Sub Form_Load()
RemoveX hWnd
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
IsPSD = 2
IsSign = 0
FNamUse = AddSlash(App.Path) & "testAvi.Avi"
MMFile = """" & FNamUse & """"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'StopMedia
'CloseMedia
'Kill FNamUse
End Sub
