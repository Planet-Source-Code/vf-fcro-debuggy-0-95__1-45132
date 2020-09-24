VERSION 5.00
Begin VB.Form SPBREAK 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "API Spy"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form30"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Break on result"
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continue"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Break Now"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E7DFD6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "API Encounted"
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
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "SPBREAK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LASTTHREADID As Long
Public IFREF As Long

Private RetA As Long
Private VTyp As Long
Private CustomD As String


Private Sub Command1_Click()
Dim RCF As Long
RCF = 1
CopyMemory ByVal IFREF, RCF, 4
RETCONFIRMATION = 1
Unload Me
End Sub

Private Sub Command2_Click()
Dim RCF As Long
RCF = 0
CopyMemory ByVal IFREF, RCF, 4



Unload Me
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
RemoveX hWnd
End Sub


Public Sub ReReadF(Data() As String, XData() As Long, ByRef ThreadId As Long, ByRef PROCRF As Long, ByRef RRadr As Long, ByRef RRtype As Long, ByRef RRcustom As String)
Dim u As Long
Check1.Value = 1
Text4 = ""
For u = 0 To UBound(Data)
Text4 = "In Thread:" & ThreadId & vbCrLf & _
Join(Data, vbCrLf)
Next u
IFREF = VarPtr(PROCRF)
RetA = RRadr
VTyp = RRtype
CustomD = RRcustom
End Sub

Private Sub Form_Unload(Cancel As Integer)
AddBPX RetA, VTyp, CustomD, CByte(Check1.Value)
End Sub
