VERSION 5.00
Begin VB.Form Form18 
   Caption         =   "Call Stack"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5430
   Icon            =   "Form18.frx":0000
   LinkTopic       =   "Form18"
   MaxButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
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
      Height          =   2565
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E7DFD6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Call Stack"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RCTX As CONTEXT
Public ShowingTH As Long

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2

RemoveX hWnd
Call SendMessage(List2.hWnd, &H194, ByVal 2000, ByVal 0&)

Dim Tabs() As Long
ReDim Tabs(1)
Tabs(0) = 0
Tabs(1) = 40
Call SendMessage(List2.hWnd, &H192, ByVal UBound(Tabs) + 1, Tabs(0))



End Sub






Public Sub ReadIt()
If ShowingTH = 0 Then ShowingTH = ActiveThread

OnScreen hWnd
Dim ThrS As Long
Dim HT As Long
HT = GetHandleOfThread(ShowingTH, ThrS)
CheckThreadCaption HT, ThrS
If HT = 0 Or ThrS = 1 Then Exit Sub
RCTX = GetContext(ShowingTH)
Caption = Caption & " ,Frame on EIP:" & Hex(RCTX.Eip)
ReadStackFrame List2, RCTX.Eip, RCTX.Ebp
End Sub


Private Sub CheckThreadCaption(ByRef HT As Long, ByRef ThrS As Long)
If ThrS = 1 Then
Caption = "Thread:" & ShowingTH & " ,Running!"
List2.Clear
ElseIf HT = 0 Then
Caption = "Thread:" & ShowingTH & " ,Not valid now!"
List2.Clear
Else
Caption = "Thread:" & ShowingTH
End If
End Sub



Private Sub List2_dblClick()
If List2.ListCount = 0 Or List2.ListIndex = -1 Then Exit Sub
Dim RAdr() As String
RAdr = Split(List2.List(List2.ListIndex), vbTab)
DISCOUNT = CLng("&H" & RAdr(0)): NextB = 0
Form16.ReleaseShow 1
End Sub
