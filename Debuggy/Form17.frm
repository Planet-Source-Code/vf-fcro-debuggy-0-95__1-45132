VERSION 5.00
Begin VB.Form Form17 
   Caption         =   "B/S"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   Icon            =   "Form17.frx":0000
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Goto EXP"
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
      Left            =   1800
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
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
      Left            =   2880
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text1 
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
      ForeColor       =   &H00E7DFD6&
      Height          =   1455
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   5175
   End
   Begin VB.VScrollBar vs1 
      Height          =   1455
      Left            =   5160
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E7DFD6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Base/Stack"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WTypeS As Byte
'0-EBP
'1-ESP

Public ShowingTH As Long 'Which Thread shows=?
Private LastState As Byte
Private LastEXP As Long
Private RCTX As CONTEXT
Private IsS As Byte





Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
ReadIt
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2

Dim Tabs() As Long
ReDim Tabs(1)
Tabs(0) = 0
Tabs(1) = 45
Call SendMessage(Text1.hWnd, &HCB, ByVal UBound(Tabs) + 1, Tabs(0))

If WTypeS = 0 Then
Command2.Caption = "Goto EBP"
Else
Command2.Caption = "Goto ESP"
End If

RemoveX hWnd
vs1.Max = 32767
vs1.Min = 0
vs1.VALUE = 16384
vs1.SmallChange = 1
vs1.LargeChange = 10
ReadIt
End Sub

Public Sub ReadIt(Optional ByVal ReRead As Byte)
If ShowingTH = 0 Then ShowingTH = ActiveThread

Dim ThrS As Long
Dim HT As Long
HT = GetHandleOfThread(ShowingTH, ThrS)
LastState = ThrS
CheckThreadCaption HT, ThrS
If HT = 0 Or ThrS = 1 Then Exit Sub
RCTX = GetContext(ShowingTH)
If WTypeS = 0 Then
If ReRead = 0 Then
ActiveBasePosition = RCTX.Ebp
End If
LastEXP = RCTX.Ebp

Read9Stack Text1, ActiveBasePosition, ActiveBasePosition, "EBP"
Label14 = "BASE PTR At:" & Hex(RCTX.Ebp)
Command2.Caption = "Goto EBP"
Else
If ReRead = 0 Then
ActiveStackPosition = RCTX.Esp
End If
LastEXP = RCTX.Esp
Read9Stack Text1, ActiveStackPosition, ActiveStackPosition, "ESP"
Command2.Caption = "Goto ESP"
Label14 = "STACK PTR At:" & Hex(RCTX.Esp)
End If
End Sub













Private Sub VS1_Change()
If IsS = 1 Then IsS = 0: Exit Sub

If vs1.VALUE = 16383 Then
If WTypeS = 0 Then
ActiveBasePosition = SubBy8(ActiveBasePosition, 4)
Else
ActiveStackPosition = SubBy8(ActiveStackPosition, 4)
End If


ElseIf vs1.VALUE = 16385 Then
If WTypeS = 0 Then
ActiveBasePosition = AddBy8(ActiveBasePosition, 4)
Else
ActiveStackPosition = AddBy8(ActiveStackPosition, 4)
End If

ElseIf vs1.VALUE < 16383 Then
If WTypeS = 0 Then
ActiveBasePosition = SubBy8(ActiveBasePosition, 12)
Else
ActiveStackPosition = SubBy8(ActiveStackPosition, 12)
End If

ElseIf vs1.VALUE > 16385 Then
If WTypeS = 0 Then
ActiveBasePosition = AddBy8(ActiveBasePosition, 12)
Else
ActiveStackPosition = AddBy8(ActiveStackPosition, 12)
End If



End If

ReRead


IsS = 1
vs1.VALUE = 16384
End Sub

Private Sub ReRead()
Dim ThrS As Long
Dim HT As Long
HT = GetHandleOfThread(ShowingTH, ThrS)
If Not (LastState = ThrS) Then
ReadIt
Else
If ThrS = 0 Or ThrS = 2 Then
RCTX = GetContext(ShowingTH)
If RCTX.Ebp = LastEXP And WTypeS = 0 Then
Read9Stack Text1, ActiveBasePosition, RCTX.Ebp, "EBP"
ElseIf RCTX.Esp = LastEXP And WTypeS = 1 Then
Read9Stack Text1, ActiveStackPosition, RCTX.Esp, "ESP"
Else
ReadIt
End If
End If
End If
End Sub

Private Sub CheckThreadCaption(ByRef HT As Long, ByRef ThrS As Long)
If ThrS = 1 Then
Caption = "Thread:" & ShowingTH & " ,Running!"
If WTypeS = 0 Then
Label14 = "BASE"
Else
Label14 = "STACK"
End If
Text1 = ""
ElseIf HT = 0 Then
Caption = "Thread:" & ShowingTH & " ,Not valid now!"
If WTypeS = 0 Then
Label14 = "BASE"
Else
Label14 = "STACK"
End If
Text1 = ""
Else
Caption = "Thread:" & ShowingTH
If WTypeS = 0 Then
Label14 = "BASE PTR At:" & Hex(RCTX.Ebp)
Else
Label14 = "STACK PTR At:" & Hex(RCTX.Esp)
End If
End If
End Sub




