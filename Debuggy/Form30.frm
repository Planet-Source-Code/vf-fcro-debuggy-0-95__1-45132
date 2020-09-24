VERSION 5.00
Begin VB.Form Form30 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "API Spy"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form30.frx":0000
   LinkTopic       =   "Form30"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Config"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   4800
      Width           =   975
   End
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
      ForeColor       =   &H00E7DFD6&
      Height          =   4515
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   7935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hide"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Log"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E7DFD6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Spy"
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
      Width           =   7935
   End
End
Attribute VB_Name = "Form30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TypeOfFF As Byte

Public Sub AddInAPI(ByRef StringX As String)
If List1.ListCount = MAXLINES Then

Dim u As Long
For u = 0 To 7
List1.RemoveItem 0
Next u

End If
List1.AddItem StringX

If SlowNOTIFY = 1 Then
List1.ListIndex = List1.ListCount - 1
If Visible = False Then Visible = True
OnScreen hWnd
DoEvents
SleepMe 150
End If


End Sub




Private Sub Command3_Click()
Select Case TypeOfFF
Case 1
KSpy.Show 1
Case 2
USpy.Show 1
Case 3
GSpy.Show 1

Case 4

End Select

End Sub

Private Sub Form_Load()
RemoveX hWnd
Call SendMessage(List1.hWnd, &H194, ByVal 2500, ByVal 0&)

End Sub
Private Sub Command1_Click()
Visible = False
End Sub

Private Sub Command2_Click()
List1.Clear
End Sub


