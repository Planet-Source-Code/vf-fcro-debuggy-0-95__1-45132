VERSION 5.00
Begin VB.Form Form21 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Table/Strings"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7275
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form21"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   4440
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
      ForeColor       =   &H00DACCC2&
      Height          =   4125
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00DCB17C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Entry / Data"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WTypeIs As Byte

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Call SendMessage(List1.hwnd, &H194, ByVal 2000, ByVal 0&)
Dim Tabs(1) As Long
Tabs(0) = 0
Tabs(1) = 40
Call SendMessage(List1.hwnd, &H192, ByVal UBound(Tabs) + 1, Tabs(0))

If WTypeIs = 0 Then
Caption = "String Resource"
ElseIf WTypeIs = 1 Then
Caption = "Message Table Resource"
End If
ReadStrings
End Sub


Private Sub ReadStrings()
On Error GoTo Dalje
Dim u As Long
List1.Clear
LockWindowUpdate List1.hwnd
For u = 0 To UBound(LDSTRINGS)
List1.AddItem Hex(LDSTRINGS(u).id) & vbTab & LDSTRINGS(u).data
Next u
LockWindowUpdate 0
Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Maximum number of Items reached!", vbExclamation, "Information!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Erase LDSTRINGS
End Sub
