VERSION 5.00
Begin VB.Form Form31 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Jump Calculator"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3915
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form31"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Long Jump"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Short Jump"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text3 
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
      Height          =   285
      Index           =   1
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   3
      Left            =   960
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "TO="
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   2
      Left            =   960
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "FROM="
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text3 
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
      Height          =   285
      Index           =   0
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
On Error GoTo Dalje
Dim AA1 As Long
Dim AA2 As Long
If CheckAdrAdr(AA1, AA2) = 0 Then Exit Sub

Dim ret As Long
If AA1 < AA2 Then
If SubBy8(AA2, AA1) > 129 Then GoTo Jimp

ElseIf AA2 < AA1 Then
If SubBy8(AA1, AA2) > 126 Then
Jimp:
MsgBox "Out of Short Jump Range!", vbExclamation, "Information": Exit Sub
End If

End If

ret = CalcShortJumpBack(AA1, AA2)
Text3(4) = "OFFSET: XX " & BBSTR(CByte(ret))

Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Calculating Error", vbCritical, "Error"
End Sub

Private Sub Command2_Click()
Dim AA1 As Long
Dim AA2 As Long
If CheckAdrAdr(AA1, AA2) = 0 Then Exit Sub

Dim ret As Long
Dim GvB As String

ret = CalcLongJumpBack(AA1, AA2)
If vbYes = MsgBox("Use 2 byte opcode?", vbYesNo, "Choose Opcode Length") Then
ret = ret - 1
GvB = " XX XX "
Else
GvB = " XX "
End If

Dim VsBytes(3) As Byte
CopyMemory VsBytes(0), ret, 4

Text3(4) = "OFFSET:" & GvB & BBSTR(VsBytes(0)) & " " & BBSTR(VsBytes(1)) _
& " " & BBSTR(VsBytes(2)) & " " & BBSTR(VsBytes(3))

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 8 Then Exit Sub
If IsValidK(Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub

End Sub


Private Function CheckAdrAdr(ByRef AA1 As Long, ByRef AA2 As Long) As Byte
On Error GoTo Dalje
If Len(Text3(0)) = 0 Or Len(Text3(1)) = 0 Then GoTo Dalje


AA1 = CLng("&H" & Text3(0))
AA2 = CLng("&H" & Text3(1))

CheckAdrAdr = 1
Exit Function

Dalje:
On Error GoTo 0
MsgBox "Error in Values!", vbInformation, "Information": Exit Function
End Function

