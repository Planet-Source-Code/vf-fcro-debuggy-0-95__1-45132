VERSION 5.00
Begin VB.Form Form29 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Config Tracer"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2130
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form29"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   2130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Find Last"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enable Range"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Accept"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2160
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
      Index           =   5
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Stop After="
      Top             =   1800
      Width           =   1095
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
      Index           =   4
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1800
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
      Index           =   1
      Left            =   960
      MaxLength       =   8
      TabIndex        =   3
      Top             =   600
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
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "TO="
      Top             =   600
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
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "FROM="
      Top             =   360
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
      Left            =   960
      MaxLength       =   8
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E7DFD6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Notify In Range"
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempFat(20) As Long
Dim IsOk As Byte
Private Sub Command1_Click()
On Error GoTo Dalje2

If Len(Text3(4)) = 0 Then MsgBox "Stop Notification isn't set!", vbExclamation, "Error": Exit Sub

If Len(Text3(0)) = 0 And Len(Text3(1)) = 0 Then
 TempFat(0) = 0
 TempFat(1) = 0
 TempFat(2) = 0
MsgBox "Range is disabled!", vbInformation, "Information"
GoTo Dalje
ElseIf (Len(Text3(0)) <> 0 And Len(Text3(1)) = 0) Or _
(Len(Text3(0)) = 0 And Len(Text3(1)) <> 0) Then
MsgBox "Error In Range!", vbCritical, "Error": Exit Sub
End If

 TempFat(0) = "&H" & Text3(0)
 TempFat(1) = "&H" & Text3(1)
 TempFat(2) = Check1.Value

If Check1.Value <> 0 Then
If TempFat(1) <= TempFat(0) Then MsgBox "Error in Range!", vbCritical, "Error": Exit Sub
Dim TName As String
Dim TName2 As String
TName = FindInModules(TempFat(0))
TName2 = FindInModules(TempFat(1))

If Len(TName) = 0 And Len(TName2) = 0 Then
If vbNo = MsgBox("Range doesn't belong to any module! Proceed?", vbYesNo, "Confirm") Then Exit Sub
End If

End If

Dalje:
Dim Stp1 As Long
Stp1 = "&H" & Text3(4)
If Stp1 = 0 Then MsgBox "Cannot notify after 0 instructions!", vbExclamation, "Error!": Exit Sub
TempFat(3) = Stp1
IsOk = 1
Unload Me
Exit Sub
Dalje2:
On Error GoTo 0
MsgBox "Error in Configuration!", vbCritical, "Error"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
FindLastEA Text3(0), Text3(1)
End Sub

Private Sub Form_Load()
IsOk = 0

Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Text3(0) = Hex(TraceConfig(0))
Text3(1) = Hex(TraceConfig(1))
Check1.Value = TraceConfig(2)
Text3(4) = Hex(TraceConfig(3))
End Sub

Private Sub Form_Unload(Cancel As Integer)
If IsOk = 1 Then
CopyMemory TraceConfig(0), TempFat(0), 84
End If
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 8 Then Exit Sub
If IsValidK(Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub

End Sub
