VERSION 5.00
Begin VB.Form Form30 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Spy Notify Configuration"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5100
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
   ScaleHeight     =   3075
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Enable Range"
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   2280
      Width           =   1455
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
      TabIndex        =   12
      Top             =   2520
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
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   11
      Text            =   "FROM="
      Top             =   2520
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
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   10
      Text            =   "TO="
      Top             =   2760
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
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Find Last"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2640
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Registry Apies"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   3255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Volume/Logic Drives Information"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   3255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Move/Rename File/Directory"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Create/ Remove Directory"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   3255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Read/Write/Copy/Map/Unmap/Delete File"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Accept"
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   1
      Top             =   2640
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Create/Open/CreateMap/OpenMap File"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
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
      TabIndex        =   13
      Top             =   2280
      Width           =   1935
   End
End
Attribute VB_Name = "Form30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
On Error GoTo Dalje2
If Index <> 0 Then Unload Me: Exit Sub


If Len(Text3(0)) = 0 And Len(Text3(1)) = 0 Then
NOTA1 = 0
NOTA2 = 0
IsENOT = 0
MsgBox "Range is disabled!", vbInformation, "Information"
GoTo Dalje

ElseIf (Len(Text3(0)) <> 0 And Len(Text3(1)) = 0) Or _
(Len(Text3(0)) = 0 And Len(Text3(1)) <> 0) Then
MsgBox "Error In Range!", vbCritical, "Error": Exit Sub
End If

 NOTA1 = "&H" & Text3(0)
 NOTA2 = "&H" & Text3(1)
 IsENOT = Check2.Value

If Check2.Value <> 0 Then
If NOTA1 >= NOTA2 Then MsgBox "Error in Range!", vbCritical, "Error": Exit Sub
Dim TName As String
Dim TName2 As String
TName = FindInModules(NOTA1)
TName2 = FindInModules(NOTA2)

If Len(TName) = 0 And Len(TName2) = 0 Then
If vbNo = MsgBox("Range doesn't belong to any module! Proceed?", vbYesNo, "Confirm") Then Exit Sub
End If

End If


Dalje:

Dim u As Long
For u = 0 To Check1.UBound
SPConfig(u) = Check1(u).Value
Next u


Unload Me
Exit Sub
Dalje2:
On Error GoTo 0
MsgBox "Error in Value!", vbCritical, "Error"
End Sub



Private Sub Command3_Click()
FindLastEA Text3(0), Text3(1)
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Dim u As Long
For u = 0 To Check1.UBound
Check1(u).Value = SPConfig(u)
Next u

Text3(0) = Hex(NOTA1)
Text3(1) = Hex(NOTA2)
Check2.Value = IsENOT

End Sub
