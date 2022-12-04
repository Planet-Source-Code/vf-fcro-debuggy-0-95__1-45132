VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form9 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Define Search Area/Save"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2445
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   2445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2280
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
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
      Left            =   240
      TabIndex        =   7
      Top             =   1200
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
      Left            =   1080
      MaxLength       =   8
      TabIndex        =   6
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
      Left            =   240
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   5
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
      Left            =   240
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   4
      Text            =   "FROM="
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Find Last"
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
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "Find Last Valid Address From"
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
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
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Accept"
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
      Left            =   240
      TabIndex        =   1
      Top             =   720
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
      Left            =   1080
      MaxLength       =   8
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo Dalje
Dim cx As Long
Dim Cxl As Long
Dim IsValidRange As Long
cx = CLng("&H" & Text3(0))
Cxl = CLng("&H" & Text3(1))

If cx > Cxl Then MsgBox "Error in Length!", vbCritical, "Error": NullV: Exit Sub
gBegAdr = cx
gLenAdr = SubBy8(Cxl, cx)

If gLenAdr > 39452672 Or gLenAdr <= 0 Then NullV: MsgBox "Max Search Area: 32MB", vbExclamation, "Information": Exit Sub
GetDataFromMem gBegAdr, DATAPW, gLenAdr, IsValidRange

If IsValidRange = 0 Then MsgBox "Invalid Memory Data! (Entire or some parts of that Range)", vbExclamation, "Information": NullV: Exit Sub
gSTARTADR = gBegAdr: gLASTADR = AddBy8(gBegAdr, gLenAdr)
gSTARTADR2 = gBegAdr: gLASTADR2 = AddBy8(gBegAdr, gLenAdr)
Unload Me
Exit Sub
Dalje:
On Error GoTo 0
NullV
MsgBox "Unknown Value Type!", vbCritical, "Error!"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Private Sub NullV()
gBegAdr = 0
gLenAdr = 0
End Sub


Private Sub Command3_Click()
On Error GoTo Dalje2
Dim cx As Long
Dim Cxl As Long
Dim MvLen As Long
Dim TempsDt() As Byte
Dim IsValidRange As Long
If Len(Text3(0)) = 0 Or Len(Text3(1)) = 0 Then MsgBox "An Empty Address", vbCritical, "Error": Exit Sub
cx = CLng("&H" & Text3(0))
Cxl = CLng("&H" & Text3(1))

If cx > Cxl Then MsgBox "Error in Length!", vbCritical, "Error":  Exit Sub
MvLen = SubBy8(Cxl, cx)
If MvLen <= 0 Then MsgBox "Invalid Range!", vbCritical, "Error": Exit Sub

If MvLen > 39452672 Then MsgBox "Max Search Area: 32MB", vbExclamation, "Information": Exit Sub
GetDataFromMem cx, TempsDt, MvLen, IsValidRange
If IsValidRange = 0 Then MsgBox "Invalid Memory Data! (Entire or some parts of that Range)", vbExclamation, "Information": Exit Sub


cd1.ShowSave
If Len(cd1.Filename) = 0 Then Exit Sub

If Dir(cd1.Filename, vbHidden Or vbReadOnly) <> "" Then
Dir "": Kill cd1.Filename
Else
Dir ""
End If
Dim FFSF As Long

FFSF = FreeFile
On Error GoTo Dalje
Open cd1.Filename For Binary As #FFSF

Put #FFSF, , TempsDt

Close #FFSF

MsgBox "File saved", vbInformation, "Information"
Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Error during Save!", vbCritical, "Error": Close #FFSF
Exit Sub
Dalje2:
On Error GoTo 0
MsgBox "Error in Value!", vbCritical, "Error"
End Sub

Private Sub Command4_Click()
If Len(Text3(0)) = 0 Then MsgBox "From Address isn't set!", vbInformation, "Information": Exit Sub
Dim LPP As Long
Dim IsValid As Byte
Dim TBuff(127) As Byte
LPP = CLng("&H" & Text3(0))
Text3(1) = ""
Do
IsValid = ReadProcessMemory(ProcessHandle, ByVal LPP, TBuff(0), 128, ByVal 0&)
LPP = LPP + 4
Loop While IsValid <> 0
LPP = LPP - 4
Do
IsValid = ReadProcessMemory(ProcessHandle, ByVal LPP, TBuff(0), 1, ByVal 0&)
LPP = LPP + 1
Loop While IsValid <> 0
LPP = LPP - 1
Text3(1) = Hex(LPP)
Erase TBuff
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
