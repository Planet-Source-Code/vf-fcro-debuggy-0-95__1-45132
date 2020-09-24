VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form27 
   Caption         =   "File Hex Editor"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form27.frx":0000
   LinkTopic       =   "Form27"
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Calc Jump"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "RVA To RAW"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   8160
      MaxLength       =   8
      TabIndex        =   6
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save File"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   5160
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3120
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.VScrollBar VS1 
      Height          =   4695
      Left            =   9240
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open File"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox TextX 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E7DFD6&
      Height          =   4695
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   9255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E7DFD6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Goto Address:"
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
      Left            =   8160
      TabIndex        =   7
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E7DFD6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "File:"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      TabIndex        =   5
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "Form27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents VSCROLL As CLongScroll
Attribute VSCROLL.VB_VarHelpID = -1
Private WithEvents PX As ProcHex
Attribute PX.VB_VarHelpID = -1
Private Sub Command1_Click()
On Error GoTo Dalje
cd1.ShowOpen
If Len(cd1.Filename) = 0 Then Exit Sub
Dim FLGF As Long
FLGF = FreeFile

Open cd1.Filename For Binary As #FLGF
ReDim FileDataPW(LOF(FLGF) - 1)
Get #FLGF, , FileDataPW
Label2 = "File:" & cd1.Filename & " ,Length:" & Hex(LOF(FLGF))
Close #FLGF
ReShow
Exit Sub


Dalje:
On Error GoTo 0
MsgBox "File Reading Error!", vbCritical, "Error": Close #FLGF
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error GoTo Dalje
Dim ARD As Long
ArrayDescriptor ARD, FileDataPW, 4
If ARD = 0 Then Exit Sub

cd1.ShowSave
If Len(cd1.Filename) = 0 Then Exit Sub
If Dir(cd1.Filename, vbHidden Or vbReadOnly) <> "" Then
Dir ""
Kill cd1.Filename
Else
Dir ""
End If

Dim FLGF As Long
FLGF = FreeFile
DoEvents
Open cd1.Filename For Binary As #FLGF
Put #FLGF, , FileDataPW
Close #FLGF

MsgBox "File Saved...", vbInformation, "Information"

Exit Sub
Dalje:
On Error GoTo 0
Close #FLGF
MsgBox "Error during save file", vbCritical, "Error"
End Sub

Private Sub Command4_Click()
Form26.Show 1
End Sub

Private Sub Command5_Click()
Form31.Show 1
End Sub

Private Sub Form_Load()
TextX.FontName = "FixedSys"
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2

TWProc1 = SetWindowLong(TextX.hWnd, -4, AddressOf TextProc)

Set VSCROLL = New CLongScroll
Set VSCROLL.Client = VS1
Set PX = New ProcHex
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetWindowLong TextX.hWnd, -4, TWProc1
Erase FileDataPW
End Sub

Private Sub PX_UpdateAdr(ByVal Address As Long, ByVal Data As Byte, CancelB As Boolean)
On Error GoTo Dalje
FileDataPW(Address) = Data
Exit Sub
Dalje:
On Error GoTo 0
CancelB = True
End Sub





Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo Dalje
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 8 Then
Exit Sub
ElseIf KeyAscii = 13 Then
If Len(Text1) = 0 Then Text1 = "": Exit Sub
Dim Gadrr As Long
Dim ARD As Long
ArrayDescriptor ARD, FileDataPW, 4
If ARD = 0 Then Exit Sub
If UBound(FileDataPW) + 1 <= 320 Then Exit Sub
Gadrr = CLng("&H" & Text1)
If Gadrr < 0 Or Gadrr > UBound(FileDataPW) Then MsgBox "Out Of Range!", vbExclamation, "Information": Text1 = "": Exit Sub
VSCROLL.Value = Int((Gadrr) / 16&) + 1
vs1_Change
End If

If IsValidK(Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Unknown Value Type", vbCritical, "Error"
End Sub

Private Sub vs1_Change()
Dim ARD As Long
ArrayDescriptor ARD, FileDataPW, 4
If ARD <> 0 Then
PrintDump2 TextX, VSCROLL.Value - 1
End If
If PX Is Nothing Then Exit Sub
PX.MAXLEN = Len(Text1)
End Sub

Private Sub ReShow()
If UBound(FileDataPW) + 1 <= 320 Then
VS1.Enabled = False
PrintDump2 TextX, 0
Else
With VSCROLL
.Min = 1
.Max = CLng(Int((UBound(FileDataPW) + 1) / 16))
If ((UBound(FileDataPW) + 1) Mod 16) = 0 Then .Max = .Max - 1
.Max = .Max - 18
.SmallChange = 1
.LargeChange = 16
.Value = 0

VS1.Enabled = True
vs1_Change
End With
End If

PX.MAXLEN = Len(Text1)
Set PX.Text1 = TextX

End Sub

Public Sub PrintDump2(ByVal TXT As TextBox, ByVal Position As Long)
Dim u As Long
Dim DaX(19) As String
For u = 0 To 19
If (Position + u) * 16& >= UBound(FileDataPW) + 1 Then Exit For
DaX(u) = GetHexDump((Position + u) * 16&, 1) & vbCrLf
Next u
TXT = Join(DaX, "")
End Sub
