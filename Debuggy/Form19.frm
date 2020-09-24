VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form19 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pics"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6045
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
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
      Height          =   3345
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Data"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   1440
      ScaleHeight     =   3585
      ScaleWidth      =   4545
      TabIndex        =   1
      Top             =   600
      Width           =   4575
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Index           =   0
         Left            =   0
         ScaleHeight     =   1935
         ScaleWidth      =   1695
         TabIndex        =   2
         Top             =   0
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E7DFD6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sizes"
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
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'cd1.Filter = "Bitmap (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif|"
Public IccType As Byte

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error GoTo Dalje
cd1.ShowSave
If Len(cd1.Filename) = 0 Then Exit Sub
Dim FFll As Long
FFll = FreeFile
If Dir(cd1.Filename, vbSystem Or vbReadOnly) <> "" Then
Dir "": Kill cd1.Filename
Else
Dir ""
End If
Open cd1.Filename For Binary As #FFll

If IccType = 12 Then
SaveCursor FFll
ElseIf IccType = 11 Then
SaveIcon FFll
Else
Put #FFll, , ResResData
End If

Close #FFll
Exit Sub
Dalje:
On Error GoTo 0
Close #FFll
MsgBox "Error during save", vbCritical, "Error"
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
ShowPics

If IccType = 1 Or IccType = 11 Then
 cd1.Filter = "Icon (*.ico)|*.ico|"
ElseIf IccType = 2 Or IccType = 12 Then
 cd1.Filter = "Cursor (*.cur)|*.cur|"
ElseIf IccType = 3 Then
 cd1.Filter = "Bitmap (*.bmp)|*.bmp|"
 ElseIf IccType = 5 Then
 cd1.Filter = "JPG (*.jpg)|*.jpg|"
End If

List1.Clear
Dim u As Long
For u = 0 To UBound(PicWidth)
List1.AddItem CStr(CLng(PicWidth(u) / 15&)) & " X " & CStr(CLng(PicHeight(u) / 15&))
Next u
End Sub
Public Sub DestroyPics()
Dim u As Long
For u = 0 To UBound(STD1)
Set STD1(u) = Nothing
Next u
Erase STD1
Erase PicWidth
Erase PicHeight
End Sub

Public Sub ShowPics()
Dim allWidth As Long
Dim allHeight As Long
Set Picture1(0).Picture = STD1(0)
Picture1(0).Width = PicWidth(0)
Picture1(0).Height = PicHeight(0)
allWidth = PicWidth(0)
allHeight = PicHeight(0)
Dim u As Long
For u = 1 To UBound(STD1)
Load Picture1(u)
Picture1(u).Left = Picture1(u - 1).Left + Picture1(u - 1).Width + 15 * 5
Picture1(u).Visible = True
If PicHeight(u) > allHeight Then allHeight = PicHeight(u)
allWidth = allWidth + PicWidth(u) + 15 * 5
Set Picture1(u).Picture = STD1(u)
Picture1(u).Width = PicWidth(u)
Picture1(u).Height = PicHeight(u)
Next u
Picture2.Width = allWidth
Picture2.Height = allHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
DestroyPics
Erase ResResData
Set BITCRI = Nothing
End Sub
