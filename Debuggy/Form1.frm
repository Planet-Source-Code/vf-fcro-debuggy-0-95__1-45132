VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resources"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save in Res"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View Resource"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   7080
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DACCC2&
      Height          =   6750
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   11775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CA9273&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resource Type / Name / Lang Id / Address / Length"
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
      TabIndex        =   1
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public ShowHinstance As Long




Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
List1_dblClick
End Sub

Private Sub Command3_Click()
On Error Resume Next
cd1.ShowSave
If Len(cd1.Filename) = 0 Then Exit Sub
If Dir(cd1.Filename, vbHidden Or vbSystem) <> "" Then
Dir ""
Kill cd1.Filename
If Err <> 0 Then On Error GoTo 0: MsgBox "Cannot delete an old file!", vbCritical, "Error": Exit Sub
Else
Dir ""
End If

DoEvents
Dim VarFileINF As Byte
If vbYes = MsgBox("Include Version Info? (If exist!)", vbYesNo, "Request") Then VarFileINF = 1

SaveToRes cd1.Filename, VarFileINF
MsgBox "Resource File Saved", vbInformation, "Information"
End Sub

Private Sub Form_Load()
 cd1.Filter = "Win32 Resource File (*.res)|*.res|"
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Call SendMessage(List1.hWnd, &H194, ByVal 6000, ByVal 0&)

Dim Tabs() As Long
ReDim Tabs(5)
Tabs(0) = 0
Tabs(1) = 170
Tabs(2) = 335
Tabs(3) = 362
Tabs(4) = 400
Tabs(5) = 450
Call SendMessage(List1.hWnd, &H192, ByVal UBound(Tabs) + 1, Tabs(0))


End Sub

Public Sub ShowIt()
Dim u As Long
List1.Clear
LockWindowUpdate List1.hWnd
Dim Tvc As String
For u = 0 To UBound(ResRC)
If IsNumeric(ResRC(u).ResType) Then
Tvc = NName(ResRC(u).ResType)
Else
Tvc = ResRC(u).ResType
End If

List1.AddItem Tvc & vbTab & ResRC(u).ResName & _
vbTab & ResRC(u).LangId & vbTab & Hex(ResRC(u).ResAddress) & vbTab & _
Hex(ResRC(u).ResLength)
Next u
LockWindowUpdate 0
End Sub

Private Sub List1_dblClick()
On Error GoTo Dalje
If List1.ListIndex = -1 Or List1.ListCount = 0 Then MsgBox "Select Resource First", vbInformation, "Require": Exit Sub
Dim SDataS() As String
Dim ARDesc As Long
Dim iSInvalid As Long
Dim IsEp As Byte
SDataS = Split(List1.List(List1.ListIndex), vbTab)

Select Case SDataS(0)

Case "Bitmap"
If 0 = FixBitmap(ResResData, CLng("&H" & SDataS(3)), CLng("&H" & SDataS(4))) Then
MsgBox "Cannot Fix Bitmap Header!", vbCritical, "Error!": Erase ResResData: Exit Sub
End If
ReDim STD1(0)
ReDim PicWidth(0)
ReDim PicHeight(0)
Set STD1(0) = GetPicture(ResResData)
If STD1(0) = 0 Then MsgBox "Error in Bitmap!", vbCritical, "Error": Erase ResResData: Exit Sub
PicWidth(0) = CLng(STD1(0).Width * (567 / 1000))
PicHeight(0) = CLng(STD1(0).Height * (567 / 1000))
Form19.IccType = 3
Form19.Caption = "Bitmap Resource"
Form19.Show 1

Case "Single Cursor"
If 0 = FixCursor(ResResData, CLng("&H" & SDataS(3)), CLng("&H" & SDataS(4))) Then
MsgBox "Cannot Fix Cursor Header!", vbCritical, "Error!": Erase ResResData: Exit Sub
End If
ReDim STD1(0)
ReDim PicWidth(0)
ReDim PicHeight(0)
Set STD1(0) = GetPicture(ResResData)
If STD1(0) = 0 Then MsgBox "Error in Cursor!", vbCritical, "Error": Erase ResResData: Exit Sub

PicWidth(0) = CLng(STD1(0).Width * (567 / 1000))
PicHeight(0) = CLng(STD1(0).Height * (567 / 1000))
Form19.IccType = 2
Form19.Caption = "Cursor Resource"
Form19.Show 1

Case "Single Icon"
If 0 = FixIcon(ResResData, CLng("&H" & SDataS(3)), CLng("&H" & SDataS(4))) Then
MsgBox "Cannot Fix Icon Header!", vbCritical, "Error!": Erase ResResData: Exit Sub
End If
ReDim STD1(0)
ReDim PicWidth(0)
ReDim PicHeight(0)
Set STD1(0) = GetIconToPicture(ResResData)
If STD1(0) = 0 Then MsgBox "Error in Icon!", vbCritical, "Error": Erase ResResData: Exit Sub

PicWidth(0) = CLng(STD1(0).Width * (567 / 1000))
PicHeight(0) = CLng(STD1(0).Height * (567 / 1000))
Form19.IccType = 1
Form19.Caption = "Icon Resource"
Form19.Show 1

Case "Group Cursor"
Set BITCRI = Nothing
If 0 = CursorGroup(CLng("&H" & SDataS(3)), CLng("&H" & SDataS(4))) Then
MsgBox "Cannot Fix Cursor Group!", vbCritical, "Error!": Erase ResResData: Exit Sub
End If
Form19.IccType = 12
Form19.Caption = "Cursor Group Resource"
Form19.Show 1

Case "Group Icon"
Set BITCRI = Nothing
If 0 = IconGroup(CLng("&H" & SDataS(3)), CLng("&H" & SDataS(4))) Then
MsgBox "Cannot Fix Icon Group!", vbCritical, "Error!": Erase ResResData: Exit Sub
End If
Form19.IccType = 11
Form19.Caption = "Icon Group Resource"
Form19.Show 1

Case "String"
LoadStr CLng("&H" & SDataS(3)), CLng("&H" & SDataS(4)), CLng(SDataS(1)), IsEp, LDSTRINGS
ArrayDescriptor ARDesc, LDSTRINGS, 4
If IsEp = 0 Or ARDesc = 0 Then MsgBox "Cannot Read Strings!", vbCritical, "Error!": Exit Sub
Form21.WTypeIs = 0
Form21.Show 1

Case "Message Table"
LoadMSGTable CLng("&H" & SDataS(3)), CLng("&H" & SDataS(4)), IsEp, LDSTRINGS
ArrayDescriptor ARDesc, LDSTRINGS, 4
If IsEp = 0 Or ARDesc = 0 Then MsgBox "Cannot Read Strings!", vbCritical, "Error!": Exit Sub
Form21.WTypeIs = 1
Form21.Show 1


Case "Dialog Box"

Dim TWSP() As Byte
GetDataFromMem CLng("&H" & SDataS(3)), TWSP, CLng("&H" & SDataS(4)), iSInvalid
If iSInvalid = 0 Then MsgBox "Error in Memory", vbCritical, "Error": Exit Sub
Dim HHDL As Long
HHDL = CreateDialogIndirectParam(APP.hInstance, ByVal VarPtr(TWSP(0)), Form1.hWnd, AddressOf dProc, 0&)
If HHDL = 0 Then MsgBox "Cannot Display Dialog Box (maybe contain unregistered Window Class or Menu)", vbExclamation, "Information": Exit Sub
ShowWindow HHDL, 0
Form22.WHDL = HHDL
Form22.Show 1
Erase TWSP

Case "Menu"

Dim TWSP2() As Byte
GetDataFromMem CLng("&H" & SDataS(3)), TWSP2, CLng("&H" & SDataS(4)), iSInvalid
If iSInvalid = 0 Then MsgBox "Error in Memory", vbCritical, "Error": Exit Sub
Dim MnHD As Long
MnHD = LoadMenuIndirect(ByVal VarPtr(TWSP2(0)))
If MnHD = 0 Then MsgBox "Error in Menu", vbCritical, "Error": Exit Sub
Dim USTRX As String
Dim chrllen As Long
Dim MnuCNT As Long
Dim IID As Long
Dim u As Long
MnuCNT = GetMenuItemCount(MnHD)
For u = 0 To MnuCNT - 1
USTRX = Space(255)
chrllen = GetMenuString(MnHD, u, USTRX, 255, &H400&)
USTRX = Left(USTRX, chrllen)
If USTRX = "" Then
IID = GetMenuItemID(MnHD, u)
Dim modd() As Byte
modd = StrConv("(Hidden Menu)" & Chr(CByte(0)), vbFromUnicode)
Call ModifyMenu(MnHD, u, &H400&, IID, VarPtr(modd(0)))
End If
Next u
Form23.MNUs = MnHD
Form23.Show 1
Erase modd
Erase TWSP2


Case Else
Dim TestHead As String
Dim IsVDD As Long
Dim VAdr As Long
Dim VLen As Long

VAdr = CLng("&H" & SDataS(3))
VLen = CLng("&H" & SDataS(4))
If VLen > 15 Then
TestHead = Space(15)
IsVDD = ReadProcessMemory(ProcessHandle, ByVal VAdr, ByVal TestHead, 15, ByVal 0&)

If StrComp(Mid(TestHead, 9, 3), "AVI") = 0 Then
'Avi File
GetDataFromMem VAdr, ResResData, VLen
Dim FFLX As Long
FFLX = FreeFile
Open AddSlash(APP.Path) & "\testAvi.Avi" For Binary As #FFLX
Put #FFLX, , ResResData
Close #FFLX
Erase ResResData
Form20.Show 1

Else 'If StrComp(Mid(TestHead, 4, 4), Chr(&HFF) & Chr(&HD8) & Chr(&HFF) & Chr(&HE0)) = 0 Then
ReDim STD1(0)
ReDim PicWidth(0)
ReDim PicHeight(0)
GetDataFromMem VAdr, ResResData, VLen
Set STD1(0) = GetPicture(ResResData)
If STD1(0) = 0 Then Erase ResResData: Exit Sub
PicWidth(0) = CLng(STD1(0).Width * (567 / 1000))
PicHeight(0) = CLng(STD1(0).Height * (567 / 1000))
Form19.IccType = 5
Form19.Caption = "JPG Resource"
Form19.Show 1


End If




End If




End Select
Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Cannot Present this type of the Resource!", vbExclamation, "Information"
End Sub
