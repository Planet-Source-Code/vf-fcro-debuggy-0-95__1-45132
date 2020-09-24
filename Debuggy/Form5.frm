VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Breakpoints"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   8
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Delete BP"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Delete BPs"
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Enable BPs"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disable BPs"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Goto BP"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.ListBox List4 
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
      Height          =   4710
      Left            =   0
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List4_dblClick
End Sub

Private Sub Command2_Click()
'Set RTRIGGER = Nothing
TRIGGERADDRESS = 0
TRIGGERFLAG = -1
RestoreAllOriginalBytes ACTIVEBREAKPOINTS
ISBPDisabled = 1
Form16.Label4.BackColor = &HFF&
End Sub

Private Sub Command3_Click()
'Set RTRIGGER = Nothing
TRIGGERADDRESS = 0
TRIGGERFLAG = -1
RestoreAllBreakPoints ACTIVEBREAKPOINTS
ISBPDisabled = 0
Form16.Label4.BackColor = &HAA00&
End Sub

Private Sub Command4_Click()
NextB = 0
Unload Me
End Sub

Private Sub Command5_Click()
'Set RTRIGGER = Nothing
TRIGGERADDRESS = 0
TRIGGERFLAG = -1
RestoreAllOriginalBytes ACTIVEBREAKPOINTS
Set ACTIVEBREAKPOINTS = Nothing
ChoosedAdr = DISCOUNT
Command4_Click
End Sub

Private Sub Command6_Click()
Dim Itms() As Long

Dim PiRef As Byte
Dim ItmsCount As Long
ItmsCount = GetSelectedItems(List4.hWnd, Itms)
If ItmsCount = 0 Then MsgBox "Nothing selected yet.", vbInformation, "Information": Exit Sub


'Set RTRIGGER = Nothing
TRIGGERADDRESS = 0
TRIGGERFLAG = -1

Dim AdrR As Long
Dim Vuz() As String
Dim u As Long
For u = 0 To ItmsCount - 1
Vuz = Split(List4.List(Itms(u)), vbTab)
AdrR = CLng("&H" & Vuz(0))
RemoveBreakPoint ACTIVEBREAKPOINTS, AdrR

Next u
Erase Itms
Erase Vuz
NextB = 0
Form16.ReleaseShow 1
PrintDump Form8.TextX, ActiveMemPos

If ACTIVEBREAKPOINTS.count = 0 Then
ChoosedAdr = DISCOUNT: Command4_Click
Else
ReadBPSS
End If

End Sub

Private Sub Command7_Click(Index As Integer)
If List4.ListCount = 0 Then Exit Sub
If List4.ListIndex = -1 Then List4.ListIndex = 0: GoTo Rer


Select Case Index

Case 0
If List4.ListIndex <> 0 Then List4.ListIndex = List4.ListIndex - 1


Case 1
If List4.ListIndex <> List4.ListCount - 1 Then List4.ListIndex = List4.ListIndex + 1

End Select

Rer:
ClearSelected List4.hWnd
List4.Selected(List4.ListIndex) = True
Dim Xs() As String
Xs = Split(List4.List(List4.ListIndex), vbTab)
ChoosedAdr = CLng("&H" & Xs(0))
NextB = 0
DISCOUNT = ChoosedAdr
Form16.ReleaseShow 1
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Call SendMessage(List4.hWnd, &H194, ByVal 1000, ByVal 0&)
Dim Tabs() As Long
ReDim Tabs(3)
Tabs(0) = 20
Tabs(1) = 40
Tabs(2) = 220
Tabs(3) = 330
Call SendMessage(List4.hWnd, &H192, ByVal UBound(Tabs) + 1, Tabs(0))

ReadBPSS


End Sub
Private Sub ReadBPSS()
Dim u As Long
Dim BTP() As Long
Dim TDx() As Byte
Dim MnM As String
Dim FWDD As Byte
Dim ISVV As Byte

Dim DissS As String
Dim TA As Long
Dim TA2 As Long
List4.Clear
SpeedUpAdding List4.hWnd, 5000, 100000
For u = 1 To ACTIVEBREAKPOINTS.count
BTP = ACTIVEBREAKPOINTS.Item(u)
GetDataFromMem BTP(0), TDx, 16
TDx(0) = CByte(BTP(1))
MnM = FindInModules(BTP(0), TA, TA2)
DASM.BaseAddress = BTP(0)
DissS = DASM.DisAssemble(TDx, 0, FWDD, 0, 0, ISVV)
List4.AddItem Hex(BTP(0)) & vbTab & DissS & vbTab & "In Module:" & MnM
Next u
End Sub
Private Sub List4_dblClick()
If List4.ListIndex = -1 Then Exit Sub
Dim Xs() As String
Xs = Split(List4.List(List4.ListIndex), vbTab)
ChoosedAdr = CLng("&H" & Xs(0))
Unload Me
End Sub
