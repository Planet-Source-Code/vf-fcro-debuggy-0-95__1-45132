VERSION 5.00
Begin VB.Form Form14 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Examine Window"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7650
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Query Window Proc"
      Height          =   255
      Left            =   5640
      TabIndex        =   15
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "STYLE="
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "EX STYLE="
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "TEXT="
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Refresh Info"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Visible/Hide"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Enable/Disable"
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change ExStyle"
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox Text4 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   4560
      MaxLength       =   8
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text3 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   840
      MaxLength       =   8
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change Style"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text2 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   720
      TabIndex        =   2
      Top             =   4200
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Text"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
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
      ForeColor       =   &H00E7DFD6&
      Height          =   3735
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   7575
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E7DFD6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Window Proc At Address:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
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
      TabIndex        =   11
      Top             =   3840
      Width           =   5535
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ACChwnd As Long
Private PRThwnd As Long
Private STLy As Long
Private EXSTLy As Long
Private IsEnb As Long
Private IsVsb As Long
Private HinstV As Long
Private Hprc As Long
Private TextPRvv As String
Private PRclassname As String
Private Wclassname As String
Private InstName As String
Private InOurPRC As Long

Private Sub Command1_Click()
Dim iret As Long
iret = SetWindowText(ACChwnd, ByVal Text2)
If iret = 0 Then
MsgBox "Unable To Change Text!", vbCritical, "Error": Exit Sub
Else
Examine
End If
End Sub

Private Sub Command2_Click()
On Error GoTo Dalje
Dim iret As Long
iret = SetWindowLong(ACChwnd, -16, CLng("&H" & Text3))
If iret = 0 Then
MsgBox "Unable To Change Style!", vbCritical, "Error"
Else
InvalidateRect PRThwnd, ByVal 0&, 1
Examine
End If
Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Unknown Value Type!", vbCritical, "Error"

End Sub

Private Sub Command3_Click()
On Error GoTo Dalje
Dim iret As Long
iret = SetWindowLong(ACChwnd, -20, CLng("&H" & Text4))
If iret = 0 Then
MsgBox "Unable To Change ExStyle!", vbCritical, "Error"
Else
InvalidateRect 0, ByVal 0&, 1
Examine
End If
Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Unknown Value Type!", vbCritical, "Error"

End Sub

Private Sub Command4_Click()
Dim iret As Long
If IsEnb = 0 Then
iret = EnableWindow(ACChwnd, 1)
If iret <> 0 Then IsEnb = 1
Else
iret = EnableWindow(ACChwnd, 0)
If iret <> 0 Then IsEnb = 0
End If
Examine
End Sub

Private Sub Command5_Click()
Dim iret As Long
If IsVsb = 0 Then
iret = ShowWindow(ACChwnd, 1)
If iret <> 0 Then IsVsb = 1
Else
iret = ShowWindow(ACChwnd, 0)
If iret <> 0 Then IsVsb = 0
End If
Examine
End Sub


Private Sub Command6_Click()
If DebuggyOut = 0 Then MsgBox "Cannot Query Window Proc!", vbExclamation, "Information": Exit Sub
Call CreateRemoteThread(ProcessHandle, ByVal 0&, 10, ByVal DebuggyOut, ByVal ACChwnd, 0, AccThreadX)

End Sub

Private Sub Command7_Click()

Unload Me
End Sub

Private Sub Command8_Click()
Examine
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Examine
End Sub


Private Sub Examine()

Dim EnString As String
Dim ViString As String
PRThwnd = GetParent(ACChwnd)
PRclassname = ClassNameEx(PRThwnd)
Wclassname = ClassNameEx(ACChwnd)
IsEnb = IsWindowEnabled(ACChwnd)
IsVsb = IsWindowVisible(ACChwnd)
STLy = GetWindowLong(ACChwnd, -16)
HinstV = GetWindowLong(ACChwnd, -6)
EXSTLy = GetWindowLong(ACChwnd, -20)
Hprc = GetClassLong(ACChwnd, -24)

If IsEnb = 0 Then
EnString = "No"
Else
EnString = "Yes"
End If
If IsVsb = 0 Then
ViString = "No"
Else
ViString = "Yes"
End If

Dim wlen As Long
wlen = GetWindowTextLength(ACChwnd)
If wlen = 0 Then
TextPRvv = " (No Text)"
Else
TextPRvv = Space(wlen + 1)
If wlen > 256 Then wlen = 256
GetWindowText ACChwnd, TextPRvv, wlen + 1
End If

InstName = FindInModules(HinstV)

Text1 = "Window Class Name:" & Wclassname & vbCrLf & _
"Hwnd:" & Hex(ACChwnd) & vbCrLf & _
"Parent Class Name:" & PRclassname & vbCrLf & _
"Parent Hwnd:" & Hex(PRThwnd) & vbCrLf & _
"Enabled:" & EnString & vbCrLf & _
"Visible:" & ViString & vbCrLf & _
"Style:" & Hex(STLy) & vbCrLf & _
"ExStyle:" & Hex(EXSTLy) & vbCrLf & _
"Hinstance:" & Hex(HinstV) & " ,In Module:" & InstName & vbCrLf & _
"Class Proc At Address:" & Hex(Hprc) & vbCrLf & _
"Text:" & TextPRvv

End Sub





Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 8 Then Exit Sub
If IsValidK(Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub

End Sub

Private Sub Text4_Change()
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 8 Then Exit Sub
If IsValidK(Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub

End Sub
