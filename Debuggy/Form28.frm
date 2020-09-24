VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form28 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Thread Tracer"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11070
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form28"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Peek Address"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Config"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go!"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   6720
      Width           =   855
   End
   Begin RichTextLib.RichTextBox rt1 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   10821
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"Form28.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C89F8E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Peek Address:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00DCB17C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E7DFD6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tracer"
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
      Width           =   11055
   End
End
Attribute VB_Name = "Form28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private GLOBALCOUNTER As Long
Public ContThread As Long
Private STTs As Byte
Private Stopped As Byte

Private Sub Command1_Click()

SetSingleStep ContThread
ContinueDebug
STTs = 0
GLOBALCOUNTER = 1

End Sub



Private Sub Command2_Click()
CTX = GetContext(ContThread)
Label1 = "Peek Address:" & Hex(CTX.Eip)
End Sub

Private Sub Command3_Click()
If Stopped = 1 Then
Unload Me
Else
Stopped = 2
End If
End Sub

Private Sub Command4_Click()
Form29.Show 1
GLOBALCOUNTER = 1
End Sub

Private Sub Command5_Click()
rt1.Text = ""
End Sub

Private Sub Form_Load()
Stopped = 1
RemoveX hWnd
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
ContThread = ActiveThread
Label2 = "Trace Thread:" & ContThread
GLOBALCOUNTER = 1
End Sub


Public Sub ProcessException(ByRef Address As Long)
'DISCOUNT = Address


If TraceConfig(2) = 0 Then GoTo Dalje


If Address >= TraceConfig(0) And Address <= TraceConfig(1) Then
Dalje:
If STTs = 1 Then
Label9 = "Status:Running in Notify Range!"
AddLine Add1(Address), rt1
GLOBALCOUNTER = GLOBALCOUNTER + 1
Else
AddLine Add1(Address), rt1
STTs = 1
End If

Else
If STTs <> 2 Then
Label9 = "Status:Running out of Notify Range!"
STTs = 2
End If

End If

If Stopped = 2 Then Unload Me: Exit Sub
If GLOBALCOUNTER >= TraceConfig(3) Then GLOBALCOUNTER = 1: STTs = 0: Label9 = "Status:Stoped": Exit Sub
SetSingleStep ContThread
ContinueDebug

End Sub

Private Sub Form_Unload(Cancel As Integer)
GLOBALCOUNTER = 1
rt1.Text = ""
UseTrace = 0
ConfigData(6) = isWinC
CTX = GetContext(ContThread)
DISCOUNT = CTX.Eip
NextB = 0
Form16.ReleaseShow 1

CopyMemory ByVal VarPtr(ConfigData(0)), ByVal VarPtr(TemConfig(0)), 80

'If ConfigData(7) = 1 Then
RestoreAllBreakPoints SPYBREAKPOINTS
'End If

If isCC = 0 And ACTIVEBREAKPOINTS.count <> 0 Then
RestoreAllBreakPoints ACTIVEBREAKPOINTS
ISBPDisabled = 0
Form16.Label4.BackColor = &HAA00&
End If

ActiveThread = ContThread
Form16.Label2 = "Selected: " & ContThread
ClearSingleStep ContThread

'ChangeStateThread ContThread, 2
'ReadThreadsFromProcess Form16.List1
'ContinueDebug

End Sub


Private Function Add1(ByRef Address As Long) As String
Dim DTX() As Byte
Dim CMDS As String
Dim AREF As String
Dim Allpr As String
Dim CHKC As String
Dim IsError As Byte
Dim CRef As String
Dim ExpSt As String
Dim IsValidBP As Byte
Dim BinBin As String
Dim ORGBYTE As Byte
Dim GFI As String
Dim IsString As Long
Dim i As Long





GetDataFromMem Address, DTX, 16
DASM.BaseAddress = Address





CMDS = DASM.DisAssemble(DTX, 0, Forward, 0, 0, IsError)

i = VALUES1

If NOTIFYVALG = 1 Then
AREF = IsStringOnAdr(IsString)
If IsString = 1 Then AREF = "(Possible) String: " & AREF
End If

ExpSt = GetFromExportsSearch(FindInModules(Address), Address)
If Len(ExpSt) <> 0 Then ExpSt = "Export:" & ExpSt


'GFI = GetFromIndex(INDEXESR, REFSR, Address)
'If Len(GFI) = 0 Then
'GFI = GetFromIndex(EINDEXESR, EREFSR, Address)
'End If


CHKC = CheckCALL(i, 1)
Allpr = ExpSt



If Len(Allpr) <> 0 And Len(CHKC) <> 0 Then
Allpr = Allpr & " ;" & CHKC
ElseIf Len(CHKC) <> 0 Then
Allpr = CHKC
End If

If Len(Allpr) <> 0 And Len(AREF) <> 0 Then
If Len(CHKC) = 0 Then
Allpr = Allpr & " ;" & AREF
End If
ElseIf Len(AREF) <> 0 Then
Allpr = AREF
End If


'If Len(Allpr) <> 0 And Len(GFI) <> 0 Then
'Allpr = Allpr & " ;" & GFI
'ElseIf Len(GFI) <> 0 Then
'Allpr = GFI
'End If

Add1 = Hex(Address) & vbTab & CMDS & vbTab & Allpr





End Function
