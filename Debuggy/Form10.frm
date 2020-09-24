VERSION 5.00
Begin VB.Form Form10 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Debugger Configuration"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2565
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   2565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Break On Menu Click"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enable API Spy"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enable WM Breakpoints"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   2535
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
      Left            =   1320
      TabIndex        =   6
      Top             =   3000
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
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Break On Destroy Window"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Break On Create Window"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Break On Create Thread"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Break On UnLoad DLL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Break On Load DLL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()


ConfigData(0) = Check1(0).Value
ConfigData(1) = Check1(1).Value
ConfigData(2) = Check1(2).Value
ConfigData(4) = Check1(3).Value
ConfigData(5) = Check1(4).Value
ConfigData(6) = Check1(5).Value
ConfigData(7) = Check1(6).Value
ConfigData(8) = Check1(7).Value

If ConfigData(7) = 0 Then

'Dim ThrS As Long
'Dim HT As Long
'HT = GetHandleOfThread(ActiveThread, ThrS)
'If ThrS = 1 And HT <> 0 Then
'SuspendThread HT
'End If
'FreezeAll

TRIGGERADDRESS = 0
TRIGGERFLAG = -1
RestoreAllOriginalBytes SPYBREAKPOINTS

'If ThrS = 1 And HT <> 0 Then
'ResumeThread HT
'End If

'UnFreezeAll

Else
RestoreAllBreakPoints SPYBREAKPOINTS
End If





Unload Me
End Sub



Private Sub Command2_Click()
Unload Me
End Sub



Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Check1(0).Value = ConfigData(0)
Check1(1).Value = ConfigData(1)
Check1(2).Value = ConfigData(2)
Check1(3).Value = ConfigData(4)
Check1(4).Value = ConfigData(5)
Check1(5).Value = ConfigData(6)
Check1(6).Value = ConfigData(7)
Check1(7).Value = ConfigData(8)
End Sub
