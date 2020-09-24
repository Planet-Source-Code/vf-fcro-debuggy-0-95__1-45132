VERSION 5.00
Begin VB.Form HModules 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Search For Hidden Modules"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form33"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      ForeColor       =   &H00E7DFD6&
      Height          =   2430
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   5415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E7DFD6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status:"
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
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   5415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C89F8E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Base Address / Found New Modules"
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
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "HModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Sa As Long
Sa = 0
List1.Clear
Label2(1) = "Status:Searching...."
DoEvents
Do
StartSrc Sa
Sa = AddBy8(Sa, &H10000)
Loop Until Sa = 0
Label2(1) = "Status:Done"
NextB = 0
Form16.ReleaseShow 1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2

Dim Tabs() As Long
ReDim Tabs(1)
Tabs(0) = 0
Tabs(1) = 40
Call SendMessage(List1.hWnd, &H192, ByVal UBound(Tabs) + 1, Tabs(0))
Call SendMessage(List1.hWnd, &H194, ByVal 1000, ByVal 0&)
End Sub


Private Sub StartSrc(ByVal MA As Long)
On Error GoTo Dalje

Dim TName As String

If TestPTR(MA) = 0 Then Exit Sub

ExPs.ModuleName = ""


TName = FindInModules(MA)
If Len(TName) <> 0 Then: Exit Sub

Dim IMPFF As Byte
Dim EXPFF As Byte
ReadPE2 MA, IMPFF, EXPFF

If Len(ExPs.ModuleName) <> 0 Then
AddInActiveModules MA
AddInExportsSearch "", MA
List1.AddItem Hex(MA) & vbTab & ExPs.ModuleName
End If


Exit Sub
Dalje:
On Error GoTo 0
End Sub
