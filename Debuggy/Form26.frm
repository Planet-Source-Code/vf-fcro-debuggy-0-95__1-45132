VERSION 5.00
Begin VB.Form Form26 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calculator"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2235
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form26"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   2235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "RVA-RAW Converter"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
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
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   720
         MaxLength       =   8
         TabIndex        =   5
         Top             =   240
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
         Left            =   720
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   4
         Top             =   720
         Width           =   1215
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
         TabIndex        =   3
         Text            =   "RAW="
         Top             =   720
         Width           =   615
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "RVA="
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1440
      Width           =   855
   End
End
Attribute VB_Name = "Form26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo Dalje
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 8 Then Exit Sub
If KeyAscii = 13 Then
FindIt CLng("&H" & Text1)
End If

If IsValidK(Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub

Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Error in Value", vbCritical, "Error"
End Sub


Private Sub FindIt(ByVal Address As Long)
Dim MmStr As String
Dim ImpX As Byte
Dim ExpX As Byte
Dim BaseA As Long

Dim TString As String


MmStr = FindInModules(Address, BaseA)
If Len(MmStr) = 0 Then MsgBox "Address " & Hex(Address) & " doesn't belongs to modules in this process!", vbInformation, "Information": Exit Sub
ReadPE2 BaseA, ImpX, ExpX

Dim u As Long
Dim AuAdr As Long
Dim Eadr As Long

For u = 0 To UBound(SECTIONSHEADER)

Eadr = AddBy8(BaseA, SECTIONSHEADER(u).VirtualAddress)
If Address >= Eadr And _
Address <= AddBy8(Eadr, SECTIONSHEADER(u).VirtualSize) Then
AuAdr = SubBy8(Address, Eadr)

If AuAdr <= AddBy8(SECTIONSHEADER(u).PointerToRawData, SECTIONSHEADER(u).SizeOfRawData) Then

SLen = lstrlen(ByVal SECTIONSHEADER(u).nameSec)
TString = Space(SLen)
CopyMemory ByVal TString, ByVal SECTIONSHEADER(u).nameSec, SLen
MsgBox "Found in Module:" & MmStr & vbCrLf & "Section:" & TString, vbInformation, "Information"
Text3 = Hex(AddBy8(SECTIONSHEADER(u).PointerToRawData, AuAdr))
Exit Sub
End If


End If
Next u

Text3 = ""
MsgBox "Address not Found in Module: " & MmStr & " (Section information)", vbExclamation, "Information"
End Sub



