VERSION 5.00
Begin VB.Form Form25 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PE Header Viewer"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5730
   LinkTopic       =   "Form25"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E7DFD6&
      Height          =   6735
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TBLS() As String
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
ReDim TBLS(15)
TBLS(0) = "Export Table"
TBLS(1) = "Import Table"
TBLS(2) = "Resource Table"
TBLS(3) = "Exception Table"
TBLS(4) = "Certificate Table"
TBLS(5) = "Relocation Table"
TBLS(6) = "Debug Data"
TBLS(7) = "Architecture Data"
TBLS(8) = "Machine Value"
TBLS(9) = "TLS Table"
TBLS(10) = "Load Configuration Table"
TBLS(11) = "Bound Import Table"
TBLS(12) = "Import Address Table"
TBLS(13) = "Delay Import Descriptor"
TBLS(14) = "COM+ Runtime Header"
TBLS(15) = "Reserved"
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
LockWindowUpdate Text1.hWnd
ReadIt
LockWindowUpdate 0
End Sub


Private Sub ReadIt()
Text1 = ""
AddTXT "-----------------"
AddTXT "PE File Header"
AddTXT "-----------------"
AddTXT "Machine:" & Hex(NTHEADER.FileHeader.Machine)
AddTXT "Number Of Sections:" & Hex(NTHEADER.FileHeader.NumberOfSections)
AddTXT "Number Of Symbols:" & Hex(NTHEADER.FileHeader.NumberOfSymbols)
AddTXT "Pointer To Symbol Table:" & Hex(NTHEADER.FileHeader.PointerToSymbolTable)
AddTXT "Size Of Optional Header:" & Hex(NTHEADER.FileHeader.SizeOfOptionalHeader)
AddTXT "Time Date Stamp:" & Hex(NTHEADER.FileHeader.TimeDateStamp)
AddTXT "Characteristic:" & Hex(NTHEADER.FileHeader.Characteristics)

AddTXT ""
AddTXT "-----------------------"
AddTXT "PE Optional Header"
AddTXT "-----------------------"
AddTXT "Magic:" & Hex(NTHEADER.OptionalHeader.Magic)
AddTXT "Minor Linker Version:" & Hex(NTHEADER.OptionalHeader.MinorLinkerVersion)
AddTXT "Major Linker Version:" & Hex(NTHEADER.OptionalHeader.MajorLinkerVersion)
AddTXT "Size Of Code:" & Hex(NTHEADER.OptionalHeader.SizeOfCode)
AddTXT "Size Of Initialized Data:" & Hex(NTHEADER.OptionalHeader.SizeOfInitializedData)
AddTXT "Size Of Uninitialized Data:" & Hex(NTHEADER.OptionalHeader.SizeOfUninitializedData)
AddTXT "Address Of Entry Point:" & Hex(AddBy8(NTHEADER.OptionalHeader.AddressOfEntryPoint, NTHEADER.OptionalHeader.ImageBase))
AddTXT "Base Of Code:" & Hex(NTHEADER.OptionalHeader.BaseOfCode)
AddTXT "Base Of Data:" & Hex(NTHEADER.OptionalHeader.BaseOfData)
AddTXT "Image Base:" & Hex(NTHEADER.OptionalHeader.ImageBase)
AddTXT "Section Alignment:" & Hex(NTHEADER.OptionalHeader.SectionAlignment)
AddTXT "Size Of Image:" & Hex(NTHEADER.OptionalHeader.SizeOfImage)
AddTXT "Size Of Headers:" & Hex(NTHEADER.OptionalHeader.SizeOfHeaders)
AddTXT "Size Of Stack Reserve:" & Hex(NTHEADER.OptionalHeader.SizeOfStackReserve)
AddTXT "Size Of Stack Commit:" & Hex(NTHEADER.OptionalHeader.SizeOfStackCommit)
AddTXT "Size Of Heap Reserve:" & Hex(NTHEADER.OptionalHeader.SizeOfHeapReserve)
AddTXT "Size Of Heap Commit:" & Hex(NTHEADER.OptionalHeader.SizeOfHeapCommit)
AddTXT "CheckSum:" & Hex(NTHEADER.OptionalHeader.CheckSum)
AddTXT "DLL Characteristic:" & Hex(NTHEADER.OptionalHeader.DllCharacteristics)
AddTXT "SubSystem:" & Hex(NTHEADER.OptionalHeader.Subsystem)
AddTXT "Minor Image Version:" & Hex(NTHEADER.OptionalHeader.MinorImageVersion)
AddTXT "Major Image Version:" & Hex(NTHEADER.OptionalHeader.MajorImageVersion)
AddTXT "Minor Operating System Version:" & Hex(NTHEADER.OptionalHeader.MinorOperatingSystemVersion)
AddTXT "Major Operating System Version:" & Hex(NTHEADER.OptionalHeader.MajorOperatingSystemVersion)
AddTXT "Minor Subsystem Version:" & Hex(NTHEADER.OptionalHeader.MinorSubsystemVersion)
AddTXT "Major Subsystem Version:" & Hex(NTHEADER.OptionalHeader.MajorSubsystemVersion)
AddTXT "Win 32 Version Value:" & Hex(NTHEADER.OptionalHeader.Win32VersionValue)
AddTXT "Loader Flags:" & NTHEADER.OptionalHeader.LoaderFlags
AddTXT "Number of Data Directories:" & Hex(NTHEADER.OptionalHeader.NumberOfRvaAndSizes)
AddTXT ""
AddTXT "----------------------"
AddTXT "Sections (Objects)"
AddTXT "----------------------"
Dim u As Long
Dim TString As String
For u = 0 To UBound(SECTIONSHEADER)
SLen = lstrlen(ByVal SECTIONSHEADER(u).nameSec)
TString = Space(SLen)
CopyMemory ByVal TString, ByVal SECTIONSHEADER(u).nameSec, SLen
AddTXT "Section: " & TString
AddTXT "Pointer To Raw Data:" & Hex(SECTIONSHEADER(u).PointerToRawData)
AddTXT "Size Of Raw Data:" & Hex(SECTIONSHEADER(u).SizeOfRawData)
AddTXT "Pointer To RVA Data:" & Hex(AddBy8(NTHEADER.OptionalHeader.ImageBase, SECTIONSHEADER(u).VirtualAddress))
AddTXT "Size Of RVA Data:" & Hex(SECTIONSHEADER(u).VirtualSize)
AddTXT "Characteristic:" & Hex(SECTIONSHEADER(u).Characteristics)
AddTXT ""
Next u

AddTXT ""
AddTXT "-------------------"
AddTXT "Data Directory"
AddTXT "-------------------"

For u = 0 To 15
AddTXT TBLS(u)
If NTHEADER.OptionalHeader.DataDirectory(u).size <> 0 Then
AddTXT "Address:" & Hex(AddBy8(NTHEADER.OptionalHeader.ImageBase, NTHEADER.OptionalHeader.DataDirectory(u).VirtualAddress))
AddTXT "Size:" & Hex(NTHEADER.OptionalHeader.DataDirectory(u).size)
AddTXT ""
Else
AddTXT "* Not Used *"
AddTXT ""
End If
Next u

End Sub


Private Sub AddTXT(ByRef StringX As String)
Text1 = Text1 & StringX & vbCrLf
End Sub

Private Sub Form_Unload(Cancel As Integer)
Erase TBLS
Text1 = ""
End Sub
