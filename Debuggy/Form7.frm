VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Registers/Flags"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3165
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   3165
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
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
      Height          =   270
      Index           =   14
      Left            =   600
      MaxLength       =   4
      TabIndex        =   65
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      Height          =   270
      Index           =   13
      Left            =   600
      MaxLength       =   4
      TabIndex        =   64
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      Height          =   270
      Index           =   12
      Left            =   600
      MaxLength       =   4
      TabIndex        =   63
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      Height          =   270
      Index           =   11
      Left            =   600
      MaxLength       =   4
      TabIndex        =   62
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      Height          =   270
      Index           =   10
      Left            =   600
      MaxLength       =   4
      TabIndex        =   61
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      Height          =   270
      Index           =   9
      Left            =   600
      MaxLength       =   4
      TabIndex        =   60
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   31
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   59
      Text            =   "GS="
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   30
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   58
      Text            =   "SS="
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   29
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   57
      Text            =   "CS="
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   28
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   56
      Text            =   "DS="
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   27
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   55
      Text            =   "ES="
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   26
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   54
      Text            =   "FS="
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox Text2 
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
      Height          =   270
      Index           =   16
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   53
      Top             =   3960
      Width           =   255
   End
   Begin VB.TextBox Text2 
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
      Height          =   270
      Index           =   15
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   52
      Top             =   3720
      Width           =   255
   End
   Begin VB.TextBox Text2 
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
      Height          =   270
      Index           =   14
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   51
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text2 
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
      Height          =   270
      Index           =   13
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   50
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   25
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "ID="
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   24
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "VIP="
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   23
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "VIF="
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   22
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "AC="
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Height          =   270
      Index           =   12
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   45
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox Text2 
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
      Height          =   270
      Index           =   11
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   44
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text2 
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
      Height          =   270
      Index           =   10
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   43
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text2 
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
      Height          =   270
      Index           =   9
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   42
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   21
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   41
      Text            =   "VM="
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   20
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "Resume="
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   19
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "Nested="
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   18
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "IOPL="
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change"
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
      Left            =   480
      TabIndex        =   37
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
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
      Left            =   1560
      TabIndex        =   36
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   17
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "OverFlow="
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Height          =   270
      Index           =   8
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   34
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   16
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "Direction="
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Height          =   270
      Index           =   7
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   32
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   15
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "Interupt="
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Height          =   270
      Index           =   6
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   30
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   14
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "Trap="
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
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
      Height          =   270
      Index           =   5
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   28
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   13
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "Sign="
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Height          =   270
      Index           =   4
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   26
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   12
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Zero="
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Height          =   270
      Index           =   3
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   24
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   11
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Auxiliary="
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Height          =   270
      Index           =   2
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   22
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   10
      Left            =   1680
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   21
      Text            =   "Parity="
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Height          =   270
      Index           =   1
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   20
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   9
      Left            =   1680
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   19
      Text            =   "Carry="
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Height          =   270
      Index           =   0
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   18
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   8
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "EBP="
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   7
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "ESP="
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   6
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "EDI="
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   5
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "ESI="
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   4
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "EDX="
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   3
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "ECX="
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   2
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "EBX="
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   1
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "EAX="
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   0
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "EIP="
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
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
      Height          =   270
      Index           =   8
      Left            =   600
      MaxLength       =   8
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      Height          =   270
      Index           =   7
      Left            =   600
      MaxLength       =   8
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      Height          =   270
      Index           =   6
      Left            =   600
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      Height          =   270
      Index           =   5
      Left            =   600
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      Height          =   270
      Index           =   4
      Left            =   600
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      Height          =   270
      Index           =   3
      Left            =   600
      MaxLength       =   8
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      Height          =   270
      Index           =   2
      Left            =   600
      MaxLength       =   8
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      Height          =   270
      Index           =   1
      Left            =   600
      MaxLength       =   8
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      Height          =   270
      Index           =   0
      Left            =   600
      MaxLength       =   8
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   66
      Top             =   2280
      Width           =   1695
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SCTX As CONTEXT
Public ShowTH As Long
Private LastKEip As Long

Private Sub Chg()
On Error GoTo Dalje
SCTX.Eip = CLng("&H" & Text1(0))
SCTX.Eax = CLng("&H" & Text1(1))
SCTX.Ebx = CLng("&H" & Text1(2))
SCTX.Ecx = CLng("&H" & Text1(3))
SCTX.Edx = CLng("&H" & Text1(4))
SCTX.Esi = CLng("&H" & Text1(5))
SCTX.Edi = CLng("&H" & Text1(6))
SCTX.Esp = CLng("&H" & Text1(7))
SCTX.Ebp = CLng("&H" & Text1(8))

SCTX.SegCs = CLng("&H" & Text1(9))
SCTX.SegDs = CLng("&H" & Text1(10))
SCTX.SegEs = CLng("&H" & Text1(11))
SCTX.SegFs = CLng("&H" & Text1(12))
SCTX.SegGs = CLng("&H" & Text1(13))
SCTX.SegSs = CLng("&H" & Text1(14))



SCTX.EFlags = Text2(0)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(1)) * 4&)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(2)) * 16&)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(3)) * 64&)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(4)) * 128&)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(5)) * 256&)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(6)) * 512&)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(7)) * 1024&)
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(8)) * 2048&)


SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(9)) * &H1000&) 'IOPL
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(10)) * &H2000&) 'NESTED TASK
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(11)) * &H8000&) 'RESUME FLAG
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(12)) * &H20000) 'VIRTUAL MODE
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(13)) * &H40000) 'ALIGNMENT CHECK
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(14)) * &H80000) 'VIRTUAL INTERUPT FLAG
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(15)) * &H100000) 'VIRTUAL INTERUPT PENDING
SCTX.EFlags = SCTX.EFlags Or (CLng(Text2(16)) * &H200000) 'ID FLAG

SetContext ShowTH, SCTX


'AddLastEip ShowTH, SCTX.Eip
ReadMem ActiveProcess, SCTX.Eip

If Form16.FBASE.ShowingTH = ShowTH And ActiveBasePosition <> SCTX.Ebp Then Form16.FBASE.ReadIt
If Form16.FSTACK.ShowingTH = ShowTH And ActiveStackPosition <> SCTX.Esp Then Form16.FSTACK.ReadIt

'Dodao!****
If LastKEip <> SCTX.Eip Then

If ISBPDisabled = 0 And ACTIVEBREAKPOINTS.count <> 0 Then
TRIGGERADDRESS = 0
TRIGGERFLAG = -1
RestoreAllBreakPoints ACTIVEBREAKPOINTS
End If

If ConfigData(7) = 1 Then
TRIGGERADDRESS = 0
TRIGGERFLAG = -1
RestoreAllBreakPoints SPYBREAKPOINTS
End If

End If
'***************

MsgBox "Thread:" & ShowTH & ", Registers changed!", vbInformation, "Information"


Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Error in Values!", vbCritical, "Error!"
ReadIt
End Sub







Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim ThrS As Long
Dim HT As Long
HT = GetHandleOfThread(ShowTH, ThrS)
CheckThreadStt HT, ThrS
End Sub

Private Sub Form_Load()
RemoveX hWnd
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2



End Sub

Public Sub ReadIt()
Dim ThrS As Long
Dim HT As Long
HT = GetHandleOfThread(ShowTH, ThrS)
If HT = 0 Then
Caption = "Thread:" & ShowTH & " ,Not Valid"

ElseIf ThrS = 1 Then
Caption = "Thread:" & ShowTH & " ,Runing"

Else
Caption = "Thread:" & ShowTH

SCTX = GetContext(ShowTH)
Text1(0) = Hex(SCTX.Eip)
Text1(1) = Hex(SCTX.Eax)
Text1(2) = Hex(SCTX.Ebx)
Text1(3) = Hex(SCTX.Ecx)
Text1(4) = Hex(SCTX.Edx)
Text1(5) = Hex(SCTX.Esi)
Text1(6) = Hex(SCTX.Edi)
Text1(7) = Hex(SCTX.Esp)
Text1(8) = Hex(SCTX.Ebp)

Text1(9) = Hex(SCTX.SegCs)
Text1(10) = Hex(SCTX.SegDs)
Text1(11) = Hex(SCTX.SegEs)
Text1(12) = Hex(SCTX.SegFs)
Text1(13) = Hex(SCTX.SegGs)
Text1(14) = Hex(SCTX.SegSs)

Text2(0) = (SCTX.EFlags And 1&)
Text2(1) = ((SCTX.EFlags And 4&) / 4&)
Text2(2) = ((SCTX.EFlags And 16&) / 16&)
Text2(3) = ((SCTX.EFlags And 64&) / 64&)
Text2(4) = ((SCTX.EFlags And 128&) / 128&)
Text2(5) = ((SCTX.EFlags And 256&) / 256&)
Text2(6) = ((SCTX.EFlags And 512&) / 512&)
Text2(7) = ((SCTX.EFlags And 1024&) / 1024&)
Text2(8) = ((SCTX.EFlags And 2048&) / 2048&)

Text2(9) = ((SCTX.EFlags And &H1000&) / &H1000&)
Text2(10) = ((SCTX.EFlags And &H2000&) / &H2000&)
Text2(11) = ((SCTX.EFlags And &H8000&) / &H8000&)
Text2(12) = ((SCTX.EFlags And &H20000) / &H20000)
Text2(13) = ((SCTX.EFlags And &H40000) / &H40000)
Text2(14) = ((SCTX.EFlags And &H80000) / &H80000)
Text2(15) = ((SCTX.EFlags And &H100000) / &H100000)
Text2(16) = ((SCTX.EFlags And &H200000) / &H200000)



LastKEip = SCTX.Eip
End If
End Sub








Private Sub CheckThreadStt(ByRef HT As Long, ByRef ThrS As Long)
If ThrS = 1 Then
Caption = "Thread:" & ShowTH & " ,Running!"
ElseIf HT = 0 Then
Caption = "Thread:" & ShowTH & " ,Not valid now!"
Else
Caption = "Thread:" & ShowTH
Chg
End If
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 8 Then Exit Sub
If IsValidK(Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 49 Then KeyAscii = 0: Exit Sub
Text2(Index).SelStart = 0
Text2(Index).SelLength = 1
Text2(Index).SelText = Chr(KeyAscii)
End Sub
