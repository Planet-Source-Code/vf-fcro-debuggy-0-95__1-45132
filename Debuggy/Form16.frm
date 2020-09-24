VERSION 5.00
Begin VB.Form Form16 
   Caption         =   "Disassembler"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form16.frx":0000
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command25 
      Caption         =   "Spy Config"
      Height          =   375
      Left            =   9120
      TabIndex        =   84
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Trace Mode"
      Height          =   375
      Left            =   9120
      TabIndex        =   83
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command23 
      Caption         =   "File Editor"
      Height          =   375
      Left            =   7560
      TabIndex        =   82
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Debug Log"
      Height          =   375
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Watches"
      Height          =   375
      Left            =   6240
      TabIndex        =   81
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Call Stack"
      Height          =   375
      Left            =   4920
      TabIndex        =   80
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command19 
      Caption         =   "ESP Browser"
      Height          =   375
      Left            =   4920
      TabIndex        =   79
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command18 
      Caption         =   "EBP Browser"
      Height          =   375
      Left            =   4920
      TabIndex        =   78
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   25
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   76
      Top             =   6240
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   24
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   75
      Top             =   6000
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   23
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   74
      Top             =   5760
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   22
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   73
      Top             =   5520
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   21
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   72
      Top             =   5280
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   20
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   71
      Top             =   5040
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   19
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   70
      Top             =   4800
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   18
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   69
      Top             =   4560
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   17
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   68
      Top             =   4320
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   16
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   67
      Top             =   4080
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   15
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   3840
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   14
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   65
      Top             =   3600
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   13
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   3360
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   12
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   3120
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   11
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   2880
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   10
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   2640
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   9
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   2400
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   8
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   2160
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   7
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   1920
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   6
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   1680
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   5
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   1440
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   4
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   1200
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   3
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   960
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   2
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   720
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   1
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   480
      Width           =   5895
   End
   Begin VB.TextBox rt2 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   0
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   240
      Width           =   5895
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   50
      Top             =   6240
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   49
      Top             =   6000
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   48
      Top             =   5760
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   47
      Top             =   5520
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   46
      Top             =   5280
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   45
      Top             =   5040
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   44
      Top             =   4800
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   43
      Top             =   4560
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   42
      Top             =   4320
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   41
      Top             =   4080
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   40
      Top             =   3840
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   39
      Top             =   3600
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   38
      Top             =   3360
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   37
      Top             =   3120
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   36
      Top             =   2880
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   35
      Top             =   2640
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   34
      Top             =   2400
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   33
      Top             =   2160
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   32
      Top             =   1920
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   31
      Top             =   1680
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   30
      Top             =   1440
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   1200
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   28
      Top             =   960
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   720
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   480
      Width           =   5175
   End
   Begin VB.TextBox rt1 
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
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   240
      Width           =   5175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Dbg Config"
      Height          =   375
      Left            =   9120
      TabIndex        =   24
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Modules"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Current EIP"
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
      Left            =   10440
      TabIndex        =   22
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Registers"
      Height          =   375
      Left            =   4920
      TabIndex        =   21
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Memory"
      Height          =   375
      Left            =   6240
      TabIndex        =   20
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Disasm/Cache"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop Debug"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Cached Imports"
      Height          =   375
      Left            =   7560
      TabIndex        =   16
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Cached Strings"
      Height          =   375
      Left            =   7560
      TabIndex        =   15
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Windows"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "BreakPoints"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8160
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
      ForeColor       =   &H000000FF&
      Height          =   1470
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   7080
      Width           =   2175
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Terminate"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Resume"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Suspend"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continue"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Single Step"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10440
      MaxLength       =   8
      TabIndex        =   4
      Top             =   6960
      Width           =   1335
   End
   Begin VB.VScrollBar vs1 
      Height          =   6255
      Left            =   11520
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.ListBox List8 
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
      Height          =   6300
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E7DFD6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Selected:"
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   0
      TabIndex        =   77
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00DCB17C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Threads in Process"
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label LabelX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E7DFD6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Goto Address:"
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
      Left            =   10440
      TabIndex        =   5
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Break"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00DCB17C&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public FBASE As New Form17
Public FSTACK As New Form17










Private Sub Command10_Click()
If ActiveThread = 0 Then Unload Form7: MsgBox "Select Thread First", vbExclamation, "Information": Exit Sub

Dim ThrS As Long
Dim HT As Long
HT = GetHandleOfThread(ActiveThread, ThrS)
If HT = 0 Then Unload Form7: MsgBox "Select Thread First", vbExclamation, "Information": Exit Sub
If ThrS = 1 Then Unload Form7: MsgBox "Thread Running,cannot display", vbExclamation, "Information": Exit Sub


Form7.ShowTH = ActiveThread
Form7.ReadIt
OnScreen Form7.hWnd
Form7.Show

End Sub

Private Sub Command11_Click()
Form12.Show
Form12.Visible = True
OnScreen Form12.hWnd
End Sub

Private Sub Command12_Click()

Dim isCC As Long
isCC = ISBPDisabled

'RestoreAllOriginalBytes SPYBREAKPOINTS
'If ConfigData(7) = 1 Then
'Set RTRIGGER = Nothing
TRIGGERADDRESS = 0
TRIGGERFLAG = -1
RestoreAllOriginalBytes SPYBREAKPOINTS
'End If



If isCC = 0 And ACTIVEBREAKPOINTS.count <> 0 Then
'Set RTRIGGER = Nothing
TRIGGERADDRESS = 0
TRIGGERFLAG = -1
RestoreAllOriginalBytes ACTIVEBREAKPOINTS
ISBPDisabled = 1
Label4.BackColor = &HFF&
MsgBox "All Breakpoints disabled!", vbInformation, "Information!"
End If

Form13.Show 1


If isCC = 0 And ACTIVEBREAKPOINTS.count <> 0 Then
RestoreAllBreakPoints ACTIVEBREAKPOINTS
ISBPDisabled = 0
Label4.BackColor = &HAA00&
End If

'If ConfigData(7) = 1 Then
RestoreAllBreakPoints SPYBREAKPOINTS
'End If

End Sub

Private Sub Command15_Click()
If SINDEXESR.count = 0 Then MsgBox "Nothing cached yet!", vbExclamation, "Information": Exit Sub
ChoosedAdr = 0
Form15.FRMTYPE = 1
Form15.Caption = "(Possible) Strings In " & ValidCRef
Form15.Show 1
If ChoosedAdr <> 0 Then NextB = 0: DISCOUNT = ChoosedAdr: ReleaseShow 1

End Sub



Private Sub Command16_Click()
If ActiveThread = 0 Then Unload Form18: MsgBox "Select Thread First", vbExclamation, "Information": Exit Sub
Dim ThrS As Long
Dim HT As Long
HT = GetHandleOfThread(ActiveThread, ThrS)
If HT = 0 Then MsgBox "Select Thread First", vbExclamation, "Information": Exit Sub
If ThrS = 1 Then MsgBox "Thread Running,cannot display", vbExclamation, "Information": Exit Sub
CTX = GetContext(ActiveThread)
DISCOUNT = CTX.Eip
NextB = 0
ReleaseShow 1
End Sub

Private Sub Command17_Click()
ChoosedAdr = 0
Form2.Show 1
If ChoosedAdr <> 0 Then NextB = 0: DISCOUNT = ChoosedAdr: ReleaseShow 1
End Sub
Public Sub InSuspend()
Command1_Click
End Sub





Private Sub Command18_Click()
If ActiveThread = 0 Then Unload FBASE: MsgBox "Select Thread First", vbExclamation, "Information": Exit Sub

Dim ThrS As Long
Dim HT As Long
HT = GetHandleOfThread(ActiveThread, ThrS)
If HT = 0 Then Unload FBASE: MsgBox "Select Thread First", vbExclamation, "Information": Exit Sub
If ThrS = 1 Then Unload FBASE: MsgBox "Thread Running,cannot display", vbExclamation, "Information": Exit Sub


FBASE.WTypeS = 0
FBASE.ShowingTH = ActiveThread
FBASE.Show
End Sub

Private Sub Command19_Click()
If ActiveThread = 0 Then Unload FSTACK: MsgBox "Select Thread First", vbExclamation, "Information": Exit Sub

Dim ThrS As Long
Dim HT As Long
HT = GetHandleOfThread(ActiveThread, ThrS)
If HT = 0 Then Unload FSTACK: MsgBox "Select Thread First", vbExclamation, "Information": Exit Sub
If ThrS = 1 Then Unload FSTACK: MsgBox "Thread Running,cannot display", vbExclamation, "Information": Exit Sub

FSTACK.WTypeS = 1
FSTACK.ShowingTH = ActiveThread
FSTACK.Show
End Sub

Private Sub Command20_Click()
If ActiveThread = 0 Then Unload Form18: MsgBox "Select Thread First", vbExclamation, "Information": Exit Sub

Dim ThrS As Long
Dim HT As Long
HT = GetHandleOfThread(ActiveThread, ThrS)
If HT = 0 Then Unload Form18: MsgBox "Select Thread First", vbExclamation, "Information": Exit Sub
If ThrS = 1 Then Unload Form18: MsgBox "Thread Running,cannot display", vbExclamation, "Information": Exit Sub



Form18.ShowingTH = ActiveThread
Form18.Show
Form18.ReadIt
End Sub





Private Sub Command21_Click()
Form24.Show
Form24.ReadExpresses
End Sub





Private Sub Command23_Click()
Form27.Show 1
End Sub

Private Sub Command24_Click()

If ActiveThread = 0 Then MsgBox "Select Thread First", vbExclamation, "Information": Exit Sub

Dim ThrS As Long
Dim HT As Long
HT = GetHandleOfThread(ActiveThread, ThrS)
If HT = 0 Then MsgBox "Thread becomes invalid now!", vbExclamation, "Information": Exit Sub
If ThrS <> 2 Then MsgBox "Trace Mode works only from Breakpoint!", vbInformation, "Information": Exit Sub



isCC = ISBPDisabled

CopyMemory ByVal VarPtr(TemConfig(0)), ByVal VarPtr(ConfigData(0)), 80


'RestoreAllOriginalBytes SPYBREAKPOINTS
'If ConfigData(7) = 1 Then
'Set RTRIGGER = Nothing
TRIGGERFLAG = -1
TRIGGERADDRESS = 0
RestoreAllOriginalBytes SPYBREAKPOINTS
'End If

If isCC = 0 And ACTIVEBREAKPOINTS.count <> 0 Then
'Set RTRIGGER = Nothing
TRIGGERFLAG = -1
TRIGGERADDRESS = 0
RestoreAllOriginalBytes ACTIVEBREAKPOINTS
ISBPDisabled = 1
Label4.BackColor = &HFF&
End If

Dim u As Long
For u = 0 To 20
ConfigData(u) = 0
Next u

UseTrace = 1
CTX = GetContext(ActiveThread)
Form28.Caption = "Thread Tracer will start from EIP:" & Hex(CTX.Eip)
Form28.Show 1

End Sub

Private Sub Command25_Click()
Form30.Show 1
End Sub

Private Sub Command4_Click()
Form10.Show 1
End Sub






Private Sub Form_Load()
TRIGGERFLAG = -1
Label4.BackColor = &HAA00&
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2

RemoveX hWnd
VS1.Max = 32767
VS1.Min = 0
VS1.Value = 16384
VS1.SmallChange = 1
VS1.LargeChange = 10

Dim u As Long
For u = 0 To 25
List8.AddItem ""
Next u



Dim Tabs() As Long
ReDim Tabs(1)
Tabs(0) = 0
Tabs(1) = 37
For u = 0 To 25
Call SendMessage(rt1(u).hWnd, &HCB, ByVal UBound(Tabs) + 1, Tabs(0))
Next u


End Sub





Private Sub Form_Unload(Cancel As Integer)
Unload Form8
Unload Form12
Unload Form11

Unload FBASE
Unload FSTACK
Unload Form7
Unload Form18
Form4.Enabled = True
Form4.Visible = True
End Sub

Private Sub List1_Click()
If List1.ListIndex = -1 Or List1.ListCount = 0 Then Exit Sub
Dim SXth() As String
SXth = Split(List1.List(List1.ListIndex), ",")

ActiveThread = CLng(SXth(0))
Label2 = "Selected: " & ActiveThread
Dim ThrS As Long
Dim HT As Long
HT = GetHandleOfThread(ActiveThread, ThrS)

If ThrS = 0 Or ThrS = 2 Then
CTX = GetContext(ActiveThread)
ReadMem ActiveProcess, CTX.Eip
FBASE.ShowingTH = ActiveThread
FBASE.ReadIt
FSTACK.ShowingTH = ActiveThread
FSTACK.ReadIt
Form18.ShowingTH = ActiveThread
Form18.ReadIt
End If
End Sub

Private Sub List8_Click()
Dim IsValidBP As Byte
Dim AdrR As Long
Dim S() As String
S = Split(rt1(List8.ListIndex), vbTab)
AdrR = CLng("&H" & S(0))

Dim IsValidMemPTR As Byte
Dim BTTX As Byte
IsValidMemPTR = TestPTR(AdrR, BTTX)
If IsValidMemPTR = 0 Then MsgBox "Cannot put Breakpoint at:" & Hex(AdrR), vbCritical, "Information": Exit Sub


If AdrR >= DEBUGGYFA And AdrR <= DEBUGGYLA Then GoTo AccDen

Call GetBreakPoint(SPYBREAKPOINTS, AdrR, IsValidBP)
If IsValidBP = 1 Then
AccDen:
MsgBox "Access Denied by Debugger itself!", vbExclamation, "Information": Exit Sub
End If

Call GetBreakPoint(ACTIVEBREAKPOINTS, AdrR, IsValidBP)
If IsValidBP = 0 Then
AddBreakPoint ACTIVEBREAKPOINTS, AdrR, ISBPDisabled
List8.List(List8.ListIndex) = "*BP*"
Else
RemoveBreakPoint ACTIVEBREAKPOINTS, AdrR
List8.List(List8.ListIndex) = ""
End If

If AdrR >= ActiveMemPos Then PrintDump Form8.TextX, ActiveMemPos
End Sub
Private Sub List8_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub
Private Sub rt1_DblClick(Index As Integer)


Select Case NotifyData1(Index)

Case Is = 1, 2, 3, 4, 8
If NotifyData2(Index) = 0 Then Exit Sub
If vbYes = MsgBox("Jump To Address:" & Hex(NotifyData2(Index)), vbYesNo, "Confirm") Then
DISCOUNT = NotifyData2(Index)
NextB = 0
ReleaseShow 1
End If

End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo Dalje
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 8 Then
Exit Sub
ElseIf KeyAscii = 13 Then
If Len(Text1) = 0 Then Text1 = "": Exit Sub
DISCOUNT = CLng("&H" & Text1): NextB = 0: ReleaseShow 1
End If

If IsValidK(Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Unknown Value Type", vbCritical, "Error"
End Sub

Private Sub vs1_Change()
Static IsS As Byte
If IsS = 1 Then IsS = 0: Exit Sub

If VS1.Value = 16383 Then
ReleaseShow 0

ElseIf VS1.Value = 16385 Then
ReleaseShow 1

ElseIf VS1.Value < 16383 Then
AddBackward25 rt1, rt2, 25, ActiveProcess, List8

ElseIf VS1.Value > 16385 Then
AddForward25 rt1, rt2, 25, ActiveProcess, List8


End If


IsS = 1
VS1.Value = 16384

End Sub

Public Sub ReleaseShow(ByVal Way As Byte)


Dim TA As Long 'Temp
Dim TA2 As Long 'Temp
Dim TName As String
Dim ICtc As String

If Way = 0 Then
AddBackward rt1, rt2, 25, ActiveProcess, List8



ElseIf Way = 1 Then
DISCOUNT = DISCOUNT + NextB
AddForward rt1, rt2, 25, ActiveProcess, List8



End If
End Sub


Private Sub Command2_Click()
If ActiveThread = 0 Then MsgBox "Select Thread First", vbExclamation, "Information": Exit Sub

Dim ChkThreadR As Long
ChkThreadR = IsRunningThread(ActiveThread)

If ChkThreadR = 0 Then
MsgBox "Thread is suspended!", vbExclamation, "Information": Exit Sub
ElseIf ChkThreadR = 1 Then
MsgBox "Thread already running!", vbExclamation, "Information": Exit Sub
End If



Dim LastEip As Long

'Dim TRIGGERFLAG As Long
'Dim IsValidTrigger As Byte
'Call GetFromTrigger(ActiveThread, IsValidTrigger, TRIGGERFLAG)

CTX = GetContext(ActiveThread)
LastEip = CTX.Eip
'If IsValidTrigger = 1 Then
If TRIGGERFLAG <> -1 Then
'SetFlagInTrigger ActiveThread, LastEip, 1
TRIGGERADDRESS = LastEip
TRIGGERFLAG = 1
SetSingleStep ActiveThread
End If

'Dodao
'If Form8.Visible = True Then
'PrintDump Form8.TextX, ActiveMemPos
'If gBegAdr <> 0 And gLenAdr <> 0 Then
'GetDataFromMem gBegAdr, DataPW, gLenAdr
'End If
'MEMINF = QueryMem(ActiveMemPos, MEMStr)
'Form8.Text4 = MEMStr
'End If


ChangeStateThread ActiveThread, 1
ReadThreadsFromProcess List1

If Unhandled = 1 Then
Unhandled = 0
ContinueDebugNotHandle
Else
ContinueDebug
End If

End Sub
Private Sub Command6_Click()
If ActiveThread = 0 Then MsgBox "Select Thread First", vbExclamation, "Information": Exit Sub

Dim ChkThreadR As Long
ChkThreadR = IsRunningThread(ActiveThread)

If ChkThreadR = 0 Then
MsgBox "Thread is suspended!", vbExclamation, "Information": Exit Sub
ElseIf ChkThreadR = 1 Then
MsgBox "Thread already running!", vbExclamation, "Information": Exit Sub
End If



'LastEip = GetLastEip(ActiveThread)
CTX = GetContext(ActiveThread)
LastEip = CTX.Eip
SetSingleStep ActiveThread

Dim IsValidBP As Byte
Call GetBreakPoint(SPYBREAKPOINTS, LastEip, IsValidBP)

If IsValidBP = 1 Then

RestoreOriginalBytes SPYBREAKPOINTS, LastEip
'SetInTrigger ActiveThread, LastEip

'Promjenio!
'TRIGGERFLAG = -1
'TRIGGERADDRESS = 0
TRIGGERFLAG = 0
TRIGGERADDRESS = LastEip

Else
Call GetBreakPoint(ACTIVEBREAKPOINTS, LastEip, IsValidBP)
'Treba unapred izmjeniti za razliku od continue koji se
'ne smije izmjeniti unapred...
'Takodjer je to potrebno je ce se dogoditi Exception umjesto single stepa..
If IsValidBP = 1 Then
RestoreOriginalBytes ACTIVEBREAKPOINTS, LastEip
'SetInTrigger ActiveThread, LastEip
TRIGGERFLAG = 0
TRIGGERADDRESS = LastEip
End If
End If

'Dodao
'If Form8.Visible = True Then
'PrintDump Form8.TextX, ActiveMemPos
'If gBegAdr <> 0 And gLenAdr <> 0 Then
'GetDataFromMem gBegAdr, DataPW, gLenAdr
'End If
'MEMINF = QueryMem(ActiveMemPos, MEMStr)
'Form8.Text4 = MEMStr
'End If

ChangeStateThread ActiveThread, 1
ReadThreadsFromProcess List1

If Unhandled = 1 Then
Unhandled = 0
ContinueDebugNotHandle
Else
ContinueDebug
End If
End Sub
Private Sub Command1_Click()
Dim ThrS As Long
Dim HT As Long
HT = GetHandleOfThread(ActiveThread, ThrS)

If HT = 0 Then Exit Sub
If ThrS = 1 Then

SuspendThread HT



ChangeStateThread ActiveThread, 0
ReadThreadsFromProcess List1
ElseIf ThrS = 0 Then
MsgBox "Thread is already suspended!", vbExclamation, "Information": Exit Sub
Else
MsgBox "Thread waiting!", vbExclamation, "Information": Exit Sub
End If



If Form8.Visible = True Then
PrintDump Form8.TextX, ActiveMemPos
If gBegAdr <> 0 And gLenAdr <> 0 Then
GetDataFromMem gBegAdr, DataPW, gLenAdr
End If
MEMINF = QueryMem(ActiveMemPos, MEMStr)
Form8.Text4 = MEMStr
End If


CTX = GetContext(ActiveThread)
ReadMem ActiveProcess, CTX.Eip
'AddLastEip ActiveThread, CTX.Eip
TouchIt ActiveThread

End Sub
Private Sub Command3_Click()
Dim HT As Long
Dim ThrS As Long
HT = GetHandleOfThread(ActiveThread, ThrS)
If HT = 0 Then Exit Sub
If ThrS = 0 Then
ResumeThread HT
ChangeStateThread ActiveThread, 1
ReadThreadsFromProcess List1
'RemoveLastEip ActiveThread
ElseIf ThrS = 1 Then
MsgBox "Thread already running!", vbExclamation, "Information"
Else
MsgBox "Thread waiting!", vbExclamation, "Information"
End If
TouchIt ActiveThread
End Sub
Private Sub Command13_Click()
If SusThreadX = 0 Then MsgBox "Cannot Terminate Thread!", vbCritical, "Information!": Exit Sub
Call CreateRemoteThread(ProcessHandle, ByVal 0&, 10, ByVal SusThreadX, ByVal ActiveThread, 0, AccThreadX)

End Sub
Private Sub Command9_Click()
If ACTIVEBREAKPOINTS.count = 0 Then MsgBox "There is no Breakpoints!", vbInformation, "Information": Exit Sub
ChoosedAdr = 0
Form5.Show 1
If ChoosedAdr <> 0 Then NextB = 0: DISCOUNT = ChoosedAdr: ReleaseShow 1

End Sub
Private Sub Command8_Click()
If WINS.count = 0 Then MsgBox "There is no any Windows!", vbInformation, "Information": Exit Sub
Form11.Show
End Sub

Private Sub Command14_Click()
If EINDEXESR.count = 0 Then MsgBox "Nothing cached yet!", vbExclamation, "Information": Exit Sub
Form15.FRMTYPE = 0
ChoosedAdr = 0
Form15.Caption = "Imports Calling For " & ValidCRef
Form15.Show 1

If ChoosedAdr <> 0 Then NextB = 0: DISCOUNT = ChoosedAdr: ReleaseShow 1

End Sub
Private Sub Command5_Click()
StopAndClear

Unload Me
End Sub
Private Sub Command7_Click()
Form8.Show
OnScreen Form8.hWnd
End Sub
