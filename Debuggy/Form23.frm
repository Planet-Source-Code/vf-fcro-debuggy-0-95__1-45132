VERSION 5.00
Begin VB.Form Form23 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menu Resource"
   ClientHeight    =   285
   ClientLeft      =   150
   ClientTop       =   675
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
   LinkTopic       =   "Form23"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   285
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu m1 
      Caption         =   "X"
   End
End
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MNUs As Long



Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Form_Resize
End Sub

Private Sub Form_Resize()
Call SetMenu(hWnd, MNUs)
DrawMenuBar hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
DestroyMenu MNUs
End Sub
