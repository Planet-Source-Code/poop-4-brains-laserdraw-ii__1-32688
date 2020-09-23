VERSION 5.00
Begin VB.Form frmVeiw 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "View Selected Picture"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5520
   Icon            =   "frmVeiw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   228
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   368
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox board 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   0
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmVeiw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub board_Click()
Unload Me
End Sub

Private Sub board_Resize()
Me.Width = board.Width * Screen.TwipsPerPixelX
Me.Height = board.Height * (Screen.TwipsPerPixelY * 1.2)
End Sub

