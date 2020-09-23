VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000004&
   Caption         =   "LaserDraw 2.0"
   ClientHeight    =   4125
   ClientLeft      =   165
   ClientTop       =   780
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmProps 
      BackColor       =   &H80000004&
      Caption         =   "Props"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2175
      Begin VB.ListBox lstOpts 
         Height          =   1185
         Index           =   0
         ItemData        =   "frmMain.frx":0000
         Left            =   240
         List            =   "frmMain.frx":0007
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton cmdAction 
         BackColor       =   &H80000004&
         Caption         =   "Laser Draw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3480
         Width           =   1695
      End
      Begin VB.ListBox lstChoices 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   0
         ItemData        =   "frmMain.frx":0015
         Left            =   240
         List            =   "frmMain.frx":0025
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Laser Source"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Timer tmrRun 
      Interval        =   10
      Left            =   1560
      Top             =   4080
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   2400
      TabIndex        =   9
      Top             =   0
      Width           =   6255
      Begin VB.PictureBox board 
         BackColor       =   &H00000000&
         Height          =   3615
         Left            =   120
         ScaleHeight     =   237
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   397
         TabIndex        =   10
         Top             =   240
         Width           =   6015
         Begin VB.Shape cir 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   5880
            Shape           =   3  'Circle
            Top             =   3480
            Visible         =   0   'False
            Width           =   135
         End
      End
   End
   Begin MSComDlg.CommonDialog cmn 
      Left            =   6960
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   -2520
      Picture         =   "frmMain.frx":0075
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Frame frmProps 
      BackColor       =   &H80000004&
      Caption         =   "Props"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2175
      Begin VB.ListBox lstOpts 
         Height          =   960
         Index           =   1
         ItemData        =   "frmMain.frx":8E0B
         Left            =   240
         List            =   "frmMain.frx":8E12
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ListBox lstChoices 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   1
         ItemData        =   "frmMain.frx":8E20
         Left            =   240
         List            =   "frmMain.frx":8E2A
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdAction 
         BackColor       =   &H80000004&
         Caption         =   "Sweep"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sweep Direction"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuChangePic 
         Caption         =   "Change Picture"
      End
      Begin VB.Menu mnuView 
         Caption         =   "View Selected Picture"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop Laser"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuDraw 
      Caption         =   "Draw"
      Begin VB.Menu mnuLaser 
         Caption         =   "Laser Draw"
      End
      Begin VB.Menu mnuSweep 
         Caption         =   "Sweep"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FWidth, FHeight
Dim SRCX As Integer, SRCY As Integer

Private Sub board_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SRCX = X
SRCY = Y

cir.Left = SRCX - (cir.Width \ 2)
cir.Top = SRCY - (cir.Height \ 2)
End Sub

Private Sub cmdAction_Click(Index As Integer)
Running = False
DoEvents
Running = True
mnuStop.Enabled = True
board.Cls

Select Case Index
Case 0
LaserDraw pic, board, lstChoices(0).ItemData(lstChoices(I).ListIndex), lstOpts(0).Selected(0), SRCX, SRCY
 
Case 1
SweepDraw pic, board, lstChoices(1).ItemData(lstChoices(1).ListIndex), lstOpts(1).Selected(0)
End Select
End Sub

Private Sub Form_Load()
Dim I As Long, N() As String

Me.Icon = LoadResPicture(101, 1)

LoadRes

For I = 0 To frmProps.UBound
lstChoices(I).ListIndex = 0
Next I

N() = Split(LoadResString(102), ",")
FWidth = Val(N(0))
FHeight = Val(N(1))

mnuLaser_Click
SRCX = board.ScaleWidth
SRCY = board.ScaleHeight
End Sub

Function SelectProps(Index As Long)
Dim I As Long

For I = 0 To frmProps.UBound
frmProps(I).Visible = True

If I <> Index Then frmProps(I).Visible = False
Next I
End Function

Function LoadRes()
pic.Picture = LoadResPicture(101, 0)
End Function

Private Sub Form_Resize()
If Me.WindowState = 0 Then
Me.Width = FWidth
Me.Height = FHeight
End If
End Sub

Private Sub mnuBlinds_Click()
SelectProps 2
End Sub

Private Sub mnuChangePic_Click()
cmn.DialogTitle = "Change Picture"
cmn.Filter = LoadResString(101)
cmn.ShowOpen

If Len(cmn.FileName) <= 0 Then Exit Sub

pic.Picture = LoadPicture(cmn.FileName)
End Sub

Private Sub mnuFade_Click()
SelectProps 3
End Sub

Private Sub mnuLaser_Click()
cir.Visible = True
SelectProps 0
End Sub

Private Sub mnuStop_Click()
Running = False
End Sub

Private Sub mnuSweep_Click()
cir.Visible = False
SelectProps 1
End Sub

Private Sub mnuView_Click()
Load frmVeiw
frmVeiw.board.Picture = pic.Picture
frmVeiw.Visible = True
End Sub

Private Sub tmrRun_Timer()
Me.Caption = IIf(Running, "Laser Draw 2.0 - Drawing... " & PercentDone & "%", "Laser Draw 2.0 - Done")
End Sub
