VERSION 5.00
Begin VB.Form frmBounce 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animated Mouse's Tail"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9540
   Icon            =   "frmBounce.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   343
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   636
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Overlay 
      Height          =   975
      Left            =   4920
      Picture         =   "frmBounce.frx":1782
      ScaleHeight     =   915
      ScaleWidth      =   1155
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   7
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   6
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   5
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   4
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   3
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   2
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   1
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox img 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   1680
      Top             =   3240
   End
   Begin VB.Image ImgBall 
      Height          =   165
      Index           =   0
      Left            =   2040
      Picture         =   "frmBounce.frx":1950
      Top             =   480
      Width           =   165
   End
End
Attribute VB_Name = "frmBounce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Shape As New clsShaped
Private Sub Command1_Click()
frmAbout.Show vbModal
End Sub
Private Sub Form_DblClick()
Dim i As Integer
For i = ImgBall.LBound + 1 To ImgBall.UBound
Unload ImgBall(i)
Next i
Unload Me
End
End Sub
Private Sub Form_Load()
Dim i As Integer
For i = ImgBall.UBound + 1 To 7
Load ImgBall(i)
ImgBall(i).Visible = True
ImgBall(i).Top = ImgBall(i - 1).Top + 11
Next i
ImgBall(0).Visible = False
Call InitVal
Call InitBall
SavePicture Overlay, "b.bmp"
Timer1.Interval = 20
Timer1.Enabled = True
Me.Height = Screen.Height
Me.Width = Screen.Width
b1.Show
b2.Show
b3.Show
b4.Show
b5.Show
b6.Show
b7.Show
b1.Height = 16 * 15
b1.Width = 16 * 15
b2.Height = 16 * 15
b2.Width = 16 * 15
b3.Height = 16 * 15
b3.Width = 16 * 15
b4.Height = 16 * 15
b4.Width = 16 * 15
b5.Height = 16 * 15
b5.Width = 16 * 15
b6.Height = 16 * 15
b6.Width = 16 * 15
b7.Height = 16 * 15
b7.Width = 16 * 15
b1.Picture = ImgBall(0)
b2.Picture = ImgBall(0)
b3.Picture = ImgBall(0)
b4.Picture = ImgBall(0)
b5.Picture = ImgBall(0)
b6.Picture = ImgBall(0)
b7.Picture = ImgBall(0)
Shape.Window b1.hwnd, LoadPicture("b.bmp"), vbMagenta
Shape.Window b2.hwnd, LoadPicture("b.bmp"), vbMagenta
Shape.Window b3.hwnd, LoadPicture("b.bmp"), vbMagenta
Shape.Window b4.hwnd, LoadPicture("b.bmp"), vbMagenta
Shape.Window b5.hwnd, LoadPicture("b.bmp"), vbMagenta
Shape.Window b6.hwnd, LoadPicture("b.bmp"), vbMagenta
Shape.Window b7.hwnd, LoadPicture("b.bmp"), vbMagenta
SetWindowPos b1.hwnd, -1, 0, 0, 0, 0, 3
SetWindowPos b2.hwnd, -1, 0, 0, 0, 0, 3
SetWindowPos b3.hwnd, -1, 0, 0, 0, 0, 3
SetWindowPos b4.hwnd, -1, 0, 0, 0, 0, 3
SetWindowPos b5.hwnd, -1, 0, 0, 0, 0, 3
SetWindowPos b6.hwnd, -1, 0, 0, 0, 0, 3
SetWindowPos b7.hwnd, -1, 0, 0, 0, 0, 3
End Sub
Private Sub Timer1_Timer()
Dim Pos As POINTAPI
GetCursorPos Pos
MoveHandler Pos.x, Pos.y
Animate
b1.Left = ImgBall(1).Left * 15
b1.Top = ImgBall(1).Top * 15
b2.Left = ImgBall(2).Left * 15
b2.Top = ImgBall(2).Top * 15
b3.Left = ImgBall(3).Left * 15
b3.Top = ImgBall(3).Top * 15
b4.Left = ImgBall(4).Left * 15
b4.Top = ImgBall(4).Top * 15
b5.Left = ImgBall(5).Left * 15
b5.Top = ImgBall(5).Top * 15
b6.Left = ImgBall(6).Left * 15
b6.Top = ImgBall(6).Top * 15
b7.Left = ImgBall(7).Left * 15
b7.Top = ImgBall(7).Top * 15
End Sub
