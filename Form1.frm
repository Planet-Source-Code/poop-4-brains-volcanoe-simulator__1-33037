VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   1890
   ClientTop       =   1350
   ClientWidth     =   7950
   DrawWidth       =   3
   FillColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   367
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   530
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Flow"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   1350
   End
   Begin VB.PictureBox fountainm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   2640
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   1
      Top             =   5760
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.PictureBox fountains 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   375
      Picture         =   "Form1.frx":12602
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   0
      Top             =   5865
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Timer Timer1 
      Left            =   3240
      Top             =   4845
   End
   Begin VB.PictureBox board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawWidth       =   10
      Height          =   4545
      Left            =   120
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   250
      TabIndex        =   2
      Top             =   360
      Width           =   3750
   End
   Begin VB.PictureBox buffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   10
      Height          =   4455
      Left            =   4080
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   250
      TabIndex        =   4
      Top             =   360
      Width           =   3750
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mountain Terrain:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your View:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim running As Long
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Select Case Command2.Caption
Case "Stop Flow"
running = False
Command2.Caption = "Continue Flow"
Case "Continue Flow"
running = True
Command2.Caption = "Stop Flow"
End Select
End Sub

Private Sub Form_Load()
running = True
Randomize
board.picture = Nothing
Me.Show
Dim c As Long, i As Long
Do
If GetTickCount > 10 Then

board.Cls
buffer.Cls
BitBlt buffer.hDC, buffer.ScaleWidth \ 2 - (250 \ 2), buffer.ScaleHeight - 100, 250, 100, fountainm.hDC, 0, 0, SRCAND
BitBlt board.hDC, board.ScaleWidth \ 2 - (250 \ 2), board.ScaleHeight - 100, 250, 100, fountainm.hDC, 0, 0, SRCAND
BitBlt board.hDC, board.ScaleWidth \ 2 - (250 \ 2), board.ScaleHeight - 100, 250, 100, fountains.hDC, 0, 0, SRCINVERT

For i = 0 To 1000
If P(i).Act = True Then
board.PSet (P(i).x, P(i).y), P(i).Color
If P(i).Hit >= 15 Then buffer.PSet (P(i).x, P(i).y), vbBlack
End If
Next i

If running = True Then
For i = 0 To 5
MakeDroplet board
Next i
End If

MoveDrops board
End If
DoEvents
Loop
End Sub

Private Sub Timer1_Timer()
BitBlt board.hDC, board.ScaleWidth \ 2 - (250 \ 2), board.ScaleHeight - 100, 250, 100, fountainm.hDC, 0, 0, SRCAND
BitBlt board.hDC, board.ScaleWidth \ 2 - (250 \ 2), board.ScaleHeight - 100, 250, 100, fountains.hDC, 0, 0, SRCINVERT
End Sub
