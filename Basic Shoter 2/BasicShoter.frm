VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Basic Shoter"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9105
   Icon            =   "BasicShoter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "BasicShoter.frx":0ECA
   MousePointer    =   2  'Cross
   ScaleHeight     =   6570
   ScaleWidth      =   9105
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   195
      Left            =   8760
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton PlayAgain 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox LeftPosition 
      Height          =   195
      Left            =   8865
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Timer Fall 
      Interval        =   10
      Left            =   1920
      Top             =   120
   End
   Begin VB.Label FinalScore 
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   840
      Left            =   2640
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   4410
   End
   Begin VB.Label GameOver 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Game Over"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   840
      Left            =   2625
      TabIndex        =   3
      Top             =   2865
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Image Ball 
      Height          =   720
      Left            =   4200
      Picture         =   "BasicShoter.frx":1184
      Top             =   120
      Width           =   720
   End
   Begin VB.Label SCORE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   975
      TabIndex        =   1
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Score :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Copyright 2005
'By PowerOfAnubis = NironSoft
'------------------------------\
' If you like this tutorial please Vote
'------------------------------/
Dim NoOnBorder As Integer
Dim BallFallSpeed As Integer
Private Sub Ball_Click()
Call RandomTopFall ' Call the sub "BallFallSpeed"
SCORE.Caption = SCORE.Caption + 20 'Every time you click on the ball you win 20 points
BallFallSpeed = BallFallSpeed + 1 'Every time you click on the ball the fall speed of the ball increases to make the game harder
End Sub

Private Sub Fall_Timer()
Ball.Top = Ball.Top + BallFallSpeed 'Ball fall code
If Ball.Top > Form1.Height Then
If SCORE.Caption > Val(Text2) Then
Open App.Path & "\Score.bestscorefile" For Binary Access Write As #2
Put #2, , SCORE.Caption
Close #2
MsgBox ("NOW THE BEST SCORE IS " & SCORE.Caption)
End If
Fall.Enabled = False
GameOver.Visible = True
PlayAgain.Visible = True
FinalScore.Visible = True
FinalScore.Caption = "Score :" & " " & SCORE.Caption 'Final Socre
End If
End Sub

Private Sub Form_Load()
Text2.Text = Form2.BestScore1.Text
BallFallSpeed = 20 'Ball fall speed = 20
NoOnBorder = Form1.Width - 1000 'With this code the ball will never be on the border of the form
Randomize 'Make random
 Value = Int(NoOnBorder * Rnd) 'Choose a number between "NoOnBorder"
  LeftPosition.Text = Str$(Value)
   Ball.Left = LeftPosition.Text 'Ball left position = LeftPosition.text
End Sub
Sub RandomTopFall()
Ball.Top = 0 - Ball.Height 'Ball is on top of the form
Randomize ' Make Random
 Value = Int(NoOnBorder * Rnd) 'Choose a number between "NoOnBorder"
  LeftPosition.Text = Str$(Value)
   Ball.Left = LeftPosition.Text 'Ball left position = LeftPosition.text
End Sub

Private Sub PlayAgain_Click()

Form2.Show
Unload Me
End Sub
