VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Basic Shoter v 2.0           "
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7725
   Icon            =   "BasicShoter2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox BestScore1 
      Height          =   195
      Left            =   7440
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   4
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   3
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Best Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   2
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shoter Game"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   4125
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd1_Click()
Form1.Show
Unload Me
     End Sub

Private Sub cmd2_Click()
MsgBox ("Best Score = " & BestScore1.Text)
     End Sub

Private Sub cmd3_Click()
     MsgBox ("Basic Shoter Game Tutorial v2" & Chr(13) & Chr(13) & "Created by PowerOfAnubis = NironSoft" & Chr(13) & Chr(13) & "Copytighted 2005")
End Sub

Private Sub cmd4_Click()
End
End Sub

Private Sub Form_Load()
Dim BestScore As String
Open App.Path & "\Score.bestscorefile" For Binary Access Read As #1
BestScore = Space$(LOF(1))
Input #1, BestScore
Close #1
BestScore1.Text = BestScore
End Sub

Private Sub Timer1_Timer()
Static I         As Long
Const strCaption As String = "Basic Shoter v 2.0           "
I = I + 1
If I > Len(strCaption) Then
I = 0
End If
Form2.Caption = Left(strCaption, I)
End Sub

Private Sub Timer2_Timer()
Label1.Left = Label1.Left - 15 '
If Label1.Left < 0 Then ' If label1's left position is smaller than 0 then
Label1.Left = Label1.Left + 10
Timer3.Enabled = True 'Timer3 = true
Timer2.Enabled = False 'Timer2 = false
End If
End Sub

Private Sub Timer3_Timer()
Label1.Left = Label1.Left + 15
If Label1.Left > Form2.Width - Label1.Width Then ' If label1's left position is bigger than label1 width
Label1.Left = Label1.Left - 10
Timer2.Enabled = True 'Timer2 = True
Timer3.Enabled = False 'Timer3 = False
End If
End Sub
