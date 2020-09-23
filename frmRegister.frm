VERSION 5.00
Begin VB.Form frm2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REGISTER"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   2850
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   4215
      Left            =   158
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "register"
      Height          =   255
      Left            =   878
      TabIndex        =   0
      Top             =   4680
      Width           =   975
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'this little function made it easier for me to load new levels.
'To edit a level, all u have to do is move the squares and the
'food around to your liking (make sure they line up in the grid!) and
'press the register button when you run it.  it will return a bunch of
'data.  just put that data in a new level loading process and it will
'create your level.
For i = 0 To 25
Text1.Text = Text1.Text & "Shape1(" & i & ").left = " & frmMain.Shape1(i).Left & vbCrLf
Text1.Text = Text1.Text & "Shape1(" & i & ").top = " & frmMain.Shape1(i).Top & vbCrLf
Next i
For j = 0 To 9
Text1.Text = Text1.Text & "Food(" & j & ").left = " & frmMain.Food(j).Left & vbCrLf
Text1.Text = Text1.Text & "Food(" & j & ").top = " & frmMain.Food(j).Top & vbCrLf
Next j
Text1.Text = Text1.Text & "Ball.left = " & frmMain.ball.Left & vbCrLf & "Ball.top = " & frmMain.ball.Top
End Sub

