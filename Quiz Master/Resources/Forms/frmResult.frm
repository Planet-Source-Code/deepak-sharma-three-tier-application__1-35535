VERSION 5.00
Begin VB.Form Result 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Result"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcorrect 
      Caption         =   "Click To See The Correct Answers"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   6000
      Width           =   6135
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   120
      X2              =   7680
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Some Facts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   360
      TabIndex        =   22
      Top             =   3120
      Width           =   1245
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks   :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   21
      Top             =   2400
      Width           =   1185
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Score         :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   20
      Top             =   1800
      Width           =   1230
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comments :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   19
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   120
      X2              =   7680
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label wrong 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   195
      Left            =   6480
      TabIndex        =   18
      Top             =   5280
      Width           =   105
   End
   Begin VB.Label correct 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   195
      Left            =   6480
      TabIndex        =   17
      Top             =   4800
      Width           =   105
   End
   Begin VB.Label totalquest 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   195
      Left            =   6480
      TabIndex        =   16
      Top             =   4320
      Width           =   105
   End
   Begin VB.Label avg 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   195
      Left            =   3240
      TabIndex        =   15
      Top             =   5280
      Width           =   105
   End
   Begin VB.Label min 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   195
      Left            =   3240
      TabIndex        =   14
      Top             =   4800
      Width           =   105
   End
   Begin VB.Label max 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   195
      Left            =   3240
      TabIndex        =   13
      Top             =   4320
      Width           =   105
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   6495
      Left            =   120
      Top             =   120
      Width           =   7575
   End
   Begin VB.Line Line6 
      X1              =   3960
      X2              =   6960
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line5 
      X1              =   3960
      X2              =   6960
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line4 
      X1              =   6240
      X2              =   6240
      Y1              =   4200
      Y2              =   5640
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About Yourself"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   4800
      TabIndex        =   12
      Top             =   3840
      Width           =   1260
   End
   Begin VB.Shape Shape2 
      Height          =   1455
      Left            =   3960
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Questions :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   4080
      TabIndex        =   11
      Top             =   4320
      Width           =   1500
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Wrong Answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   4080
      TabIndex        =   10
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Correct Answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   4080
      TabIndex        =   9
      Top             =   4800
      Width           =   1860
   End
   Begin VB.Line Line3 
      X1              =   2880
      X2              =   2880
      Y1              =   4200
      Y2              =   5640
   End
   Begin VB.Line Line2 
      X1              =   840
      X2              =   3720
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Average Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   960
      TabIndex        =   7
      Top             =   5280
      Width           =   1365
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   960
      TabIndex        =   6
      Top             =   4800
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum Score "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   960
      TabIndex        =   5
      Top             =   4320
      Width           =   1485
   End
   Begin VB.Line Line1 
      X1              =   840
      X2              =   3720
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About This Quiz:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1680
      TabIndex        =   4
      Top             =   3840
      Width           =   1425
   End
   Begin VB.Shape Shape1 
      Height          =   1455
      Left            =   840
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label lblmarks 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   1800
      TabIndex        =   3
      Top             =   1800
      Width           =   105
   End
   Begin VB.Label lblremarks 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   1800
      TabIndex        =   2
      Top             =   2400
      Width           =   105
   End
   Begin VB.Label lblcomment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   1770
      TabIndex        =   1
      Top             =   1200
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Result Status"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   3210
   End
End
Attribute VB_Name = "Result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcorrect_Click()
   
    Test.Show
    Test.CorrectAnswer
    Test.cmdcorrect.ZOrder
    Test.cmdcorrect.Default = True
    Test.Frame2.Visible = True
    Test.Frame3.Visible = True
    Me.Hide
    
End Sub

Private Sub Form_Load()

  Me.lblmarks = Test.Quiz.GetResult(Marks) & " Out Of 100"
  Me.lblcomment = Test.Quiz.GetResult(Comments)
  Me.lblremarks = Test.Quiz.GetResult(Remarks)
  
  Me.max = Test.Quiz.Facts(Maximum)
  Me.min = Test.Quiz.Facts(Minimum)
  Me.avg = Test.Quiz.Facts(Average)
  
  Me.totalquest = Test.Quiz.Facts(TotalQuestion)
  Me.wrong = Test.Quiz.Facts(WrongAnswers)
  Me.correct = Test.Quiz.Facts(CorrectAnswers)
 
End Sub
