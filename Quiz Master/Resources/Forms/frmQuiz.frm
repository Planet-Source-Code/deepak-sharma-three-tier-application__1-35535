VERSION 5.00
Begin VB.Form Test 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "3 Tyre Application (VB - Com - Access)"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   615
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   2535
      Begin VB.Shape Shape2 
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "= Right Answer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   195
         Left            =   600
         TabIndex        =   26
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "= Wrong Answer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   195
         Left            =   600
         TabIndex        =   25
         Top             =   360
         Width           =   1410
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00008000&
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   120
         Width           =   165
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   615
      Left            =   5880
      TabIndex        =   18
      Top             =   120
      Width           =   3375
      Begin VB.Shape Shape3 
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label Stat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   960
         TabIndex        =   21
         Top             =   0
         Width           =   75
      End
      Begin VB.Label user 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   75
      End
      Begin VB.Label correct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   195
         Left            =   1680
         TabIndex        =   19
         Top             =   360
         Width           =   75
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   9135
      Begin VB.OptionButton Choice 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   0
         Left            =   240
         OLEDropMode     =   1  'Manual
         TabIndex        =   6
         Tag             =   "1"
         Top             =   1440
         Width           =   8295
      End
      Begin VB.OptionButton Choice 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Tag             =   "2"
         Top             =   2040
         Width           =   8295
      End
      Begin VB.OptionButton Choice 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Tag             =   "3"
         Top             =   2640
         Width           =   8295
      End
      Begin VB.OptionButton Choice 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Tag             =   "4"
         Top             =   3240
         Width           =   8295
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "Next"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdfinish 
         Caption         =   "Finish"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   7
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdcorrect 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   17
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Question 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   8655
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   6
         Height          =   4815
         Left            =   0
         Top             =   0
         Width           =   9135
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   9120
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   0
         X2              =   9120
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   0
         X2              =   9120
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Question No :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Question :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4920
         TabIndex        =   14
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining Question :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2280
         TabIndex        =   13
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label lbltime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   7680
         TabIndex        =   12
         Top             =   240
         Width           =   105
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   7080
         Picture         =   "frmQuiz.frx":0000
         Stretch         =   -1  'True
         Top             =   45
         Width           =   480
      End
      Begin VB.Label total 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   240
         Left            =   6480
         TabIndex        =   11
         Top             =   240
         Width           =   105
      End
      Begin VB.Label Questionno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   105
      End
      Begin VB.Label Remaning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   240
         Left            =   4320
         TabIndex        =   9
         Top             =   240
         Width           =   105
      End
      Begin VB.Label Msg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   3840
         Width           =   8460
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   0
         X2              =   9120
         Y1              =   3720
         Y2              =   3720
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   800
      Left            =   120
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   720
      Top             =   240
   End
   Begin VB.Image Image3 
      Height          =   600
      Left            =   7560
      Picture         =   "frmQuiz.frx":0442
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   840
      Picture         =   "frmQuiz.frx":0D0C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quiz Master"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   2520
   End
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------
'Created and designed by Deepak Sharma
'-------------------------------------
'Client     : VB
'Middleware : COM
'Server     : Access
'-------------------------------------
'          Vote Me Please
'-------------------------------------

Public Quiz As QuizMaster.Quiz
Dim i As Integer
Dim B As Integer
Dim num As Integer
Dim UserChoice As Boolean
Dim Select_Option As Integer

Private Sub Choice_Click(Index As Integer)

  Select_Option = Index
  Msg.Caption = ""
  
End Sub

Private Sub cmdcorrect_Click()
   
    CorrectAnswer
    
    If Me.Remaning.Caption = "0" Then
    
        Me.cmdcorrect.Caption = "Quit"
        num = num + 1
      
    End If
    
    If num = 2 Then
    
        MsgBox "Please Send Your Suggestions To deepakmailto@rediffmail.com. And Don't Forget To Vote Me.", vbInformation
        End
     
    End If
    
End Sub

Private Sub cmdfinish_Click()

    Quiz.UserChoice Choice(Select_Option).Tag
    Result.Show
    Me.Hide

End Sub

Private Sub CMDNEXT_Click()

    Display_Questions
  
End Sub

Public Sub Display_Questions()
  
    If Quiz.RemaningQuestions = 1 Then cmdfinish.ZOrder
    ShowQuestions

End Sub

Public Sub CorrectAnswer()
  
    DisplayQuestion
 
    Stat.Caption = Quiz.ShowQuestionStatus(Status)
    user.Caption = "User Option : " & Quiz.ShowQuestionStatus(UserOption)
    correct.Caption = "Correct Option : " & Quiz.ShowQuestionStatus(CorrectOption)
    
    SelectUserOption
  
    Quiz.ShowQuestionStatus (Increment)
  
End Sub

Public Sub SelectUserOption()

    If Quiz.ShowQuestionStatus(UserOption) = Quiz.ShowQuestionStatus(CorrectOption) Then

        UserChoice = True

    Else

        UserChoice = False

    End If

    Select Case Quiz.ShowQuestionStatus(UserOption)
  
    Case "A"
    
        i = 0
    
    Case "B"
  
        i = 1
  
    Case "C"
  
        i = 2
  
    Case "D"
  
        i = 3
  
    End Select
    
  
    For B = 0 To Choice.UBound
  
        If B = i Then
            Choice(B).Value = True
            Choice(B).ForeColor = IIf(UserChoice = True, &H8000&, &HFF&)
        Else
            Choice(B).ForeColor = &H404080
        End If
        
    Next
  
End Sub

Public Sub ShowQuestions()
  
    If AllOptionsDeselect = False Then
     
             Msg.Caption = "Select atleast one option from them"
     
    Else
  
        Quiz.UserChoice Choice(Select_Option).Tag
        DisplayQuestion
    
    End If

End Sub

Private Sub Form_Load()

  Set Quiz = New QuizMaster.Quiz
  
  total.Caption = Quiz.TotalQuestions

  DisplayQuestion
  
  cmdnext.ZOrder
  
  num = 0
  Frame2.Visible = False
  Frame3.Visible = False
 
End Sub

Private Sub Timer1_Timer()

  Me.lbltime.Caption = Time
    
End Sub

Public Sub ClearOptions()

    For i = 0 To Choice.UBound
    
       Choice(i).Value = False
       
    Next
   
End Sub

Public Sub DisplayQuestion()
 
  Question.Caption = Quiz.Question
  
  Choice(0).Caption = Quiz.Options(Options1)
  Choice(1).Caption = Quiz.Options(Options2)
  Choice(2).Caption = Quiz.Options(Options3)
  Choice(3).Caption = Quiz.Options(Options4)
  
  Questionno.Caption = Quiz.CurrentQuestion
  Remaning.Caption = Quiz.RemaningQuestions
  
  ClearOptions
  Msg.Caption = ""
  
End Sub

Public Function AllOptionsDeselect() As Boolean
   
    For i = 0 To Choice.UBound
    
       If Choice(i).Value = False Then
          
          AllOptionsDeselect = False
       
       Else
       
          AllOptionsDeselect = True
          Exit Function
       
       End If
       
    Next
  
End Function

Private Sub Timer2_Timer()

    If Msg.Visible = True Then
    
       Msg.Visible = False
       
    ElseIf Msg.Visible = False Then
    
       Msg.Visible = True
       
    End If
  
End Sub
