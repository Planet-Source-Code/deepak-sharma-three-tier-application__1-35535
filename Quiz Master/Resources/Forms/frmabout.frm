VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form About 
   BackColor       =   &H0000C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Instructions"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3015
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   7335
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start Quiz"
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
         Left            =   2880
         TabIndex        =   4
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A )     All The Questions Containing Equal Marks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   3900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B )     You Can See The Correct Answer After The End Of Test."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   5010
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C )     You Have To Select At Least One Option In Each Questions."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   5340
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D )     You Cannot Go To Previous Questions So Be Careful."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   4860
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3015
      Left            =   360
      TabIndex        =   9
      Top             =   960
      Width           =   7335
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   2775
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   7095
         ExtentX         =   12515
         ExtentY         =   4895
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "res://C:\WINNT\system32\shdoclc.dll/offcancl.htm#http:///"
      End
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   7080
      Picture         =   "frmabout.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Label BACK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<< BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6960
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   7800
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Know About This Application."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   2
      Top             =   4320
      Width           =   3360
   End
   Begin VB.Label Link 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   7680
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Read Before Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2460
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BACK_Click()
  Frame1.ZOrder
  BACK.Visible = False
End Sub

Private Sub cmdStart_Click()
    Unload Me
    Test.Show
End Sub

Private Sub Form_Load()

Frame1.ZOrder
Me.WebBrowser1.Navigate "About:blank"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Link.FontUnderline = False
End Sub

Private Sub Link_Click()
  Frame2.ZOrder
  BACK.Visible = True
  
  Me.WebBrowser1.Navigate App.Path & "\Tutorial\index.html"
  
End Sub

Private Sub Link_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Link.FontUnderline = True
End Sub

