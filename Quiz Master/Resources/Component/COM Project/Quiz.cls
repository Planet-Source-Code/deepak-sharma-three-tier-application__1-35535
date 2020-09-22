VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Quiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Type Quest
   Question_No As Integer
   Correct_Option As Integer
   User_Option As Integer
End Type

Enum Options
  Options1 = 1
  Options2 = 2
  Options3 = 3
  Options4 = 4
End Enum

Enum Result
  Marks = 1
  Comments = 2
  Remarks = 3
End Enum

Enum Correct_Answers
  Status = 1
  UserOption = 2
  CorrectOption = 3
  Increment = 4
End Enum

Enum Fact
  Maximum = 1
  Minimum = 2
  Average = 3
  TotalQuestion = 4
  CorrectAnswers = 5
  WrongAnswers = 6
End Enum

Private Questions() As Quest

Private Sub Class_Initialize()

   Path = Replace(App.Path, "Component", "") & "\Database\Quiz.mdb"
         
   If cn.State = 1 Then cn.Close
   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Path & ";Persist Security Info=False"
   cn.CursorLocation = adUseClient
   
   rs.Open "select * from questions", cn, adOpenDynamic, adLockOptimistic
   
   ReDim Questions(rs.RecordCount - 1)
   
   I = 0: num = 0: k = 0: B = 0: D = 0: Correct_Answer = 0
      
End Sub

Public Function Question() As String

   GenerateQuestion
      
   Question = rs.Fields("Question")
 
End Function

Public Function Options(Opt As Options) As String
  
    Select Case Opt
    
    Case 1
  
          Options = "A) " & rs.Fields("option1")
       
    Case 2
  
          Options = "B) " & rs.Fields("option2")
    
    Case 3
  
          Options = "C) " & rs.Fields("option3")
  
    Case 4
  
          Options = "D) " & rs.Fields("option4")
  
    End Select
  
End Function

Private Function GenerateQuestion()

    RandomQuestionNo
 
    If num < UBound(Questions()) Then
    
        rs.AbsolutePosition = Questions(num).Question_No
        num = num + 1
        
    End If

End Function

Public Function TotalQuestions() As Integer
   
   TotalQuestions = rs.RecordCount
   
End Function

Private Function RandomQuestionNo()
  
   Randomize

   If I < rs.RecordCount Then

Jump:

        Quest_No = (Rnd * (TotalQuestions - 1)) + 1
  
        For Incre = 0 To UBound(Questions())
       
           If Quest_No = Questions(Incre).Question_No Then
                
                 GoTo Jump
            
            Else
           
                 AddQuestion = True
        
            End If
       
        Next

        
        If AddQuestion = True Then
        
          Questions(I).Question_No = Quest_No
          
          'Store the correct option of particular question
          rs.AbsolutePosition = Quest_No
          Questions(I).Correct_Option = Val(rs.Fields("correct_option"))
          
        End If
     
        I = I + 1
        D = D + 1
        
   End If
  
End Function

Public Function UserChoice(choices As String)

   If k < UBound(Questions()) + 1 Then
   
     Questions(k).User_Option = Val(choices)
     k = k + 1
     
   End If
   
End Function

Public Function CurrentQuestion() As Integer
   
   CurrentQuestion = D
   
End Function

Public Function RemaningQuestions() As Integer
   
   RemaningQuestions = TotalQuestions - D
   
End Function

Public Function GetResult(Disp As Result) As String
    
    If Total <= 40 Then
      
      Comment = "Sorry You Didn't Clear This Quiz."
      Remark = "Bad Luck This Time,Better Luck Next Time"
      
    ElseIf Total <= 60 Then
    
      Comment = "Congratulation You Have Successfully Completed This Quiz."
      Remark = "Good Knowldge"
      
    ElseIf Total <= 90 Then
    
      Comment = "Congratulation You Have Successfully Completed This Quiz."
      Remark = "Excellent Knowldge"
    
    ElseIf Total >= 100 Then
    
      Comment = "Congratulation You Have Successfully Completed This Quiz."
      Remark = "Unbelievable Knowldge. You Are The Master Of Knowldge"
      
    End If
    
    Select Case Disp
    
    Case 1 'Marks
    
        For J = 0 To UBound(Questions())
        
            If Questions(J).User_Option = Questions(J).Correct_Option Then
          
                  Total = Total + 10
              
                  Correct_Answer = Correct_Answer + 1
              
            Else
             
                 Total = Total + 0
          
            End If
    
        Next
        
        'Add the score to database
        Add_Score
       
        GetResult = Total
    
    Case 2 'Comments
       
       GetResult = Comment
    
    Case 3 'Remarks
       
       GetResult = Remark
       
    End Select

    num = 0
    D = 1
    
End Function

Private Function Add_Score()
    
    If Find.State = 1 Then Find.Close
    Find.Open "score", cn, adOpenDynamic, adLockOptimistic
        
    If Find.RecordCount < 2 Then
    
        If Find.RecordCount = 0 Then
    
            cn.Execute "insert into score values(" & Total & ")"
        
        ElseIf Total <> Find.Fields(0) Then
        
            cn.Execute "insert into score values(" & Total & ")"
    
        End If
    
    Else
    
        If Find.State = 1 Then Find.Close
        Find.Open "select Max(scores) from score", cn, adOpenDynamic, adLockOptimistic
    
        If Total > Find.Fields(0) Then
      
            cn.Execute "Update score set scores=" & Total & " where scores=" & Find.Fields(0)
      
        Else
    
            If Find.State = 1 Then Find.Close
            Find.Open "select Min(scores) from score", cn, adOpenDynamic, adLockOptimistic
    
            If Total < Find.Fields(0) Then
      
                cn.Execute "Update score set scores=" & Total & " where scores=" & Find.Fields(0)
      
            End If
            
        End If
    
    End If
      
End Function

Public Function Facts(fac As Fact) As String

    If Find.State = 1 Then Find.Close
    
    Select Case fac
    
    Case 1  'Max
    
        Find.Open "select Max(scores) from score", cn, adOpenDynamic, adLockOptimistic
        Facts = Find.Fields(0)
    
    Case 2  'Min
    
        Find.Open "select Min(scores) from score", cn, adOpenDynamic, adLockOptimistic
        Facts = Find.Fields(0)
    
    Case 3  'Avg
    
        Find.Open "select Avg(scores) from score", cn, adOpenDynamic, adLockOptimistic
        Facts = Find.Fields(0)
    
    Case 4  'Total
        
        Facts = TotalQuestions
    
    Case 5  'Correct
       
        Facts = Correct_Answer
    
    Case 6  'Wrong
    
       Facts = TotalQuestions - Correct_Answer
    
    End Select
    

End Function

Public Function ShowCorrectAnswer() As Boolean

    If CurrentQuestion = TotalQuestions Then

        ShowCorrectAnswer = True
    
    Else
    
        ShowCorrectAnswer = False
  
    End If

End Function

Public Function ShowQuestionStatus(corr As Correct_Answers) As String
 
    Select Case corr

    Case 1    'Status

           ShowQuestionStatus = IIf(Questions(B).Correct_Option <> Questions(B).User_Option, _
                              "Wrong Answer", "Correct Answer")
                               

    Case 2    'UserChoice

          ShowQuestionStatus = Choose(Questions(B).User_Option, "A", "B", "C", "D")
 

    Case 3    'CorrectOption

         ShowQuestionStatus = Choose(Questions(B).Correct_Option, "A", "B", "C", "D")
     
  
    Case 4    'Increment
  
         If B < UBound(Questions()) Then B = B + 1
     
         If D < TotalQuestions Then D = D + 1

    End Select
 
End Function

Public Function Arr() As String
    Dim z, m
    
    For z = 0 To UBound(Questions())
        
          m = m & "Question : " & Questions(z).Question_No & " User Option : " & Questions(z).User_Option & " correct option : " & Questions(z).Correct_Option & vbCrLf
          
    Next
    
    Arr = m
    
End Function
