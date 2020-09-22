<%@ Language=VBScript %>
<!--#include file=Const.inc-->
<%Response.Buffer=true
Dim QuestionNo,remaning,total,Question,msg

	'set quiz=server.CreateObject("QuizMaster.Quiz")
		
Set Quiz=session("quiz")
		
If Quiz.CurrentQuestion<10 then		
	
   Caption="Next"		
		 
Else	   
	  
   Caption="Finish"
	   
End if
	
	
Select case Request.Form("nav")
	
Case "Next" 
	    
	    Call Display_Questions
	
Case "Finish"

		Quiz.UserChoice Request.Form("choice")
		
		Response.Redirect "result.asp"
	
End Select
		


'******** Procedure/Method/Functions ********

Public Sub Store()

	session("Restore_question")=Question
	session("Restore_questionno")=questionno
	session("Restore_remaning")=remaning
	session("Restore_total")=total

End Sub	

Public Sub DisplayQuestion()

	QuestionNo=quiz.CurrentQuestion 
	Remaning=quiz.RemaningQuestions 
	Total=quiz.TotalQuestions 		
	Question=quiz.Question

End Sub	

Public Sub Restore()
	
	Question=Session("Restore_question")
	Questionno=Session("Restore_questionno")
	Remaning=Session("Restore_remaning")
	Total=Session("Restore_total")	

End Sub

Public Sub Display_Questions()

	If Request.Form("choice")="" then
		  
	    Msg="Select atleast one option from them"		  		  
	    Call Restore
		  
	Else
		
       Quiz.UserChoice Request.Form("choice")					
	   Call DisplayQuestion				
	   Call Store
		
	End If

End Sub

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<form method=post action="quiz master.asp" name=myform>
<br><br>

<table border=1 width='80%' align=center>  

<tr height=30><td>
	
	<%If Question="" then	
 
		call DisplayQuestion
		call Store				
		
	  End if
	%>	
  
	Question No :&nbsp;<%=QuestionNo%>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	Remaning Question :&nbsp;<%=Remaning%>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	Total Question :&nbsp;<%=Total%>
 
</tr></td>		
 
<tr height=40><td>	
	    
   <%=Question%>
		 
</tr></td>		

<tr><td>
	
   &nbsp;<br>
   &nbsp;&nbsp;&nbsp;<input type=radio name=choice value="1" onclick="javascript:myform.message.value=''"><%=quiz.Options(Options1)%><br><br>
   &nbsp;&nbsp;&nbsp;<input type=radio name=choice value="2" onclick="javascript:myform.message.value=''"><%=quiz.Options(Options2)%><br><br>
   &nbsp;&nbsp;&nbsp;<input type=radio name=choice value="3" onclick="javascript:myform.message.value=''"><%=quiz.Options(Options3)%><br><br>
   &nbsp;&nbsp;&nbsp;<input type=radio name=choice value="4" onclick="javascript:myform.message.value=''"><%=quiz.Options(Options4)%><br>&nbsp				

</tr></td>

<tr><td>

   <center><input type=text name=message size=40 value="<%=msg%>" style="position:relative;border:0;text-align:center;color:red;font-weight:bold"></center>

</tr></td>
  
<tr height=40><td>

   <center><input type=submit Name=Nav value=<%=Caption%> style="position:relative;width:50"></center>            
  
</tr></td>  
  
</table>
</form>
</BODY>
</HTML>
<%




%>