<%@ Language=VBScript %>
<%

if Request.Form("Start")="Start Quiz" then
 
 set quiz=server.CreateObject("QuizMaster.Quiz")
 set session("quiz")=quiz 
 quiz.Question 
 Response.Redirect "Quiz Master.asp"
 
end if
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<form method=post id=form1 name=form1>
<br><br><br><br><br><br><br><br><center>
<input type=submit Name=Start value="Start Quiz">
</center>


</form>
</BODY>
</HTML>
