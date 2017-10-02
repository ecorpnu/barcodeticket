<%
if request.cookies("login") <> "true" then
response.redirect("login.asp")
end if
%>

<%
Dim TypePage, FileName, FS, A
Dim TextContent

TypePage = Request("TypePage")
Select Case TypePage
  Case "CH":  '_____________________event heading
     FileName = "ch.htm"
  Case "IL":  '_____________________ImageLink
     FileName = "IL.htm"
  Case "TI":  '_____________________ImageLink for ticket
     FileName = "TI.htm"
  Case "TI2":  '_____________________2nd ImageLink for ticket
     FileName = "TI2.htm"
  Case "CI":  '_____________________Event intro
     FileName = "CI.htm"
  Case "CD":  '_____________________Event details
     FileName = "cd.htm"
  Case "CDate":  '__________________Event date
     FileName = "cdate.htm"
  Case "CL":  '_____________________Purchase link
     FileName = "cl.htm"
  Case "CE":  '_____________________Purchase condition
     FileName = "ce.htm"
  Case "CEmail":  '_________________Email Notification
     FileName = "cemail.htm"
  Case "Low":  '_____________________Purchase condition
     FileName = "low.htm"
  Case "Bcc":  '_____________________Email bcc
     FileName = "cebcc.htm"
  Case "tpp":  '_____________________Host URL 
     FileName = "pathprint.htm"
  Case "pp":  '_____________________Paypal print path
     FileName = "pp.htm"
  Case "ppemail":  '_____________________Paypal email
     FileName = "ppemail.htm"  
  Case else: '_____________________error case, default is heading
     FileName = "default.htm"
	 TypePage = "R"
End Select

FileName = Server.Mappath("text/" & FileName)
Set FS = CreateObject("Scripting.FileSystemObject")

If Request("Submit") = "Submit" then
  TextContent = Request("TextContent")
  Set A = FS.CreateTextFile(FileName, true)
  A.writeline(TextContent)  
  A.close
End If

Set A = FS.OpenTextFile(FileName, 1, False)
%>
<html>
<head>
<title></title>

</head>

<body bgcolor="#FFFFFF">
<p class="head1">Edit Text in File:</p>
<p class="Black"><b><%= FileName %></b></p>
<form method="post" action="admin_text.asp" method="post">
  <input type="hidden" name="TypePage" value="<%= TypePage %>">
<p><a href="main.asp"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">back</font></a></p>
<p>
  <textarea name="TextContent" rows="13" cols="70">
<%
Do While Not A.AtEndOfStream
        Response.write(A.ReadLine() & chr(13) & chr(10))
Loop
A.Close
Set A = Nothing
Set FS = Nothing
%>	
</textarea>
</p>
  
 <p> 
   <input type="submit" name="Submit" value="Submit">
   <input type="reset" name="Submit2" value="Reset">
<%
If Request("Submit") = "Submit" then
%>
   &nbsp&nbsp;<b><font face="Verdana, Arial, Helvetica, sans-serif" color="#CC0000">Update data Complete ...</font></b>
<%
End If
%>
  </p>
<p><a href="main.asp"><font face="Verdana, Arial, Helvetica, sans-serif" size="1">back</font></a></p>
</form>
<p>&nbsp;</p>
</body>
</html>
