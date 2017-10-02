<!--#INCLUDE FILE="ADOVBS.INC"-->
<!--#include file ="_OpenConnection.asp"-->
<!--#INCLUDE FILE="errornumeric.asp"-->
<% 
'To make sure the whole page is process before output
response.buffer= true 
%>

<!--#include file ="1.asp"-->

<% ' error check
if TRIM( Request( "id" ) ) = "" or request("productID") ="" then
  errorForm "ppbuy.asp", "You are not authorised here."
end if

%>

<%              'check to see if ID is here
                prodID=request("productID")        
		strSQL7 = "Select imagepath FROM product WHERE product_id="&prodID
	   	set rs7=conn.execute(strSQL7)
                If rs7.eof then
                errorForm "ppbuy.asp", "You are not authorised here."      
                end if
%>



<%   'check to see if ID is in database
testid = Request( "id" )
mySQL2 = "SELECT * FROM card WHERE card_ID=" & testid
Set RS2 = conn.execute(mySQL2)
If rs2.eof then
errorForm "ppbuy.asp", "You are not authorised here."      
end if

rs2.close
set rs2 = nothing
%>


<html>
<head>
</head>
<body>

Thank you for your purchase 
<IMG SRC="images/rulecut1.gif" width="788" Height="15">


<table width="600" align="center" cellspacing="0" cellpadding="0" border="0" bgcolor="#E1DDC6">
<tr><!-- Section with main graphic and 2 top corners -->
    <td valign="top" rowspan="2" align="left" width="9"><img src="images/Dor_tlcorner.gif" width="9" height="9" border="0"></td>
    <td align="center">
<img src="images/white_w2.gif" width="9" height="9" border="0" vspace="6">
</td>
    <td valign="top" rowspan="2" align="right" width="9"><img src="images/Dor_trcorner.gif" width="9" height="9" border="0"></td>
</tr><!-- end Section with main graphic and 2 top corners for enomitron -->

<tr>

<td>


<%
jobid = Request( "id" )

' Construct SQL
mySQL = "SELECT * FROM card WHERE "
mySQL = mySQL & "card_ID=" & jobid

' Ready DB Connection


' Open Recordset
Set RS = Server.CreateObject( "ADODB.Recordset" )
rs.ActiveConnection = Conn
rs.CursorLocation = adUseClient
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
RS.Open mySQL

If not rs.eof then
 


    Response.Write " <table> "
    Response.Write " <tr> "
    Response.Write "<td>"
'	if rs("buyfor") = "adult" then
'		response.write "Adult " & rs("adultcost") & "<br>"
'	end if
'	if rs("buyfor") = "child" then
'		response.write "Child : $" & rs("childcost") & "<br>"
'	end if
    Response.Write  "Auth:  "  & RS("refnum") 
    Response.Write "<br>"
    Response.Write " Name: " & RS("card_name")
    Response.Write "<br>"
    Response.Write " Type: " & RS("type_ticket")
    Response.Write "<br>"
    
    Response.Write " Event: " & RS("Des")
    Response.Write "<br>"
    Response.Write " Date: " & RS("Date1" )
    Response.Write "<br>"
    Response.Write " Email: " & RS("emailuser")
    
    Response.write ""
    Response.Write "</td>"

rs("clientemail")=now()
rs.Update 

    Response.write " <td> "
set barcode=Server.CreateObject("CseBarcodeDll.Barcodeclass")

ParsedString =Barcode.Encode(RS("SEDID"))

         BlankorSpace = True
         For Checkvalue = 1 To Len(ParsedString)
             If (BlankorSpace) Then 
               	 ImageGIF = Mid(ParsedString, Checkvalue, 1) & "b.gif"
             Else
                ImageGIF = Mid(ParsedString, Checkvalue, 1) & "s.gif"
             End If
             Response.Write "<IMG SRC=bars/" &ImageGIF& ">"
             
             BlankorSpace = Not BlankorSpace
         Next
             Response.Write "<br>"
             Response.Write "<center>" 
             Response.Write RS("SEDID")
             Response.Write "</center>"
Response.Write "</td>"
Response.write "<td>"


%>




<td><IMG SRC=<!-- #include file="text/TI2.htm"--> height ="130" width="32"></td>
<td>
<% 
                'prodID=request("productID")        
		'strSQL7 = "Select imagepath FROM product WHERE product_id="&prodID
	   	'set rs7=conn.execute(strSQL7)
                'If not rs7.eof then 
		If len(rs7("imagepath")) > 5 Then        
                           
			pathstr=rs7("imagepath")
         
			else
        
        		pathstr=ReadFile("text/ti.htm")
        
			end if
                        'end if
			rs7.close 
                        set rs7 = nothing %>


<img src="<%= pathstr %>" width="110" height="120">
</td>
<p>
</tr>
</table>
<% else %>
<% end if %>
<tr><!-- bottom corners -->
    <td height="9" align="left"><img src="images/Dor_blcorner.gif" width="9" height="9" border="0"></td>
    <td height="9"><img src="images/blank.gif" width="1" height="1" border="0"></td>
    <td height="9" align="right"><img src="images/Dor_brcorner.gif" width="9" height="9" border="0"></td>
</tr><!-- end bottom corners -->
</table>
<table>
<td width="740" height="6" valign="top"></td>


<IMG SRC="images/rulecut1.gif" width="788" height="15">

</tr>

</TABLE>

</center>

<br>
<b>
<UL>
<li>Please print this.
</li>
</UL>
</b><p><b><center><font color=#990000><u>IMPORTANT : YOU HAVE 3 MINUTES TO PRINT THIS PAGE !</u></font></center>
</b>

<br>


<!--#include file ="_closeConnection.asp"-->





 
<p>
<p>Bring this printed copy and remember your password </b>
<p></tr></td></table>


<form name="frm" action="ppbuy.asp">

<script language="javaScript">
<!--hide
{
mytime=setTimeout("document.frm.submit()",180000);
}
//-->
</script>

<input type=hidden name=hidrecord value="0">

</form>



</body>
</html>

<%
function ReadFile(fname)
	FileName = Server.Mappath(fname)
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set a = fs.OpenTextFile(FileName, 1)
	ReadFile = a.readline()
	a.close
	set a = nothing
	set fs = nothing
end function
%>
