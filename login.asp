<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
</head>

<body>
<center><table BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH="80%" >
<tr BGCOLOR="#FFCC00">
<td></td>

<td HEIGHT="25">
<center>
          <font color="#000000"><b>Welcome to Barcodeticket (v.1.1) Admin 
          Page.</b></font> 
        </center>
</td>
</tr>
</table></center>

<center><table BORDER=0 CELLSPACING=0 CELLPADDING=10 WIDTH="80%" >
<caption><center></center></caption>

<tr>
<td ALIGN=LEFT VALIGN=TOP WIDTH="50%">
<hr>
        <div align="left">Please <a href="mailto:chris@barcodeticket.com"> feel 
          free to email us for customisation.</a> Barcode ticket comes with administrative 
          requirements such as checking for verified tickets, your sales report, 
          updates posted on the fly through a web interface and much more. </div>
		 <br><br>
        <font size="2"></font></a> </td>

<td ALIGN=LEFT VALIGN=TOP WIDTH="50%">
        <table BORDER=0 CELLSPACING=0 CELLPADDING=3 width="293" >
          <tr BGCOLOR="#CCCCFF">
            <td ALIGN=LEFT><b><font size=+1>Login to administer</font></b></td>
</tr>

<tr ALIGN=LEFT BGCOLOR="#EEEEEE">
<td>
			  <FORM name=form action=main.asp method=post>
              <table>
                <tr> 
                  <td COLSPAN="2"><b>Please sign in.</b></td>
                </tr>
                <tr VALIGN=CENTER> 
                    <td><b><font size=-1>Your Login:</font></b></td>
                  <td>
                    <input TYPE="TEXT" NAME="Login_Name" SIZE="20">
                  </td>
                </tr>
                <tr VALIGN=CENTER> 
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr VALIGN=CENTER> 
                  <td><b><font size=-1>Password:&nbsp;</font></b></td>
                  <td>
                    <input type=password size=20 name=password>
                  </td>
                </tr>
                <tr VALIGN=CENTER> 
                  <td></td>
                  <td> <br>
                    <input type=submit value="Sign In">
                  </td>
                </tr>
                <tr> 
                  <td></td>
                  <td></td>
                </tr>
              </table></form>
</td>
</tr>
</table>
</td>
</tr>
</table></center>


</body>
</html>



<%
function ReadFile(fname)
	FileName = server.mappath(fname)
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set a = fs.OpenTextFile(FileName, 1)
	ReadFile = a.readline()
	a.close
	set a = nothing
	set fs = nothing
end function
%>


