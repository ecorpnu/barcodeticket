<%@LANGUAGE ="VBScript"%>


<!--#include file ="_OpenConnection.asp"-->
<!--#INCLUDE FILE="ADOVBS.INC"-->

<%
if (request.cookies("login")="") or (Request.Cookies("login") = "false") then
	if request.form("Login_Name")="" or request.form("password")="" then
		response.redirect("login.asp")
	end if
        
	strUserID = Request("Login_Name")
	strPassword = Request("Password")
	strEmail = Request("Email")
	strSQL = "Select * FROM users WHERE Login_Name = '" & strUserID & "' AND Password = '" & strPassword & "'"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open strSQL, conn, adOpenStatic, adLockOptimistic, adCmdText
        'set rs=conn.execute(strSQL)
	If not rs.EOF Then 
		n = now
		Response.Cookies("login")="true"
	'	response.cookies("login").expires = date & " " & hour(n)+2 & ":" & minute(n) & ":" & second(n)
                rs("logintime")=now()
                rs.update
                rs.close
                set rs= nothing 
	else
		response.redirect "login.asp"
	end if
end if
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<HTML>
<HEAD><TITLE>Admin Page for Front Page</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<META content="barcode, bar code, tickets, events, instant " name=keywords>
<META content="MSHTML 5.50.4134.100" name=GENERATOR>
<script language="JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</HEAD>

<BODY text=#000000 vLink=#0000ff aLink=#0000ff link=#0000ff bgColor=#ffffff><!--Legal Logo Start-->
<DIV align=center>
<TABLE cellSpacing=0 cellPadding=7 width=600 border=0>
  <TBODY>
  <TR>
      <TD>&nbsp;</TD>
    <TD align=right>
        <DIV align=right></DIV>
      </TD></TR></TBODY></TABLE><!--Legal Logo End--><!-- Begin Main Top Menu -->
<TABLE cellSpacing=0 cellPadding=0 width=600 border=0>
  <TBODY>
  <TR>
      <TD>&nbsp;</TD>
    </TR></TBODY></TABLE><!-- End Main Top Menu --><!--Begin Normal Table-->
<TABLE cellSpacing=0 cellPadding=6 width=600 bgColor=#003399 border=0>
  <TBODY>
  <TR bgColor=#003399>
    <TD height=75>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" bgColor=#ffffff 
        border=0><TBODY>
        <TR>
          <TD width="77%" bgColor=#00ccff>
              <TABLE height="100%" cellSpacing=0 cellPadding=0 width="100%" 
            border=0>
                <TBODY> 
                <TR>
                <TD vAlign=top align=left width=15><IMG height=10 
                  src="images/dn_inside-ul.gif" 
                  width=10></TD>
                  <TD width="570">&nbsp;</TD>
                <TD vAlign=top align=right width=15><IMG height=10 
                  src="images/dn_inside-ur.gif" 
                  width=10></TD></TR><!-- start tts_user -->
              <TR>
			  	<td width=15>&nbsp;</td>
                <TD width=570>
                  <TABLE width="95%" align=center bgColor=#00ccff border=0>
                    <TBODY>
                    <TR>
                      <TD height=2>
                        <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                        align=center bgColor=#0034a0 border=0>
                          <TBODY>
                          <TR>
                            <TD vAlign=top align=left width=10 height=10><IMG 
                              height=10 
                              src="images/dn_inverse-ul.gif" 
                              width=10></TD>
                            <TD height=10 rowSpan=2>
                                <DIV align=center><B><FONT 
                              face="Arial, Helvetica, sans-serif" color=#ffffff 
                              size=+1>BarCode Ticket Management System</FONT></B></DIV>
                              </TD>
                            <TD vAlign=top align=right width=10 height=10><IMG 
                              height=10 
                              src="images/dn_inverse-ur.gif" 
                              width=10></TD></TR>
                          <TR>
                            <TD vAlign=bottom align=left width=10 
                              height=10><IMG height=10 
                              src="images/dn_inverse-ll.gif" 
                              width=10></TD>
                            <TD vAlign=bottom align=right width=10 
                              height=10><IMG height=10 
                              src="images/dn_inverse-lr.gif" 
                              width=10></TD></TR></TBODY></TABLE>
                    <TR>
                        <TD height=12> 
                          <DIV align=center><FONT 
                        face="Arial, Helvetica, sans-serif" size=2></FONT></DIV>
                             <FORM name=add action=addevent.asp method=post>
                            <TABLE cellSpacing=3 cellPadding=3 width="100%" 
                        align=center bgColor=#00ccff border=0>
                              <TBODY> 
                              <!-- #include file="menu.asp" -->
                              <TR vAlign=center bgColor=#00ccff> 
                                <TD colSpan=2 height="228"> 
                                  <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                              align=center bgColor=#0034a0 border=0>
                                    <TBODY> 
                                    <TR> 
                                      <TD vAlign=top align=left width=10 
                                height=10><IMG height=10 
                                src="images/dn_inverse-ul.gif" 
                                width=10></TD>
                                      <TD height=10 rowSpan=2> 
                                        <DIV align=center><FONT 
                                face="Arial, Helvetica, sans-serif" 
                                color=#ffffff size=2> <%=now %></FONT></DIV>
                                      </TD>
                                      <TD vAlign=top align=right width=10 
                                height=10><IMG height=10 
                                src="images/dn_inverse-ur.gif" 
                                width=10></TD>
                                    </TR>
                                    <TR> 
                                      <TD vAlign=bottom align=left width=10 
                                height=10><IMG height=10 
                                src="images/dn_inverse-ll.gif" 
                                width=10></TD>
                                      <TD vAlign=bottom align=right width=10 
                                height=10><IMG height=10 
                                src="images/dn_inverse-lr.gif" 
                                width=10></TD>
                                    </TR>
                                    </TBODY> 
                                  </TABLE>
                                  <center>Your Login is verified <br><br>
                                  <b>Welcome to Barcodeticket System.</b></center>
                                    <!--#include file ="main2.asp"-->
                                  <DIV></DIV>
                                </TD>
                              </TR>
                              <TR bgColor=#00ccff> 
                                <TD vAlign=top colSpan=2> 
                                  <CENTER>
                                    <P><FONT 
                              face="Arial, Helvetica, sans-serif"> </FONT> </P>
                                  </CENTER>
                                </TD>
                              </TR>
                              </TBODY> 
                            </TABLE>
 

      </form></TD></TR><!-- end body -->
                    <TR bgColor=#00ccff>
                      <TD vAlign=top height=2>
                        </TD></TR></TBODY></TABLE>
					</TD>
					<td width=15>&nbsp;</td>
				</TR>
              <TR>
                <TD vAlign=bottom align=left width=15><IMG height=10 
                  src="images/dn_inside-ll.gif" 
                  width=10></TD>
                  <TD width="570">&nbsp;</TD>
                <TD vAlign=bottom align=right width=15><IMG height=10 
                  src="images/dn_inside-lr.gif" 
                  width=10></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE><!--End Normal Table--><!-- Begin Privacy Policy -->

<!--#include file ="endfor.asp"-->
<!--#include file ="_closeconnection.asp"-->




</body>
</html>


