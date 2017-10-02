<%
if request.cookies("login") <> "true" then
response.redirect("login.asp")
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include file ="_OpenConnection.asp"-->
<!--#INCLUDE FILE="ADOVBS.INC"-->

<!--#include file ="1.asp"-->

<% ' error check
if TRIM( Request( "product_ID" ) ) = "" then
  errorForm "seeevents.asp", "You must have an event."
end if

%>
<HTML><HEAD><TITLE>Barcode Ticket System</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<META 
content="barcode, bar code, tickets, events, instant " 
name=keywords>
<META content="MSHTML 5.50.4134.100" name=GENERATOR></HEAD>
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
                <TD>&nbsp;</TD>
                <TD vAlign=top align=right width=15><IMG height=10 
                  src="images/dn_inside-ur.gif" 
                  width=10></TD></TR><!-- start tts_user -->
              <TR>
			  	<td width=15>&nbsp;</td>
                <TD>
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
                        
                            
                          <TABLE cellSpacing=3 cellPadding=3 width="100%" 
                        align=center bgColor=#00ccff border=0>
                            <!-- #include file="menu.asp" -->
                            <TBODY> 
                            <TR vAlign=center bgColor=#00ccff> 
                              <TD colSpan=2 height="228" valign="top"> 
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
                                color=#ffffff size=2>Status <%=now %></FONT></DIV>
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
                                <CENTER>
                                  <%

'On Error Resume Next

' Declare the variables

Dim rst
Dim strSQL

strSQL = "SELECT * FROM product WHERE product_ID =" & request("product_ID")

Set rst = Server.CreateObject("ADODB.Recordset")
rst.Open strSQL, conn, adOpenStatic, _
 adLockOptimistic, adCmdText 

If Request("Submit1")= "Activate" Then
  ' Update record

rst("Description")= Request.Form("Description")
rst("Unit_price_adult")= Request.Form("Unit_price_adult")
rst("Unit_price_child")= Request.Form("Unit_price_child")
sdate = trim(request("Event_date"))
'if mid(sdate,5,1) = "/" then
'	sdate = left(sdate,3) & "0" & right(sdate,6)
'end if
sdate = sdate & " " & request("selecthour") & ":" & request("selectminute1") & request("selectminute2") & ":00"
rst("Event_date")= cdate(sdate)
rst("Event2")= Request.Form("Event2")
rst("numticket")= Request.Form("numticket")
if request.form("OrgComments") = "" then
	rst("OrgComments") = " "
else
	rst("OrgComments")= Request.Form("OrgComments")
end if
if request("imagepath") = "" then
	rst("imagepath") = ReadFile("text/ti.htm")
else 
	rst("imagepath")=request("imagepath")
end if
rst("IPAddress")= Request.ServerVariables("REMOTE_ADDR") 
rst("status")="Active"
rst.Update

  If Err.Number = 0 Then
    Response.Write  " Record/Event Updated!"
  Else
    Response.Write  " Record/Event Update Failed!"
  End If
Else
  ' Delete record
  rst("status")="Deactivated"
  rst.Update
  If Err.Number = 0 Then
    Response.Write  " Record/Event Deactivated!"
  Else
    Response.Write  " Record Deactivation Failed!"
  End If
End If

' Clean up time
rst.Close



%>
                                  <br>
                                  <A HREF="seeevents.asp">Your Current Events 
                                  (Summary) is as appended here:</a> 
                                </center>
                                
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



 

      </TD></TR><!-- end body -->
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
                <TD>&nbsp;</TD>
                <TD vAlign=bottom align=right width=15><IMG height=10 
                  src="images/dn_inside-lr.gif" 
                  width=10></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE><!--End Normal Table--><!-- Begin Privacy Policy -->
<!--#include file ="endfor.asp"-->
<!--#include file ="_closeconnection.asp"-->
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