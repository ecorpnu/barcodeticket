<%
if request.cookies("login") <> "true" then
response.redirect("login.asp")
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include file ="_OpenConnection.asp"-->
<!--#INCLUDE FILE="ADOVBS.INC"-->


<HTML><HEAD><TITLE>Barcode Ticket System</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<META 
content="barcode, bar code, tickets, events, instant " 
name=keywords>
<META content="MSHTML 5.50.4134.100" name=GENERATOR>
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
                <TD>&nbsp;</TD>
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
                        
                            <TABLE cellSpacing=3 cellPadding=3 width="100%" 
                        align=center bgColor=#00ccff border=0>
                              <!-- #include file="menu.asp" -->
                              <TBODY> 
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
                                color=#ffffff size=2>All 'Deactived' Events posted to <%=now %> </FONT></DIV>
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
<center><A HREF="Seeevents.asp">Click here to see Active Events</a></center>
                                  <P> 
                                  <CENTER>

                                    <%

Dim rst
Dim strSQL2

strSQL2 = "Select * FROM product where status = 'Deactivated' order by event_date desc"
Set rst = Server.CreateObject("ADODB.Recordset")
rst.Open strSQL2,conn, adOpenstatic,_
adlockReadonly, adCmdText

%>
                                    <table BORDER=1 CELLSPACING=0 CELLPADDING=3 WIDTH="80%" BGCOLOR="#CCCCFF">
                                      <td><font size=-1><B>Status</B></font></td>
                                      <TD><font size=-1><B>Type of Event</B></font></TD>
                                      <TD><font size=-1><B> Adult Cost</B></font></TD>
                                      <!--<TD><font size=-1><B> Child Cost</B></font></TD>-->
                                      <TD><font size=-1><B>Event Details</B></font></TD>
                                      <TD><font size=-1><B>Event Date</B></font></TD>
                                      
                                      <TD><font size=-1><B>IP Address</B></font></TD>
                                      <TD><font size=-1><B>Posted</B></font></TD>
                                      <TD><font size=-1><B>Sold</B></font></TD>     
                                      </TR>
                                      <%
set rst = conn.execute(strSQL2)
If not rst.eof then
	page = CInt(Request.QueryString("Page"))
	if Trim(Page) = "" then	page = 0
	strtxt = "SELECT COUNT(*) AS TotalItem FROM product where status = 'Deactivated'"
	set rs = conn.execute(strtxt)
	TotalItem = rs("TotalItem")
	rs.close
	set rs=nothing
	rst.move page
	i=page+1
	While (not rst.eof) and (i-page-10<=0)
%>
                                      <TR> 
                                        <center>
                                          <TD><font size=-1><A HREF="Update.asp?prodId=<%= rst("product_ID")%>"><%=rst("status") %></a><br>
                                          <% if rst("event_date") < now() then 
                                           response.write "Expired" 
                                           end if %> </font></TD>
                                          <TD><font size=-1><%= rst("Description") %></font></TD>
                                          <TD><font size=-1><%= formatcurrency(rst("Unit_Price_adult"),2) %></font></TD>
                                          <!--<TD><font size=-1><%= rst("Unit_Price_child") %></font></TD>-->
                                          <TD><font size=-1><%= rst("Event2") %></font></TD>
                                          <TD><font size=-1><%= rst("Event_date") %></font></TD>
                                          
                                          <TD><font size=-1><%= rst("IPaddress") %></font></TD>
                                          <TD><font size=-1><%= rst("Datepost") %></font></TD>
				          <TD><font size=-1>
<%                                        
strsold = "SELECT COUNT(*) AS Totalsold FROM Orders where product_id ="& rst("product_ID")
	set rst6 = conn.execute(strsold)
	response.write rst6("TotalSold")
	rst6.close
	set rst6=nothing
%></font></TD>
                                        </center>
                                      </TR>
                                      <%
		i = i + 1
		rst.movenext
	WEND
else

%>
                                      <center>
                                        <table BORDER=1 CELLSPACING=0 CELLPADDING=3 WIDTH="80%" BGCOLOR="#EEEEE">
                                          <tr> 
                                            <td> 
                                              <center>
                                                You do not have any events yet 
                                              </center>
                                            </td>
                                          </tr>
                                        
                                      <%
response.write ""

end if 

%>
                                    </Table>
                                    <%
rst.close
%>
                                    <!--#include file ="_CloseConnection.asp"-->
                                  </center>
                                  <P></P>
                                  <DIV><center>
								  <%
									'Computed Previous position
									if page <> 0 then
									  page1 = page - 10
									%>
									  <a href=seedeevents.asp?page=<% =page1 %>>[Previous]</a>&nbsp&nbsp
									<%
									end if
									'Computed Next position
									if (page < TotalItem) then
									  page2 = page + 10
									  if page2 < TotalItem then
									%>
									    <a href=seedeevents.asp?page=<% =page2 %>>[Next]</a>
									<%
									  end if
									end if
									%></center>
								  </DIV>
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
				<td width=15></td>
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
