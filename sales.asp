<%
if request.cookies("login") <> "true" then
response.redirect("login.asp")
end if
%>
<!--#include file ="_OpenConnection.asp"-->
<!--#INCLUDE FILE="ADOVBS.INC"-->


<%
bydate = ""
if request.form("bydate")="bydate" then
	fromdate = request.form("fromdate") 
	untildate = request.form("untildate")
	'bydate = " where order_date > {ts '" & fromdate & " 00:00:00'} and order_date < {ts '" & untildate & " 00:00:00'}"
        bydate = " where order_date > " & prepdate(fromdate) & " and order_date < " & prepdate(untildate)
else
vfd = request.querystring("fd")
vud = request.querystring("ud")
if vfd <> "" and vud <> "" and len(vfd)=10 and len(vud)=10 then
	fromdate = request.querystring("fd") 
	untildate = request.querystring("ud")
	'bydate = " where order_date > {ts '" & fromdate & " 00:00:00'} and order_date < {ts '" & untildate & " 00:00:00'}"
         bydate = " where order_date > " & prepdate(fromdate) & " and order_date < " & prepdate(untildate) 
end if
end if
%>



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">


<HTML><HEAD><TITLE>Barcode Ticket System</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<META 
content="barcode, bar code, tickets, events, instant " 
name=keywords>
<META content="MSHTML 5.50.4134.100" name=GENERATOR>
<script language="javascript">
function check(f) {
	if ((f.fromdate.value == 'YYYY-MM-DD')||(f.untildate.value == 'YYYY-MM-DD'))
	{
		alert('Invalid date');
		return false;
	}
	var fdate = f.fromdate.value;
	var udate = f.untildate.value;
	fdate = new Date(fdate.substr(0,4),fdate.substr(5,2)-1,fdate.substr(8,2));
	udate = new Date(udate.substr(0,4),udate.substr(5,2)-1,udate.substr(8,2));
	if ( fdate.getTime() >= udate.getTime() )
	{
		alert('From Date must be smaller than Until Date');
		return false;
	}
	return true;
}
</script>
</HEAD>

<BODY text=#000000 vLink=#0000ff aLink=#0000ff link=#0000ff bgColor=#ffffff><!--Legal Logo Start-->
<!--  PopCalendar Start (frame name and id must match)--->
<iframe src="popcjs.htm" name="popFrame" id="popFrame" scrolling="no" frameborder="0" style="border:2 ridge;visibility:hidden;position:absolute;z-index:65535"></iframe>
<script>document.onclick=function() {document.getElementById("popFrame").style.visibility="hidden";}</script>
<!--  PopCalendar End --->

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
                  <TD width="570"> 
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
                                color=#ffffff size=2>All sales posted to <%=now %></FONT></DIV>
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
								  <form method="post" action="sales.asp" name="basedate" onsubmit="return check(this);">
								  	Search Order from:&nbsp;	
									<input id=dc1 readOnly size=12 name="fromdate" value=<%
												if bydate="" then 
													response.write """YYYY-MM-DD"""
												else
													response.write """" & fromdate & """"
												end if
												%>>
                                  	<input onClick="popFrame.fPopCalendar('dc1','dc1',event);" type=button value="V" name="button">
								  	&nbsp;until:&nbsp;
									<input id=dc2 readOnly size=12 name="untildate" value=<%
												if bydate="" then 
													response.write """YYYY-MM-DD"""
												else
													response.write """" & untildate & """"
												end if
												%>>
                                  	<input onClick="popFrame.fPopCalendar('dc2','dc2',event);" type=button value="V" name="button">
									<input type="hidden" name="bydate" value="bydate">
									<input type="submit" value="Submit">
								  </form>
                                  <P> 
                                  <CENTER>
                                    <%
									Dim rst
									Dim strSQL2
									if (bydate="") then
									strSQL2 = "Select * FROM orders" & bydate 
                                                                        else
                                                                        strSQL2 = "Select * FROM orders" & bydate 
                                                                        strSQL2 = strSQL2 & "order by Order_date desc " 
                                                                        end if 
									Set rst = Server.CreateObject("ADODB.Recordset")
									rst.Open strSQL2,conn, adOpenstatic,adlockReadonly, adCmdText
									' Number of Orders
									rowcount = 0
									While not rst.eof
									rowcount=rowcount + 1
									Rst.Movenext
									Wend
									response.write " Total Orders  " 
									response.write rowcount
									
									strcost = "Select sum(costprice) from orders" & bydate
									set rstamount = server.createObject("ADODB.Recordset")
									rstamount.open strcost,conn, adOPenstatic,_
									adlockReadonly, adCmdText
									
									response.write " Amount " 
									response.write rstamount(0)
									%>
                                    <center>
                                      <table border=1 cellspacing=0 cellpadding=3 width="80%" bgcolor="#CCCCFF">
                                        <!--<td><font size=-1><b>Comments</b></font></td>-->
                                        <td><font size=-1><b>Order Date</b></font></td>
                                        <td><font size=-1><b>Description </b></font></td>
                                        <td><font size=-1><b>Ticket Type</b></font></td>
                                        <td><font size=-1><b>Cost of Ticket</b></font></td>
                                        <td><font size=-1><b>Event Date </b></font></td>
                                        <td><font size=-1><b>Customer Name </b></font></td>
                                        <td><font size=-1><b>Customer Email </b></font></td>
                                        </tr>
                                        <%
										set rst = conn.execute(strSQL2)
										If not rst.eof then
											page = CInt(Request.QueryString("Page"))
											if Trim(Page) = "" then	page = 0
											strtxt = "SELECT COUNT(*) AS TotalItem FROM orders" & bydate
											set rs = conn.execute(strtxt)
											TotalItem = rs("TotalItem")
											rs.close
											set rs=nothing
											rst.move page
											i=page+1
											While (not rst.eof) and (i-page-10<=0)
										%>
                                        <tr> 
                                          <center>
                                            <!--<td><font size=-1><%= rst("comment_client") %></font></td>-->
                                            <td><font size=-1><%= rst("Order_date") %></font></td>
                                            <td><font size=-1><%= rst("Description") %></font></td>
                                            <td><font size=-1><%= rst("Ticket_cat") %></font></td>
                                            <td><font size=-1><%= rst("CostPrice") %></font></td>
                                            <td><font size=-1><%= rst("Event_date") %></font></td>
                                            <td><font size=-1><%= rst("Card_name") %></font></td>
                                            <td><font size=-1><a href= mailto:<%= rst("Cust_email") %>><%= rst("Cust_email") %></a></font></td>
                                          </center>
                                        </tr>
										<%
												i=i+1
												rst.movenext
											WEND
										else
										%>
										<tr><td colspan="8">
                                        <center>
                                          <table border=1 cellspacing=0 cellpadding=3 width="80%" bgcolor="#EEEEE">
                                            <tr> 
                                              <td> 
                                                <center>
                                                  <b>You do not have any orders yet </b>
                                                </center>
                                              </td>
                                            </tr>
                                          </table>
                                        </center>
										</td></tr>
										<%
										response.write ""
										end if 
										%>
                                      </table>
										<%
										rst.close
										%>
                                    </center>
                                    <P></P>
                                  </CENTER>
                                  <DIV><center>
								  <%
									'Computed Previous position
									if page <> 0 then
									  page1 = page - 10
									%>
									  <a href=sales.asp?page=<% =page1 %>&fd=<% =server.urlencode(fromdate) %>&ud=<% =server.urlencode(untildate) %>>[Previous]</a>&nbsp&nbsp
									<%
									end if
									'Computed Next position
									if (page < TotalItem) then
									  page2 = page + 10
									  if page2 < TotalItem then
									%>
									  <a href=sales.asp?page=<% =page2 %>&ud=<% =server.urlencode(untildate) %>&fd=<% =server.urlencode(fromdate) %>>[Next]</a>
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


