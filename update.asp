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
if TRIM( Request( "prodid" ) ) = "" then
  errorForm "seeevents.asp", "You must choose an event."
end if

%>

<HTML><HEAD><TITLE>Barcode Ticket System</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<META 
content="barcode, bar code, tickets, events, instant " 
name=keywords>
<META content="MSHTML 5.50.4134.100" name=GENERATOR>
<script language="javascript">


function checkform1(f)
{
	if (document.form.Event_date.value == 'YYYY-MM-DD')
	{
		alert('Invalid date');
		return false;
	}
	if ((document.form.description.value == '') || (document.form.Event2.value == '') || (document.form.Unit_Price_Adult.value == ''))
	{
		alert('Please complete this form completely');
		return false;
	}
	var now = new Date();
	var edate = document.form.Event_date.value;
	var fdate = new Date(edate.substr(0,4),edate.substr(5,2)-1,edate.substr(8,2));
	if ( (fdate.getTime() < now.getTime()) )
	{
		alert('You can not choose this date, try another date in the future');
		return false;
	}
	//if (document.form.Unit_price_child.value == '') document.form.Unit_price_child.value = '0';
	//return true;
}





</script>
</HEAD>

<BODY text=#000000 vLink=#0000ff aLink=#0000ff link=#0000ff bgColor=#ffffff>

<!--  PopCalendar Start (frame name and id must match)--->
<iframe src="popcjs.htm" name="popFrame" id="popFrame" scrolling="no" frameborder="0" style="border:2 ridge;visibility:hidden;position:absolute;z-index:65535"></iframe>
<script>document.onclick=function() {document.getElementById("popFrame").style.visibility="hidden";}</script>
<!--  PopCalendar End --->

<!--Legal Logo Start-->
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
					  <TABLE cellSpacing=3 cellPadding=3 width="100%" 
                        align=center bgColor=#00ccff border=0>
                            <tbody> </TBODY>
                          </TABLE>
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
                              width=10></TD>
							  </TR>
							  </TBODY>
							  </TABLE>
                    <TR>
                        <TD height=12> 
                          
                       <FORM METHOD="Post" name="form" Action="update2.asp">
                            
                          
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
                                color=#ffffff size=2><%=now %></FONT></DIV>
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

' Brings out the event once a Product_ID is identified.

Dim rst
Dim strSQL2



strSQL2 = "Select * FROM product WHERE product_ID = " & request("prodID")  

Set rst = Server.CreateObject("ADODB.Recordset")
rst.Open strSQL2,conn, adOpenstatic,_
adlockReadonly, adCmdText

If Not rst.eof then
dateofevent = rst("event_date")

%>
							</center>
                            <table>
                              <tr> 
                                <TD>Alert-ID </TD>
                                <TD> 
                                  <% = rst("product_ID")%>
                                </TD>
                              </tr>
                              <TR> 
                                <TD>Description</TD>
                                <TD> 
                                  <INPUT maxLength=54 name="description" value="<%=rst("description") %>" size=20>
                                </td>
                              </tr>
                              <TR> 
                                <TD>Unit Price Adult</TD>
                                <TD> 
                                  <INPUT maxLength=54 name="Unit_Price_Adult" value="<%=rst("Unit_Price_Adult") %>" size=20>
                                Currency format 0.00 </td> 
                              </tr>
                              <!--<TR> 
                                <TD>Unit Price Child</TD>
                                <TD>--> 
                                  <INPUT Type=hidden name="Unit_price_child" value="0" size=20>
                                </td> 
                              </tr>
                               <TR> 
                                <TD>Image Path<br> (images/123.jpg or leave blank)</TD>
                                <TD> 
                                  <INPUT maxLength=54 name="imagepath" value="<%=rst("imagepath") %>" size=20>
                                </td>
                              </tr>
                              <TR> 
                                <TD>Event Details</TD>
                                <TD> 
                                  <INPUT maxLength=54 name="Event2" value="<%=rst("Event2") %>" size=20>
                                </td>
                              </tr>
                              <TR> 
                                <TD>Number of Tickets</TD>
                                <TD> 
                                  <INPUT maxLength=54 name="Numticket" value="<%=rst("NumTicket") %>" size=20>
                                </td>
                              </tr> 
                              <TR> 
                                <TD>Date of Event</TD>
                                <TD> 
								  <%
								    y = year(dateofevent)
									m = month(dateofevent)
									d = day(dateofevent)
									if len(m) = 1 then m = "0" & m
									if len(d) = 1 then d = "0" & d
								  %>
                                  <INPUT readOnly id=dc1 maxLength=12 name="Event_date" value="<%= y & "-" & m & "-" & d %>" size=10>
								  <INPUT onclick="popFrame.fPopCalendar('dc1','dc1',event);" type=button value="V"><b> / </b>
                                          <select name="selecthour">
										  	  <%
											  h = hour(dateofevent)
											  'if h>12 then
											  '	h = h-12
											  'end if
											  %>
										      <option value=00 <% if (h="0") then response.write "selected" %>>00</option>
											  <option value=01 <% if (h="1") then response.write "selected" %>>01</option>
											  <option value=02 <% if (h="2") then response.write "selected" %>>02</option>
											  <option value=03 <% if (h="3") then response.write "selected" %>>03</option>
											  <option value=04 <% if (h="4") then response.write "selected" %>>04</option>
											  <option value=05 <% if (h="5") then response.write "selected" %>>05</option>
											  <option value=06 <% if (h="6") then response.write "selected" %>>06</option>
											  <option value=07 <% if (h="7") then response.write "selected" %>>07</option>
											  <option value=08 <% if (h="8") then response.write "selected" %>>08</option>
											  <option value=09 <% if (h="9") then response.write "selected" %>>09</option>
											  <option value=10 <% if (h="10") then response.write "selected" %>>10</option>
											  <option value=11 <% if (h="11") then response.write "selected" %>>11</option>
											  <option value=12 <% if (h="12") then response.write "selected" %>>12</option>
											  <option value=13 <% if (h="13") then response.write "selected" %>>13</option>
											  <option value=14 <% if (h="14") then response.write "selected" %>>14</option>
											  <option value=15 <% if (h="15") then response.write "selected" %>>15</option>
											  <option value=16 <% if (h="16") then response.write "selected" %>>16</option>
											  <option value=17 <% if (h="17") then response.write "selected" %>>17</option>
											  <option value=18 <% if (h="18") then response.write "selected" %>>18</option>
											  <option value=19 <% if (h="19") then response.write "selected" %>>19</option>
											  <option value=20 <% if (h="20") then response.write "selected" %>>20</option>
											  <option value=21 <% if (h="21") then response.write "selected" %>>21</option>
											  <option value=22 <% if (h="22") then response.write "selected" %>>22</option>
											  <option value=23 <% if (h="23") then response.write "selected" %>>23</option>
                                          </select><b> : </b>
										  <select name="selectminute1">
										      <% if len(minute(dateofevent))=2 then %>
										      <option value=0 <% if left(minute(dateofevent),1)="0" then response.write "selected" %>>0</option>
											  <option value=1 <% if left(minute(dateofevent),1)="1" then response.write "selected" %>>1</option>
											  <option value=2 <% if left(minute(dateofevent),1)="2" then response.write "selected" %>>2</option>
											  <option value=3 <% if left(minute(dateofevent),1)="3" then response.write "selected" %>>3</option>
											  <option value=4 <% if left(minute(dateofevent),1)="4" then response.write "selected" %>>4</option>
											  <option value=5 <% if left(minute(dateofevent),1)="5" then response.write "selected" %>>5</option>
											  <% else %>
											  <option value=0 "selected">0</option>
											  <option value=1>1</option>
											  <option value=2>2</option>
											  <option value=3>3</option>
											  <option value=4>4</option>
											  <option value=5>5</option>
											  <% end if %>
                                          </select>
										  <select name="selectminute2">
										  	  <% if len(minute(dateofevent))=2 then %>
										      <option value=0 <% if right(minute(dateofevent),1)="0" then response.write "selected" %>>0</option>
											  <option value=1 <% if right(minute(dateofevent),1)="1" then response.write "selected" %>>1</option>
											  <option value=2 <% if right(minute(dateofevent),1)="2" then response.write "selected" %>>2</option>
											  <option value=3 <% if right(minute(dateofevent),1)="3" then response.write "selected" %>>3</option>
											  <option value=4 <% if right(minute(dateofevent),1)="4" then response.write "selected" %>>4</option>
											  <option value=5 <% if right(minute(dateofevent),1)="5" then response.write "selected" %>>5</option>
											  <option value=6 <% if right(minute(dateofevent),1)="6" then response.write "selected" %>>6</option>
											  <option value=7 <% if right(minute(dateofevent),1)="7" then response.write "selected" %>>7</option>
											  <option value=8 <% if right(minute(dateofevent),1)="8" then response.write "selected" %>>8</option>
											  <option value=9 <% if right(minute(dateofevent),1)="9" then response.write "selected" %>>9</option>
											  <% else %>
											  <option value=0 <% if minute(dateofevent)="0" then response.write "selected" %>>0</option>
											  <option value=1 <% if minute(dateofevent)="1" then response.write "selected" %>>1</option>
											  <option value=2 <% if minute(dateofevent)="2" then response.write "selected" %>>2</option>
											  <option value=3 <% if minute(dateofevent)="3" then response.write "selected" %>>3</option>
											  <option value=4 <% if minute(dateofevent)="4" then response.write "selected" %>>4</option>
											  <option value=5 <% if minute(dateofevent)="5" then response.write "selected" %>>5</option>
											  <option value=6 <% if minute(dateofevent)="6" then response.write "selected" %>>6</option>
											  <option value=7 <% if minute(dateofevent)="7" then response.write "selected" %>>7</option>
											  <option value=8 <% if minute(dateofevent)="8" then response.write "selected" %>>8</option>
											  <option value=9 <% if minute(dateofevent)="9" then response.write "selected" %>>9</option>
											  
											  <% end if %>
                                          </select>
										  <!-- <select name="selectap">
										      <option value=AM <% if (right(rst("event_date"),2)="AM") then response.write "selected" %>>AM</option>
											  <option value=PM <% if (right(rst("event_date"),2)="PM") then response.write "selected" %>>PM</option>
                                          </select> -->
                                </td>
                              </tr>
                              <TR> 
                                <TD>Comments</TD>
                                <TD> 
                                  <INPUT maxLength=54 name="OrgComments" value="<%=rst("OrgComments") %>" size=20>
                                </td>
                              </tr>
                              <TR> 
                                <TD>IPAddress</TD>
                                <TD><%=rst("IPAddress") %></td>
                              </tr>
                              <INPUT NAME = product_ID TYPE= "HIDDEN" VALUE ="<%=rst("product_ID")%>">
                            </Table>
                            <BR>
                            <INPUT TYPE = "SUBMIT" NAME="SUBMIT1" VALUE="Activate" onclick="return checkform1()">
                            <INPUT TYPE = "SUBMIT" NAME="SUBMIT2" VALUE="Deactivate">
                            <BR>
                            <BR>
                            <% 
else
%>
                            <B>NO record was found</B> 
                            <%
End if
%>
                            <%
rst.close
%>
                            <P></P>
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
				  
				</form>


 

      </TD></TR><!-- end body -->
                    <TR bgColor=#00ccff>
                      <TD vAlign=top height=2>
                        <DIV align=center></DIV></TD></TR></TBODY></TABLE>
						</TD>
						
                  <TD width="15">&nbsp;</TD>
						</TR>
              <TR>
                <TD vAlign=bottom align=left width=15><IMG height=10 
                  src="images/dn_inside-ll.gif" 
                  width=10></TD>
                  <TD width="570">&nbsp;</TD>
                <TD vAlign=bottom align=right width=15><IMG height=10 
                  src="images/dn_inside-lr.gif" 
                  width=10></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></td></tr></TBODY></table><!--End Normal Table--><!-- Begin Privacy Policy -->
<!--#include file ="endfor.asp"-->
<!--#include file ="_closeconnection.asp"-->
