<%
if Request.Cookies("login") <> "true" then
response.redirect("login.asp")
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include file ="_OpenConnection.asp"-->
<!--#INCLUDE FILE="ADOVBS.INC"-->

<% if request("add")="adddata" then 
        SET recSearch=Server.CreateObject("ADODB.RecordSet")

		sdate = trim(request("event_date"))
		'if mid(sdate,5,1) = "/" then
		'	sdate = left(sdate,3) & "0" & right(sdate,6)
		'end if		
		'sdate = sdate & " " & request("selecthour") & ":" & request("selectminute1") & request("selectminute2") & " " & request("selectap")
		sdate = sdate & " " & request("selecthour") & ":" & request("selectminute1") & request("selectminute2") & ":00"
		'strSQL2 = "SELECT * FROM product WHERE event_date={ts '" & sdate & "'} AND event2='" & request("event2") & "'"
		'strSQL2 = "SELECT * FROM product WHERE event_date={ts '" & sdate & "'} AND event2='" & request("event2") & "'AND description='" & request("description") & "'"
                 strSQL2 = "SELECT * FROM product WHERE event_date= " & prepdate(sdate) & " AND event2='" & request("event2") & "'AND description='" & request("description") & "'"
                set rs = conn.execute(strSQL2)
		if (rs.bof) and (rs.eof) then
			strSQL2 = "SELECT * FROM product" 
			Set rst2 = Server.CreateObject("ADODB.Recordset")
			rst2.Open strSQL2, conn, adOpenStatic, adLockOptimistic, adCmdText 

			rst2.addnew
			rst2("description")= request("description")
			rst2("unit_price_adult")=request("unit_price_adult")
			'rst2("unit_price_child")=request("unit_price_child")
                        rst2("status")="Active"
			rst2("event_date")= sdate
			rst2("event2")=request("event2")
                        rst2("datepost")=now()
			if request("comments") = "" then
				rst2("orgcomments") = " "
			else rst2("orgcomments")=request("comments")
			end if
                        if request("imagepath") = "" then
				rst2("imagepath") = ReadFile("text/ti.htm")
			else rst2("imagepath")=request("imagepath")
			end if

            rst2("IPaddress")=Request.ServerVariables("REMOTE_ADDR")
            rst2("numticket")=request("numticket")
            rst2.Update
			eventexist = false
			rst2.Close
		else eventexist = true
		end if

'  If Err.Number = 0 Then
'    Response.Write  " Thank you. Your Event Added. "  
'  Else
'    Response.Write  " <center>Please email us with your details. An error has occured.</center>"
'  End If
end if
   
	       
%>
<HTML><HEAD><TITLE>Barcode Ticket System</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<META 
content="barcode, bar code, tickets, events, instant " 
name=keywords>
<META content="MSHTML 5.50.4134.100" name=GENERATOR>
<script language="javascript">
function checkform(f)
{
	if (f.event_date.value == 'YYYY-MM-DD')
	{
		alert('Invalid date');
		return false;
	}
	if ((f.description.value == '') || (f.event2.value == '') || (f.unit_price_adult.value == '') ||
		(f.numticket.value == '')) 
	{
		alert('Please complete this form completely');
		return false;
	}
	var now = new Date();
	var edate = f.event_date.value;
	var fdate = new Date(edate.substr(0,4),edate.substr(5,2)-1,edate.substr(8,2));
	if (fdate.getTime() < now.getTime())
	{
		alert('You can not choose this date, try another date in the future');
		return false;
	}
	//if (f.unit_price_child.value == '') f.unit_price_child.value = '0';
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
				  <td width=570>&nbsp;</td>
                  <TD vAlign=top align=right width=15><IMG height=10 
                  src="images/dn_inside-ur.gif" 
                  width=10></TD>
                </TR><!-- start tts_user -->
              <TR>
			  	   <td width=15>&nbsp;</td>
				   
                  <td width=570> 
                    <table width="95%" align=center bgcolor=#00ccff border=0>
                      <tbody> 
                      <tr> 
                        <td height=2> 
                          <table cellspacing=0 cellpadding=0 width="100%" 
                        align=center bgcolor=#0034a0 border=0>
                            <tbody> 
                            <tr> 
                              <td valign=top align=left width=10 height=10><img 
                              height=10 
                              src="images/dn_inverse-ul.gif" 
                              width=10></td>
                              <td height=10 rowspan=2> 
                                <div align=center><b><font 
                              face="Arial, Helvetica, sans-serif" color=#ffffff 
                              size=+1>BarCode Ticket Management System</font></b></div>
                              </td>
                              <td valign=top align=right width=10 height=10><img 
                              height=10 
                              src="images/dn_inverse-ur.gif" 
                              width=10></td>
                            </tr>
                            <tr> 
                              <td valign=bottom align=left width=10 
                              height=10><img height=10 
                              src="images/dn_inverse-ll.gif" 
                              width=10></td>
                              <td valign=bottom align=right width=10 
                              height=10><img height=10 
                              src="images/dn_inverse-lr.gif" 
                              width=10></td>
                            </tr>
                            </tbody>
                          </table>
                      <tr> 
                        <td height=12> 
                          <div align=center><font face="Arial, Helvetica, sans-serif" size=2>
						  <% if (request("add")="adddata") and (not eventexist) then
								  If Err.Number = 0 Then
									    Response.Write  "<b>Thank you. Your Event Added.<b>"
								  Else
									    Response.Write  "<b>Please email us with your details. An error has occured.<b>"
								  End If
						      end if
						  %>
						  </font></div>
                          <form name=add action=addevent.asp method=post onSubmit="return checkform(this);">
                            <table cellspacing=3 cellpadding=3 width="100%" 
                        align=center bgcolor=#00ccff border=0>
                              <!-- #include file="menu.asp" -->
                              <tbody> 
                              <tr valign=center bgcolor=#00ccff> 
                                <td colspan=2 height="228"> 
                                  <table cellspacing=0 cellpadding=0 width="100%" 
                              align=center bgcolor=#0034a0 border=0>
                                    <tbody> 
                                    <tr> 
                                      <td valign=top align=left width=10 
                                height=10><img height=10 
                                src="images/dn_inverse-ul.gif" 
                                width=10></td>
                                      <td height=10 rowspan=2> 
                                        <div align=center><font 
                                face="Arial, Helvetica, sans-serif" 
                                color=#ffffff size=2>Add Your New Event: Follow 
                                          the example below. Note the time <%=now %></font></div>
                                      </td>
                                      <td valign=top align=right width=10 
                                height=10><img height=10 
                                src="images/dn_inverse-ur.gif" 
                                width=10></td>
                                    </tr>
                                    <tr> 
                                      <td valign=bottom align=left width=10 
                                height=10><img height=10 
                                src="images/dn_inverse-ll.gif" 
                                width=10></td>
                                      <td valign=bottom align=right width=10 
                                height=10><img height=10 
                                src="images/dn_inverse-lr.gif" 
                                width=10></td>
                                    </tr>
                                    </tbody> 
                                  </table>
                                  <center>
								  	<% if (eventexist) then response.write "<b>Event has already been added in database. Please select another time.<br>No same event can be in the same time slot</b><br><br>" %>
                                    <table width="100%">
                                      <tr> 
                                        <td>Type of Ticket</td>
                                        <td> 
                                          <input size=30 value="Conference etc" name="description">
                                        </td>
                                      </tr>
                                      <tr> 
                                        <td>Event Description</td>
                                        <td> 
                                          <input size=30 value="name of event and venue" name="event2">
                                        </td>
                                      </tr>
                                      <tr> 
                                        <td>Date and Time of Event </td>
                                        <td> 
                                          <input id=dc1 readOnly size=12 name="event_date" value="YYYY-MM-DD">
                                          <input onClick="popFrame.fPopCalendar('dc1','dc1',event);" type=button value="V" name="button">
                                          <b> / </b> 
                                          <select name="selecthour">
                                            <option value=00 selected>00</option>
                                            <option value=01>01</option>
                                            <option value=02>02</option>
                                            <option value=03>03</option>
                                            <option value=04>04</option>
                                            <option value=05>05</option>
                                            <option value=06>06</option>
                                            <option value=07>07</option>
                                            <option value=08>08</option>
                                            <option value=09>09</option>
                                            <option value=10>10</option>
                                            <option value=11>11</option>
                                            <option value=12>12</option>
											<option value=13>13</option>
											<option value=14>14</option>
											<option value=15>15</option>
											<option value=16>16</option>
											<option value=17>17</option>
											<option value=18>18</option>
											<option value=19>19</option>
											<option value=20>20</option>
											<option value=21>21</option>
											<option value=22>22</option>
											<option value=23>23</option>
                                          </select>
                                          <b> : </b> 
                                          <select name="selectminute1">
                                            <option value=0 selected>0</option>
                                            <option value=1>1</option>
                                            <option value=2>2</option>
                                            <option value=3>3</option>
                                            <option value=4>4</option>
                                            <option value=5>5</option>
                                          </select>
                                          <select name="selectminute2">
                                            <option value=0 selected>0</option>
                                            <option value=1>1</option>
                                            <option value=2>2</option>
                                            <option value=3>3</option>
                                            <option value=4>4</option>
                                            <option value=5>5</option>
                                            <option value=6>6</option>
                                            <option value=7>7</option>
                                            <option value=8>8</option>
                                            <option value=9>9</option>
                                          </select>
                                          <!-- <select name="selectap">
                                            <option value=AM selected>AM</option>
                                            <option value=PM>PM</option>
                                          </select> -->
                                        </td>
                                      </tr>
                                      <tr> 
                                        <td>Cost of Event (Cents ok) This version has no Child</td>
                                        <td> 
                                          <input maxlength=7 size=5 value="10.00" name=unit_price_adult>
                                          <!--/ 
                                          <input maxlength=3 size=3 value="0" name=unit_price_child>
                                          &nbsp; &nbsp; (Adult/Child)--> Currency Format 0.00 </td>
                                      </tr>
                                      <tr>  
                                        <td>Number of Tickets</td>
                                        <td> 
                                          <input size=30 value=10 name="numticket">
                                        </td>
                                      </tr>
                                      <tr> 
                                        <td>Image Path (edit after "images/") or leave blank</td>
                                        <td> 
                                          <input size=30 value=images/abc.jpg name="imagepath">
                                        </td>
                                      </tr>
                                        <td>Comment/Description/Refund Policy 
                                        </td>
                                        <td> 
                                          <textarea name=Comments rows=10 wrap=virtual cols=40 >your comments here </textarea>
                                        </td>
                                      </tr>
                                      <input type=hidden value=adddata name=add>
                                      <tr> 
                                        <td> 
                                          <input type=submit value="Add Your Event" name="submit">
                                        </td>
                                        <td>&nbsp;</td>
                                      </tr>
                                    </table>
                                  </center>
                                  <div></div>
                                </td>
                              </tr>
                              <tr bgcolor=#00ccff> 
                                <td valign=top colspan=2> 
                                  <center>
                                    <p><font 
                              face="Arial, Helvetica, sans-serif"> </font> </p>
                                  </center>
                                </td>
                              </tr>
                              </tbody> 
                            </table>
                          </form>
                        </td>
                      </tr>
                      <!-- end body -->
                      <tr bgcolor=#00ccff> </tr>
                      </tbody>
                    </table>
                  </td>
                  <TD width="15">&nbsp; </TD>
                </TR>
              <TR>
                  <TD vAlign=bottom align=left width=15><IMG height=10 
                  src="images/dn_inside-ll.gif" 
                  width=10></TD>
				  <td width=570>&nbsp;</td>
                  <TD vAlign=bottom align=right width=15><IMG height=10 
                  src="images/dn_inside-lr.gif" 
                  width=10></TD>
                </TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE><!--End Normal Table--><!-- Begin Privacy Policy -->

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
