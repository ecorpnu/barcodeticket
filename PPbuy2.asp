<%
if Request.Cookies("demopp") <> "true" then
response.redirect("ppbuy.asp")
end if
%>
<!--#include file ="_OpenConnection.asp"-->
<!--#INCLUDE FILE="ADOVBS.INC"-->

<% 
   adults=cint(request("numseatadult"))

   nn=adults
   adultcost=request("d")
   childcost=request("c")
   productID=request("e")
   Avai=request("f")
   Costicket=(adults*adultcost)
   dol=costicket
   Fnum=avai-nn
   'response.write nn
   
   'response.write Costicket
   'response.write "yyy"
   'response.write Fnum
%>

<% If Request("e") <> "" OR Request.Form("card_name") <> "" Then %>


<% if request.cookies("salescompleted")= "true" then 
   response.cookies("salescompleted") = "false" 
   'response.redirect "erroradmin.asp"
end if

%>



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

</HEAD>


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
                  <TD width="570">&nbsp;</TD>
                <TD vAlign=top align=right width=15><IMG height=10 
                  src="images/dn_inside-ur.gif" 
                  width=10></TD></TR><!-- start tts_user -->
              <TR>
			  	<td width=15></td>
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
                              size=+1><!--#INCLUDE FILE="header.asp"--></FONT></B></DIV>
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
                            <TBODY> 
                            <tr> 
                              <td valign=top width="42%">&nbsp; </td>
                              <td valign=top width="53%" bgcolor=#00ccff>&nbsp;</td>
                            </tr>
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
                                color=#ffffff size=2>Check the details below. Use back button to make changes </FONT></DIV>
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
                                   
                                    <FORM ACTION="https://www.paypal.com/cgi-bin/webscr" METHOD="POST">


                                    Email: <% response.write (request.form("cust_email")) %> 
                                    <br>
                                    First Name: <% response.write (request.form ("first_name")) %> 
                                    <br>
                                    Last Name: <% response.write (request.form ("last_name")) %> 
                                    <br>
                                    Address: <% response.write (request.form ("cust_address1")) %> 
                                    <br>
                                    City: <% response.write (request.form ("cust_city")) %> 
                                    <br>
                                    Post/Zip Code: <% response.write (request.form ("cust_zip")) %> 
                                    <br>
                                    Event: <% response.write (request.form ("event3")) %> 
                                    <br>
                                    Date : <% response.write (request.form ("detval")) %> 
                                    <br>
                                    Number of Ticket Adult: <% response.write (request.form ("numseatadult")) %> 
                                    <br>
                                    
                                    
                                    <br>
                                    Total Cost: <%= formatcurrency(dol,2) %> <br>
<% 				   
                                   SET recSearch1=Server.CreateObject("ADODB.RecordSet")
			           recSearch1.ActiveConnection = Conn
				   recSearch1.CursorLocation = adUseClient
				   recSearch1.CursorType = adOpenDynamic
				   recSearch1.LockType = adLockOptimistic    
                                   product_id=request.form ("e")  
                                   'eventdate=request.form ("detval") 
                                   'response.write strpid
                                   'response.write eventdate                                 
                                   strSql="Select numticket from product where product_id="&product_id
			           recsearch1.open strsql
                                   if recsearch1("numticket") > 0 then
                                   num= recsearch1("numticket")
				   response.write " Ticket Availability " 
                                   response.write num 
                                   response.write "<p><center>"                                  
				   response.write "<INPUT TYPE=""image"" SRC=""http://images.paypal.com/images/x-click-but01.gif"" BORDER=""0"" NAME=""submit"" ALT=""Make payments with PayPal - it's fast, free and secure!"" width=""62"" height=""31"">"
                                   response.write "<p>"
                                   response.write "<FONT face=""Arial, Helvetica, sans-serif"">" 
                                   response.write "Once you click Paypal above, you will leave for paypal secure site. Please complete all steps until you see your ticket(s) being printed out." 
                                   response.write "</FONT></center>"                                  
			           else
                                   
                                   response.write "No more Tickets."
                                   response.write "<script>soldout = true;</script>"
                                   end if
                                   recsearch1.close
                                   set recsearch1 = nothing    	
                                
                                   
			           
                                   %>


<INPUT TYPE="hidden" NAME="redirect_cmd" VALUE="_xclick">
<INPUT TYPE="hidden" NAME="cmd" VALUE="_ext-enter">

<INPUT TYPE="hidden" NAME="business" VALUE="<%= ReadFile("text/ppemail.htm") %>">
<input type="hidden" name="notify_url" value="<%= ReadFile("text/pp.htm") %>">
<INPUT TYPE="hidden" NAME="undefined_quantity" VALUE="0">                         

<INPUT TYPE="hidden" NAME="return" VALUE="<%= ReadFile("text/pathprint.htm") %>/pp3webpay.asp">
<INPUT TYPE="hidden" NAME="cancel_return" VALUE="<%= ReadFile("text/pathprint.htm") %>/cancel.asp">                                      
                                   

                                      

<INPUT type=hidden name="amount" value="<%=formatnumber((dol),2)%>">
<INPUT TYPE="hidden" NAME="item_name" VALUE="<%= request.form("pidi")%>">
<input type=hidden name="item_number" value="<%= ProductID %>">
<input type=hidden name="on1" value="<%= nn %>" >
<INPUT TYPE="hidden" NAME="os1" VALUE="<%= request.form("numseatadult")%>">
<INPUT TYPE="hidden" NAME="on0" VALUE="<%= request.form("event3")%>">
<INPUT TYPE="hidden" NAME="os0" VALUE="<%= request.form("detval")%>">
<INPUT TYPE="hidden" NAME="custom" VALUE="<%= request.form("pw11")%>">
<input type="hidden" name="first_name" value="<%= request.form("first_name")%>">
<input type="hidden" name="last_name" value="<%= request.form("last_name")%>">
<input type="hidden" name="address_street" value="<%= request.form("cust_address1")%>">
<input type="hidden" name="address_zip" value="<%= request.form("cust_zip")%>">
<input type="hidden" name="address_city" value="<%= request.form("cust_city")%>">
<input type="hidden" name="payer_email" value="<%= request.form("cust_email")%>">
<input type="hidden" name="invoice" value="<%= session.sessionid %>">                           




                                     
                                                    

				 </form>
                                </center>
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
                  <TD width="570">&nbsp;</TD>
                <TD vAlign=bottom align=right width=15><IMG height=10 
                  src="images/dn_inside-lr.gif" 
                  width=10></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE><!--End Normal Table--><!-- Begin Privacy Policy -->
<!--#include file ="endfor.asp"-->
<!--#include file ="_CloseConnection.asp"-->

<% else %>
<% response.expires = 0
response.redirect ("ppbuy.asp")
end if
%>
	
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
