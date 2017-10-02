<% Response.Cookies("demopp") = "true" %>
<!--
00ccff => ffe9b0
0034a0 => 3564a9
-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include file ="_OpenConnection.asp"-->
<!--#INCLUDE FILE="ADOVBS.INC"-->


<HTML><HEAD><TITLE>Barcode Ticket System</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<META 
content="barcode, bar code, tickets, events, instant " 
name=keywords>
<META content="MSHTML 5.50.4134.100" name=GENERATOR>
<script language="JavaScript">
<!--
var soldout = false;
var numticket = 0;

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name;
		switch(nm){
					case 'cust_email': nm = 'Email';break;
					case 'first_name': nm = 'First Name';break;
                                        case 'last_name': nm = 'Last Name';break;
					case 'cust_address1': nm = 'Address';break;
					case 'cust_city': nm = 'City';break;
					case 'cust_zip': nm = 'Zip Code';break;
					case 'pw11': nm = 'Password';break;
                                        
			  }
	if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+'E-mail address is required.\n';
      } else if (test!='R') {
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (val<min || max<val) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}

function checkform(f)
{
	//check sold out
	if (soldout) 
	{
	  alert('Ticket was sold out');
	  return false;
	}
	//check date valid
/*	var date1 = new Date(f.detval.value);
	var date2 = new Date();
	date1.toUTCString();
	date2.toUTCString();
	var diff = date1.getTime() - date2.getTime();
	if (diff < 600000)
	{
		alert('This ticket closed');
		return false;
	}*/
	//check num ticket
	var ticketorder = parseInt(f.numseatadult.value); // + parseInt(f.numchild.value)
	if (ticketorder > numticket)
	{
		alert('Max order is ' + numticket);
		return false;
	}
	return true;
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
                        face="Arial, Helvetica, sans-serif" size=2><B>Use the 
                            form below to buy a Ticket : All Fields must be filled. 
                            </B></FONT></DIV>
                        <FORM name=form2 action=ppbuy2.asp method=post onSubmit="return checkform(this);">
					<%     
						SET recSearch=Server.CreateObject("ADODB.RecordSet")
				        recSearch.ActiveConnection = Conn
						recSearch.CursorLocation = adUseClient
						recSearch.CursorType = adOpenDynamic
						recSearch.LockType = adLockOptimistic
				      	strSql="Select distinct(description) from product where event_date > " & prepdate(now) & " and status = 'Active' and numticket > '1'"
				        recsearch.open strsql 
				        strpid=request.form("pidi")   ' cache the pid for select to work
					%>


                            <TABLE cellSpacing=3 cellPadding=3 width="100%" 
                        align=center bgColor=#00ccff border=0>
                              <TBODY> 
                              <TR> 
                                <TD vAlign=top width="42%"> 
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
                                color=#ffffff size=2>Type of Ticket</FONT></DIV>
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
                                </TD>
                                <TD vAlign=top width="53%" bgColor=#00ccff><FONT 
                              face="Arial, Helvetica, sans-serif" 
                              size=-1> 
                                  <select name="pidi" onchange="Eventout()">
                                    <option value="" <%if strpid="" then response.write " selected" end if %> > 
                                    Select Type </option>
                                    <% 
									if not recsearch.eof then
									recsearch.movefirst 
									do until recsearch.eof
									%> 
                                    <option value="<%=recsearch(0)%>"
				    				<%if recsearch(0)=cstr(strpid) then response.write " selected"%>> 
                                    <%=recsearch(0)%></option>
                                    <%recsearch.movenext
                    				loop
									end if
                   					recsearch.close %> 
                                  </select>
                                  </FONT></TD>
                                <INPUT type=hidden value="Select a type of ticket" name=pidi_req>
                              </TR>
                              <tr> 
                                <td valign=top width="42%"> 
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
                                color=#ffffff size=2>Event Descriptions</font></div>
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
                                </td>
                                <td valign=top width="53%" bgcolor=#00ccff> <% if request.querystring("mode")="calculate" and request.querystring("level")>0 then 
                SET recSearch=Server.CreateObject("ADODB.RecordSet")
		        recSearch.ActiveConnection = Conn
		        recSearch.CursorLocation = adUseClient
	    	    recSearch.CursorType = adOpenkeyset
            	recSearch.LockType = adLockOptimistic
                event2=request.form("pidi")
                k=request.form("event3")
                if event2<>cstr("") or event2=cstr(0) then 
                	response.write "    </font><select name=""event3"" onchange=""Eventout2()"">"%><option value="" > 
                                  Select Event</option> <%
				  
                                  strSql="Select distinct(event2) from product where description='" & strpid & "'" & " and event_date>" & prepdate(now) & " and status = 'Active' and numticket > '1'"
                recsearch.open strsql 
                if recsearch.recordcount>0 then recsearch.movefirst
                do until recsearch.eof
				%><option value="<%=recsearch(0)%>"
				<%
				if recsearch(0) = cstr(k) then response.write " selected" end if
					response.write ">" & recsearch(0) & "</option>"
             	recsearch.movenext
             	loop
             	recsearch.close
				%>
			</select> <%  end if
             end if %> 
                                  <input type=hidden value="You must enter a value for the entertainment" name=event3_req>
                                </td>
                              </tr>
                              <tr> 
                                <td valign=top width="42%"> 
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
                                color="#ffffff" size="2">Date of Event</font><br>
                                          <font face="Arial, Helvetica, sans-serif" 
                                color="#ffffff" size="1">YYYY-MM-DD 24s hours</font></div>
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
                                </td>
                                <td valign=top width="53%" bgcolor=#00ccff> <% if request.querystring("mode")="calculate" and request.querystring("level")>1 then 
                SET recSearch=Server.CreateObject("ADODB.RecordSet")
		        recSearch.ActiveConnection = Conn
		        recSearch.CursorLocation = adUseClient
	    	    recSearch.CursorType = adOpenkeyset
            	recSearch.LockType = adLockOptimistic
                findate=request.form("event3")
                kk=request.form("detval")
                 
                if event2<>cstr("") or event2=cstr(0) then   'or event3<>cstr("") or event3=cstr(0)  then 
                if findate<>cstr("") or findate=cstr(0) then 
                 response.write " </font><select name=""detval"" onchange=Eventout3()>"%><option value="" <% 'if kk="" then response.write " selected"%>> 
                                  Select Date</option> <% 
                                
			 	strSql="Select Event_date from product where description='" & strpid & "'" & " and event2='" & findate & "' and status = 'Active' and numticket > '1'"
                                datenow = cdate(now)
                recsearch.open strsql 
                if recsearch.recordcount>0 then recsearch.movefirst
                do until recsearch.eof
	       			if cdate(recsearch(0)) > datenow then
				%><option value="<%= internationaldate(recsearch(0)) %>" <% if isdate(kk) then
														if cdate(recsearch(0)) = cdate(cstr(kk)) then
															response.write " selected" 
														end if
														end if
					response.write ">" & internationaldate(recsearch(0)) & "</option>"
					end if
        	    	 recsearch.movenext
	             loop
             	recsearch.close %>
				</select> <% end if
		  end if
           end if%> 
                                  <input type=hidden value="You must enter a value for date" name=detval_req>
                                </td>
                              </tr>
                              <tr> 
                                <td valign=top width="42%"> 
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
                                color=#ffffff size=2>Cost of Tickets // Availability 
                                          </font></div>
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
                                </td>
                                <td valign=top width="53%" bgcolor=#00ccff> <%
					if request.querystring("mode")="calculate" and request.querystring("level")>2 and request.form("detval")<>"" then
			           SET recSearch1=Server.CreateObject("ADODB.RecordSet")
			           recSearch1.ActiveConnection = Conn
				   	   recSearch1.CursorLocation = adUseClient
					   recSearch1.CursorType = adOpenDynamic
					   recSearch1.LockType = adLockOptimistic
			           eventdate=request.form("detval")
                                   
                                    strSql="Select Unit_price_adult,Unit_price_child, product_id, numticket, Orgcomments from product where description='"&strpid&"' and event2='"&findate&"' and Event_date=" & prepdate(eventdate)  
			           'strSql="Select Unit_price_adult,Unit_price_child, product_id, numticket, Orgcomments from product where description='"&strpid&"' and event2='"&findate&"' and Event_date={ts '" & eventdate & "'}"
			           recsearch1.open strsql
			           if recsearch1.recordcount>0 then
           					pricestr="Adult = " & recsearch1(0) 
						   Tickets="    // Availability = " & recsearch1(3)
                                                   Policy="    *Policy* = " & recsearch1(4)
				           d= recsearch1(0)
				           c= recsearch1(1)
				           e= recsearch1(2)
				           f= recsearch1(3)
			           else
				           pricestr=""
			           end if
			           recsearch1.close
					   response.write "<script>numticket=" & f & ";</script>"
		           end if 
					'330
					response.write pricestr
					if clng(f) > clng(ReadFile("text/low.htm")) then
						response.write tickets
                                                  
					else
						if (request.querystring("level")>2) and (eventdate <> "") then
						if (cint(f) <= 1) then
							response.write " // Sold Out "
							response.write "<script>soldout = true;</script>"
						else
							response.write tickets
						end if
						' sent email to notify tickets low
						sFrom = "chris@barcodeticket.com"
						sTo = ReadFile("text/cemail.htm")
						sSubject = "Ticket Low - " & request.form("event3")
						message = "Ticket Low Notification at " & now & vbcrlf 
						message = message & "Type of ticket = " & strpid & vbcrlf
						message = message & "Event = " & request.form("event3") & vbcrlf
						message = message & "Date of event = " & eventdate & vbcrlf
						message = message & "Availability = " & f
						sBody = message
						sCc = ""
						sBcc = ""
						iPriority = 1
						SendMail sFrom, sTo, sSubject, sBody, sCc, sBcc, iPriority
						end if
					end if
				%> </td>
                                <input type=hidden value=<%=d%> name=d>
                                <input type=hidden value=<%=c%> name=c>
                                <input type=hidden value=<%=e%> name=e>
                                <input type=hidden value=<%=f%> name=f>
                              </tr>
                              <tr> 
                                <td valign=top width="42%"> 
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
                                color=#ffffff size=2>Ticket Policy</font></div>
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
                                </td>
                                <td valign=top width="53%" bgcolor=#00ccff> <% response.write Policy %> 
                                </td>
                              </tr>
                              <TR> 
                                <TD vAlign=top width="42%"> 
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
                                color=#ffffff size=2> Email </FONT></DIV>
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
                                </TD>
                                <TD vAlign=top width="53%" bgColor=#00ccff> 
                                  <INPUT type="text"
                              size=30 value= "Cust email" name="cust_email">
                                  <INPUT type=hidden value="You must enter your email" name="cust_email_req">
                                </TD>
                              </TR>
                              <TR> 
                                <TD vAlign=top width="42%"> 
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
                                color=#ffffff size=2>Your First Name</FONT></DIV>
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
                                </TD>
                                <TD vAlign=top width="53%" bgColor=#00ccff> 
                                  <INPUT 
                              maxLength=79 size=30 value="Your First Name" name="first_name">
                                  <INPUT type=hidden value="You must enter a value for card name" name="card_name_req">
                                </TD>
                              </TR>
                              <TR> 
                                <TD vAlign=top width="42%"> 
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
                                color=#ffffff size=2>Your Last Name</FONT></DIV>
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
                                </TD>
                                <TD vAlign=top width="53%" bgColor=#00ccff> 
                                  <INPUT 
                              maxLength=79 size=30 value="Your Last Name" name="last_name">

                                </TD>
                              </TR>

                                <tr> 
                                <td valign=top width="42%"> 
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
                                color=#ffffff size=2>Your Address</font></div>
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
                                </td>
                                <td valign=top width="53%" bgcolor=#00ccff> 
                                  <input 
                              size=30 value="Your Address" name="cust_address1">
                                  <INPUT type=hidden value="You must enter your address" name="cust_address1_req">
                                </td>
                              </tr>
                              <tr> 
                                <td valign=top width="42%"> 
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
                                color=#ffffff size=2>City</font></div>
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
                                </td>
                                <td valign=top width="53%" bgcolor=#00ccff> 
                                  <input 
                              size=30 value="Your City" name="cust_city">
                                  <INPUT type=hidden value="You must enter a value for city/town" name="cust_city_req">
                                </td>
                              </tr>
                              <tr> 
                                <td valign=top width="42%"> 
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
                                color=#ffffff size=2>ZipCode</font></div>
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
                                </td>
                                <td valign=top width="53%" bgcolor=#00ccff> 
                                  <input 
                              size=30 value="Your Zip" name="cust_zip">
                                  <INPUT type=hidden value="You must enter a value for zip" "name=cust_zip_req">
                                </td>
                              </tr>
                              <tr> 
                                <td valign=top width="42%"> 
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
                                color=#ffffff size=2>Contact Num</font></div>
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
                                </td>
                                <td valign=top width="53%" bgcolor=#00ccff> 
                                  <input size=30 value="1234567890" name="contact_num">
                              
                                  
                                </td>
                              </tr>
                              <tr> 
                                <td valign=top width="42%"> 
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
                                color=#ffffff size=2>Number of Tickets</font></div>
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
                                </td>
                                <td valign=top width="53%" bgcolor=#00ccff> 
                                  <SELECT NAME="numseatadult" size=1>
                                    <OPTION selected VALUE="1">1 Adult</OPTION>
                                    <OPTION VALUE="2">2 Adults</OPTION>
                                    <OPTION VALUE="3">3 Adults</OPTION>
                                    <OPTION VALUE="4">4 Adults</OPTION>
                                    <OPTION VALUE="5">5 Adults</OPTION>
                                  </SELECT>
                                  <!--<SELECT NAME="numchild" size=1>
                                    <OPTION selected VALUE="0">0 Child</OPTION>
                                    <OPTION VALUE="1">1 Child</OPTION>
                                    <OPTION VALUE="2">2 Children</OPTION>
                                    <OPTION VALUE="3">3 Children</OPTION>
                                    <OPTION VALUE="4">4 Children</OPTION>
                                    <OPTION VALUE="5">5 Children</OPTION>
                                  </SELECT>-->
                                  
                                </td>
                              </tr>
                              <tr> 
                                <td valign=top width="42%"> 
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
                                color=#ffffff size=2>Set Password for Tickets</font></div>
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
                                </td>
                                <td valign=top width="53%" bgcolor=#00ccff> 
                                  <input 
                              maxlength=40 type= password size=30 name=pw11>
                                  <input type=hidden value="You must enter a password" name=pw11_req>
                                </td>
                              </tr>
                              <TR bgColor=#00ccff> 
                                <TD vAlign=top colSpan=2> 
                                  <CENTER>
                                    <P><FONT 
                              face="Arial, Helvetica, sans-serif"> <br>
                                      <INPUT type=submit value="Continue" onClick="MM_validateForm('cust_email','','RisEmail','first_name','','R','cust_address1','','R','cust_city','','R','cust_zip','','R','pw11','','R','last_name','','R');return document.MM_returnValue">
                                      </FONT> </P>
                                  </CENTER>
                                </TD>
                              </TR>
                              </TBODY> 
                            </TABLE>

                          </FORM></TD></TR><!-- end body -->
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
dim level
if strpid <> "" then level=1
if request.form("event3") <> "" then level=2
if request.form("detval") <> "" then level=3

Function SendMail(sFrom, sTo, sSubject, sBody, sCc, sBcc, iPriority)
	dim myCDO
	set myCDO = Server.CreateObject("CDONTS.NewMail")
	if IsObject(myCDO) then
		myCDO.From = sFrom
		myCDO.To = sTo
		myCDO.Subject = sSubject
		myCDO.Body = sBody	
     	myCDO.importance = iPriority	
   		myCDO.Cc = sCc
		myCDO.Bcc = sBcc
		myCDO.Send
 		set myCDO = nothing	
		SendMail = True
	else
		SendMail = False
	end if
end function
%>

<Script language="javascript">
function Eventout()
{
document.form2.method="post";
document.form2.action="ppbuy.asp?mode=calculate&level=1";
document.form2.submit();
}
function Eventout2()
{
document.form2.method="post";
document.form2.action="ppbuy.asp?mode=calculate&level=2";
document.form2.submit();
}
function Eventout3()
{
document.form2.method="post";
document.form2.action="ppbuy.asp?mode=calculate&level=3";
document.form2.submit();
}

if (soldout) alert('This ticket was sold out.');

</script>

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


'Function PrepDate(DateToConvert)
'	y = year(DateToConvert)
'	m = month(DateToConvert)
'	if len(m) = 1 then m = "0" & m
'	d = day(DateToConvert)
'	if len(d) = 1 then d = "0" & d
 '       strDate = "{ts '" & y & "-" & m & "-" & d & " " & FormatDateTime(DateToConvert, 4) & ":00'}"
'	PrepDate = strDate
'end function

Function InternationalDate(DateToConvert)
	y = year(DateToConvert)
	m = month(DateToConvert)
	if len(m) = 1 then m = "0" & m
	d = day(DateToConvert)
	if len(d) = 1 then d = "0" & d
    strDate = y & "-" & m & "-" & d & " " & FormatDateTime(DateToConvert, 4) & ":00"
	InternationalDate = strDate
end function

%>
