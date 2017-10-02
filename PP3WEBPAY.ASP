<!--#include file ="_OpenConnection.asp"-->
<!--#include file ="adovbs.inc"-->
<%
'To make sure user comes from paypal
WhereFrom=request.ServerVariables("HTTP_Referer")

If whereFrom="" or whereFrom="https://www.paypal.com/cgi-bin/webscr?__track=_xclick-flow:p/xcl/pay/buy-confirm:_xclick-payment-confirm-submit" then

else
	     'catch illegal loggers
             ForReading = 1
	     ForWriting = 2
	     ForAppending = 8
             Filebase = Server.MapPath(".")
             Set fso = Server.CreateObject("Scripting.FilesystemObject")
	     loggerfile = "logger.log"	
             Set f = fso.OpenTextFile(Filebase & "\" & loggerfile, ForAppending, true)
	        
             f.writeline request.ServerVariables("HTTP_Referer")
            
             f.writeline Now()
	     f.close
    	     Set f = nothing
	    
response.redirect("ppbuy.asp")
end if


'To make sure transaction is genuine.

txn_id = request.form("txn_id")
if request.form("txn_id") <>"" then
IPAddress= Request.ServerVariables("REMOTE_ADDR") 

Bouchon = CStr(Hex(4466)) & CStr(Hex(345232)) & Cstr(hex(3432)) 
sessionx= Bouchon

Set rsx = Server.CreateObject("ADODB.Recordset")
	 
 strsqlx = "select * from preorder where xsudsjssaas='" & sessionx & "' AND paypaltnx ='" & txn_id & "'"
        'response.write strsqlx
       rsx.Open strSQLx, conn, adOpenStatic, adLockOptimistic, adCmdText 
                 
       if not rsx.eof then

       rsx("L441242782L")=IPAddress
       
       rsx.update
       else

       response.redirect("ppbuy.asp")
       rsx.close
       set rsx = nothing
 	end if

else
	response.redirect("ppbuy.asp")

end if

%>



<%
if Request.Cookies("demopp") <> "true" then
response.redirect("ppbuy.asp")
end if
%>


<%
'constant declaration
emailfile = "email.log"
csendform = "somevalue"
logfile = "trans.log"
formfile = "formfile.log"
describe = "ticket"
%>


<%
'THESE ARE RETURNING VARS FROM PAYPAL..


cust_name1=request.form("first_name")
cust_name2=request.form("last_name")
card_name=cust_name1+cust_name2
'response.write card_name
event3=Request.Form("option_name1")
detval=Request.Form("option_selection1")
pw11 = Request.Form("custom")
productid = request.form("ITEM_NUMBER")
numseatadult = request.form("option_selection2")
nn=request.form("option_name2")

amount = request.form("payment_gross")
'card_name = request.form("last_name")
cust_address1 = request.form("address_street")
cust_city = request.form("address_city")
cust_zip = request.form("address_zip")
cust_country = request.form("address_country")
cust_email = request.form("payer_email")
payment_status = Request.Form("payment_status")


strsql = "select * from product where product_id=" & cint(productid) & " "
set rs = conn.execute(strsql)
while not rs.eof
	adultcost = rs("unit_price_adult")
	description = rs("description")
        numticket=rs("numticket")
	rs.movenext
wend
fnum=numticket-numseatadult
set rs = nothing
%>

<%
browser = ""
if instr(1,request.servervariables("HTTP_USER_AGENT"),"IE",1)<>0 then
	browser = "ie"
end if
%>

<% 
'To make sure the whole page is process before output
response.buffer= true 
%>

<%
'No cache
'response.Expires = 0
'response.ExpiresAbsolute = Now() - 1
'response.addHeader "pragma","no-cache"
'response.addHeader "cache-control","private"
'Response.CacheControl = "no-cache" 
%>

<%
Response.Buffer = True
'to prevent loading twice
if Request.Cookies("salescompleted")= "true" then
	Response.Cookies("salescompleted") = "false"
	response.redirect "ppbuy.asp"
else
'to log visitors *********
'prelog
end if
%>

<HTML>
<HEAD>
<title>PayPal Process Page</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">


<style type="text/css">
<!--
BODY, TABLE, LI 
{  font: 12px Verdana }
.f1
{  font: 14px Verdana }
-->
</style>

</HEAD>


<BODY bgcolor="white">
<% 'VBScript Main 

conn.begintrans                            
if numseatadult > 0 and Request.Form("payment_status")="Completed" then

		   nadult = numseatadult
	           for nn = 1 to cint(numseatadult)
			   strMyKey=""  
	           set rs1 = nothing
 	           set rst2 = nothing         
     	           set rst5 = nothing
			set objPrint = server.createobject("pp3.releaseprint")
			objprint.dsn = dsn
			keygen = objprint.generaterandom
			set objprint = nothing
   	       	   kk=keyGen
			   
			   if nadult > 0 then
			   		nadult = nadult - 1
					tt = "Adult"
					tp = adultcost
			   end if
        	   sale tt, tp, nn, numseatadult			'uploads all the information to credit card 
    	       'prints the ticket details

		' select image SUGGEST YOU USE THIS TO EXTRACT DATA WHERE REQUIRED...
		if productID<>"" then
			strSQL7 = "Select imagepath FROM product WHERE product_id=" & productID
			set rs7=conn.execute(strSQL7)
			pathstr = ""
			If len(rs7("imagepath")) > 5 Then        
				pathstr=rs7("imagepath")
			else
        		pathstr=ReadFile("text/ti.htm")
			end if
			rs7.close
		end if 

			   'Response.write "<p>"
			   Response.Write "<font face=verdana size=2>Order ID."
			   Response.Write kk
			   set objPrint = server.createobject("pp3.releaseprint")
			   objPrint.PrintTicket kk, tt, tp, nn, numseatadult, ReadFile("text/ti2.htm"), pathstr, amount, card_name, description, event3, detval
			   set objPrint = nothing
	           space ' provide space between tickets
			   if browser="ie" then
			   		spaceie nn
				else
					spacenets nn
			   end if
     	   next       
else Response.Write "Some errors have occured please contact webmaster.<p>"
     response.write ReadFile("text/cemail.htm")
end if
           
if numseatadult > 0 and Request.Form("payment_status")="Completed" then           
   orderadult
   resetTicket
end if

conn.committrans


' End of Main Script 


' Send the Input form
sub SendForm()
   
   ' On Error Resume Next 
   Set fso = Server.CreateObject("Scripting.FileSystemObject")
   ' Check to see if the file exists.
   if fso.FileExists(FormFile) then
     Set f = fso.GetFile(FormFile)
     Set fileHTML = f.OpenAsTextStream(1,-2)
     ' Strip off everything before the BODY TAG,
     ' then output the File to the client.
     Body = False 
     'Loop until the end of the file
     while Not fileHTML.AtEndOfStream
        if Instr(1,sendHTML, "<BODY", 1) <> 0 then 
            ' If this line has html after the close of the body tag
            ' then we will dismiss it. If this is the case
            ' you would have to write some parsing code
            Body = True
        end if
        if Instr(1, sendHTML, "</BODY", 1) <> 0 then body = False
        if Body then Response.Write(sendHTML) 
        sendHTML = fileHTML.ReadLine
     Wend
     ' Close the File
     fileHTML.close 
   else
     Failed("File Not Found")
   end if
end sub

  
' Send the Response page


'Sub Sale, start here
sub sale(tt, tp, nn, nt)
	Const adOpenStatic = 3
	Const adLockOptimistic = 3
	Const adCmdText = 1
	'update
	strSQL2 = "SELECT * FROM card" 
	Set rst2 = Server.CreateObject("ADODB.Recordset")
	rst2.Open strSQL2, conn, adOpenStatic, adLockOptimistic, adCmdText 

	rst2.addnew
	rst2("Status")=now()
        
	rst2("emailuser")=cust_email
	rst2("card_name")=card_name
	rst2("cust_address1")=cust_address1
	rst2("cust_city")=cust_city
	rst2("cust_country")=cust_country
	rst2("numseat")=numseatadult
	rst2("first_name")=request.form("first_name")
	rst2("cust_zip")=cust_zip
	rst2("last_name")=request.form("last_name")
	rst2("sedid")= left(keyGen,12)
	rst2("des")=event3
	rst2("pw1")=pw11
	rst2("date1")=cdate(detval)
	rst2("product_id")=productid
	rst2("type_ticket")=description
	rst2("numseatadult")=numseatadult
        rst2("numchild")=0
	rst2("dollars")=amount
	rst2("childcost")=0
	rst2("adultcost")=adultcost
	rst2("refnum")=txn_id
        rst2("authcode")=" "
        rst2("amountender")=" "
        rst2("contact_num")=" "
        rst2("verify")="No"
	rst2("buyfor") = tt
	rst2.Update

	If Err.Number = 0 Then
    	Response.Write  "<font face=verdana size=2> Thank you. An receipt email is being sent to you now ."  
	
        Else
    	Response.Write  "<font face=verdana size=2> <center> An error has occured. Contact webmaster.</center>"
	End If

	' Clean up time
	rst2.Close
	
%>

	 Print this..</font>

<%
	' put email address in email.log

	'ForReading = 1
	'ForWriting = 2
	'ForAppending = 8

	'Where to save the email and output log file relative to web server root

	'Filebase = Server.MapPath(".")

	'if cust_email <> "" then
		'Set fso = Server.CreateObject("Scripting.FilesystemObject")
		'if emailfile <> "" then
        	'Set f = fso.OpenTextFile(Filebase & "\" & emailfile, ForAppending, true)
	        'f.writeline cust_email
            'f.writeline card_name
            'f.writeline amount
            'f.writeline Now()
	     '   f.close
    	    'Set f = nothing
	    'end if
	'end if

	'sent request to purchaser now 
strSQL=" Select * FROM card Where card_name='" & card_name & "' AND sedid = '" & left(keyGen,12) & "'"

	
	set rs1 = Server.CreateObject("ADODB.Recordset")
	set rs1 = conn.execute(strSQL)

	if cust_email <> "" then
    	ansmsg = " Thank you " & card_name & "." & vbCrlf _
			 & " Event: " & event3 & vbCrlf _
                         & " Time: " & cdate(detval) & "  for " & nn & " of " & nt & " " & vbCrlf _
              		 & " Your password   " & pw11 & vbCrlf _
                         & " Please click to print your receipt if you have not done so" & vbCrlf _
			 & " " & ReadFile("text/pathprint.htm") & "/ticpt.asp?id=" & rs1("card_id") & "&productID=" & productID & vbCrlf _
                         & " If you have any queries or problems, please write to us at " & vbCrlf _   
    		         & " " &  vbCrlf _ 
	                 & " " & ReadFile("text/cemail.htm") & vbCrlf _
                         & " " & NOW() & vbCrlf 
'252

		        sFrom = ReadFile("text/cemail.htm")
    			sTo = cust_email
    			sBody = ansmsg
    			sCc = ""
    			sBcc = ReadFile("text/cebcc.htm") 
    			iPriority = 1
    			sSubject = "Thank you for purchasing  " & description & " for " & event3
    			SendMail sFrom, sTo, sSubject, sBody, sCc, sBcc, iPriority 



	else
        Response.write "no email was sent"
        end if

	'close rs1
	rs1.close
	

    
end sub


function  SendMail (sFrom, sTo, sSubject, sBody, sCc, sBcc, iPriority)
    dim myCDO
    'set myCDO = Server.CreateObject("CDONTS.NewMail")
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


sub Failed(Message)
    ' We Got an error.
    ' You should put code in here that controls
    ' What to do if you get an error
	Response.Write("<CENTER><H1>"+Message+"</H1></CENTER>")
    Response.Write("Email Us NOW with your Name")
end sub


'to prevent loading twice
Response.Cookies("salescompleted") = "true"
%>      

<%
sub orderadult()
	Const adOpenStatic = 3
	Const adLockOptimistic = 3
	Const adCmdText = 1
	if numseatadult >= 1 then
		for n=1 to numseatadult
			strSQL3 = "orders" 
			Set rst3 = Server.CreateObject("ADODB.Recordset")
			rst3.Open strSQL3, conn, adOpenDynamic, adLockOptimistic, adCmdTable 
			rst3.addnew
			rst3("card_name")=card_name
			rst3("cust_email")=cust_email
			
			rst3("Order_date")=cdate(now)
                        rst3("event_date")=cdate(detval)
			rst3("Description")=description
                        'if request("comments") = "" then
				'rst3("Comment_client") = " "
			'else rst3("Comment_client")=request("comments")
			'end if 
			rst3("product_id")=productid
			rst3("Ticket_cat")="Adult"
                        rst3("costprice")=adultcost
                        rst3("authcode")=" "
			rst3("refnum")=txn_id
                        rst3("admin")=" " ' log the admin
                        rst3("comment_client")=" "
			rst3.Update
			rst3.close
			set rst3=nothing

		next
	end if
end sub

%>





<%
sub resetTicket()
	if productid <> "" then
        
        Set rst5 = Server.CreateObject("ADODB.Recordset")
	 
	strsql5 = "select * from product where product_id=" & cint(productid) & " "
        rst5.Open strSQL5, conn, adOpenStatic, adLockOptimistic, adCmdText 
        rst5("numticket")=fnum 
	rst5.update
        rst5.close
        set rst5=nothing 
	end if
end sub
%>

<%
sub space()
	'To align the tickets when it is printed
	response.write "<font face=verdana size=2>To email us : " & ReadFile("text/cemail.htm") & "</font><br>"
        response.write "<br>"
end sub
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


sub spaceie(nn)
	if (nn=2) then
		response.write "<br><br><br><br><br><br><br>"
		response.write "<br><br><br><br><br><br>"
	end if
	if (nn=4) then
		response.write "<br><br><br><br><br><br><br>"
		response.write "<br><br><br><br><br><br>"
	end if
	if (nn=6) then
		response.write "<br><br><br><br><br><br><br>"
		response.write "<br><br><br><br><br><br>"
	end if
	if (nn=8) then
		response.write "<br><br><br><br><br><br><br>"
		response.write "<br><br><br><br><br><br>"
	end if
end sub

sub spacenets(nn)
	if (nn=2) then
		response.write "<br><br><br><br><br><br><br><br>"
		response.write "<br><br><br><br><br><br><br><br><br><br>"
	end if
	if (nn=4) then
		response.write "<br><br><br><br><br><br><br><br>"
		response.write "<br><br><br><br><br><br><br><br><br><br>"
	end if
	if (nn=6) then
		response.write "<br><br><br><br><br><br><br><br>"
		response.write "<br><br><br><br><br><br><br><br><br><br><br>"
	end if
	if (nn=8) then
		response.write "<br><br><br><br><br><br><br><br>"
		response.write "<br><br><br><br><br><br><br><br><br><br><br><br>"
	end if
end sub
%>

<%
sub prelog
ForReading = 1
ForWriting = 2
ForAppending = 8
'Where to save preliminary info, in case something happen along the way
Filebase = Server.MapPath(".")
if cust_email <> "" then
	FileName2 = "visitor.log"
	Set fso = Server.CreateObject("Scripting.FilesystemObject")
    if emailfile <> "" then
        Set f = fso.OpenTextFile(Filebase & "\" & Filename2, ForAppending, true)
        f.writeline cust_email
        f.writeline card_name
        f.writeline Now()
        f.close
        Set f = nothing
    end if
end if

if cust_email <> "" then
	ansmsg = " Customer Name " & card_name & " ." & vbCrlf _
            & " Address " & cust_address1 & vbCrlf _
			& " City " & cust_city & vbCrlf _
			& " Zip " & cust_zip & vbCrlf _ 
			& " " &  vbCrlf _ 
          	& " " & NOW() & vbCrlf 
'& " State " & cust_state & vbCrlf _
'& " Country " & cust_country & vbCrlf _ 
    sFrom = "Prelog" 
    sTo = ReadFile("text/cemail.htm")
    sBody = ansmsg
    sCc = ""
    sBcc = ""
    iPriority = 1
    sSubject = "You have a sale for barcodeticket -admin - prelog "
    SendMail sFrom, sTo, sSubject, sBody, sCc, sBcc, iPriority 
else
    'Response.write "no email was sent"
end if
end sub
%>


<% if cust_email <> "" then %>
<a href="ppbuy.asp">Buy Another Ticket</a>
<% end if %>
<% Session.abandon %>
<p>

<!--#include file ="_CloseConnection.asp"-->


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

</BODY>
</HTML>


