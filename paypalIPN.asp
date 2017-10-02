
<%@LANGUAGE="VBScript"%>
<!--#include file ="_OpenConnection.asp"-->
<!--#INCLUDE FILE="ADOVBS.INC"-->
<% IF NOT Request.Form("item_number") = "" THEN '001 %>

<%
'Declare our variables we will be receiving

Dim str, OrderID, payment_status, Txn_id, payer_email, my_email, amount_paid
Dim objHttp

'Request the variables we declare above from PayPal
str = Request.Form
OrderID = Request.Form("item_number")
payment_status = Request.Form("payment_status")
Txn_id = Request.Form("txn_id")
payer_email = Request.Form("payer_email")
my_email = ReadFile("text/cemail.htm")
amount_paid = Request.Form("payment_gross")

 
' read post from PayPal system and add 'cmd'

str = str & "&cmd=_notify-validate"


' Post back to PayPal system to validate

set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
' set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
' set objHttp = Server.CreateObject("Microsoft.XMLHTTP")
objHttp.open "POST", "https://www.paypal.com/cgi-bin/webscr", false 
objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
objHttp.Send str 


' Check notification validation
if (objHttp.status <> 200 ) then '002
' HTTP error handling
'Now we see if the payment is pending, verified, or denied 
elseif (objHttp.responseText =  "VERIFIED"  OR objHttp.responseText = "PENDING") AND payment_status = "Completed" THEN 
'elseif (objHttp.responseText =  "INVALID"  OR objHttp.responseText = "PENDING") AND payment_status = "Completed" THEN 

'insert special code to make sure it came from paypal.


        Bouchon = CStr(Hex(4466)) & CStr(Hex(345232)) & Cstr(hex(3432)) 

	strSQL2 = "SELECT * FROM preorder" 
	Set rst2 = Server.CreateObject("ADODB.Recordset")
	rst2.Open strSQL2, conn, adOpenStatic, adLockOptimistic, adCmdText 

	rst2.addnew
	rst2("Status")=now()
        rst2("paypaltnx")=Txn_id
	rst2("xsudsjssaas")=Bouchon
	rst2.Update

	' Clean up time
	rst2.Close

     
'It is verified or pending so we process the payment with the code below

'Begin Send an Email to Customer and Vendor to notify them of their purchase and download URL
'Dim objCDO

'Set objCDO = Server.CreateObject("CDONTS.NewMail")
'objCDO.From = my_email
'objCDO.To = payer_email
'objCDO.CC = my_email
'objCDO.Subject = XX
'objCDO.BodyFormat = 0 
'objCDO.MailFormat = 0 
'objCDO.Body = HTML
'objCDO.Body = "<html>" & vbCrLf & "<head>" & vbCrLf & "<title>Purchase Success</title>" & vbCrLf & "<meta http-equiv= Content-Type content= text/html; charset= iso-8859-1>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<body bgcolor= #FFFFFF text= #000000>" & vbCrLf & "<p>Thank you for purchasing Barcodeticketdemo50 </p>" & vbCrLf & " <p>Here is the link to download your copy: http://www.barcodeticket.com/au/</p>" & vbCrLf & "<p>. The serial number is 1234567890. If you need further information, email me anytime at <a href = mailto:" & my_email &">" & my_email & "</a></p>" & vbCrLf & "<p> </p>" & vbCrLf & "<p>Regards,</p>" & vbCrLf & " <p>//Chris Kwan</p>" & "</body>" & vbCrLf & "</html>"
'objCDO.Send()
'Set objCDO = Nothing 

'End Send an Email to Customer and Vendor


'elseif (objHttp.responseText = "VERIFIED" OR objHttp.responseText = "FAILED") then
elseif (objHttp.responseText = "INVALID" OR objHttp.responseText = "FAILED") then


' If we get an Invalid response from PayPal, then the payment is messed up and we notify the customer

'Begin Send a Problem Email to Customer and Vendor, 
Dim FailobjCDO
Set FailobjCDO = Server.CreateObject("CDONTS.NewMail")
FailobjCDO.From = my_email
FailobjCDO.To = payer_email
FailobjCDO.CC = my_email
FailobjCDO.Subject = "Purchase Problem at your paypal account "
FailobjCDO.Body = "Dear Customer,"&chr(10)&"It seems there was a problem while processing your paypal account."&chr(10)&"Check your PayPal account to make sure everything is correct."&chr(10)&"If you HAVE sent payment and still received this email, contact me to resolve the issue. Item Number: "& OrderID & chr(10) &"If you need further help, please contact me at: " & my_email & chr(10)&chr(10)&"Regards,"&chr(10)&"//System Administrator"&chr(10)& now()
FailobjCDO.Send()
Set FailobjCDO = Nothing
'End Send an Email to Customer and Vendor
else 
' error
end if '002
set objHttp = nothing



END IF '001


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


<!--#include file ="_closeConnection.asp"-->
