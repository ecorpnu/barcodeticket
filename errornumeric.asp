<%  
if not IsNumeric(request( "id")) then
response.redirect "error2.asp"
end if

if not IsNumeric(request( "productid")) then
response.redirect "error2.asp"
end if

if not IsNumeric(request( "prodid")) then
response.redirect "error2.asp"
end if



%>