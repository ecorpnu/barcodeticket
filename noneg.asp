<!--#include file ="1.asp"-->

<% ' error check
x=Request.form("f")

if x =< "0" or x= "" then
  errorForm "6.asp", "Sorry please try again or ticket sold out."
end if

%>


<% 

'x = request.form("f")

'if x >= "1" then

'else

'response.expires = 0

'response.redirect "default.asp"

'end if

%>