<% 

Response.write "There is a problem with your request. " 
Response.write "<br>"
Response.write "Please make sure you got the URL right and if there is any further problem please contact your administrator."
Response.write "<br>"
Response.write " We have logged this error on  "
Response.write Now()
Response.write "<br>"
response.write Request.ServerVariables("REMOTE_HOST")
%>