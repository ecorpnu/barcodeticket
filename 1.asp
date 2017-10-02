<%
'''''''''''''''''''''''''''''''''''''''''''''''''
'  ERROR FORM 
'''''''''''''''''''''''''''''''''''''''''''''''''
SUB errorFORM( byVal backpage, byVal errorMSG )
  %>

  <center>
  <p>
  <table width=400 cellpadding=5 cellspacing=0
   bgcolor="lightblue" border=2>
  <tr><td> 
  There was a problem with the information you entered</b>
  <blockquote>
  <b><%=errorMSG %></b>
  </blockquote>
  <form method="post" action="<%= backpage %>">
  
  <input type="submit" value=" << Return ">
  </form>
  </td></tr>
  </table>
  </center>

  <%
  Response.End
END SUB

%>