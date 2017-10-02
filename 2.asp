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
  <b>Caution!!! By deleting expired records, your database will be more efficient.</b>
  <blockquote>
  <b><%=errorMSG %></b>
  </blockquote>
  <form method="post" action="<%= backpage %>">
  
  <input type="submit" value=" Clean Now ">
  <p>
  <b>or
   <p>    <a href= "main.asp" >To return to main menu</a> </b>
  </form>
  </td></tr>
  </table>
  </center>

  <%
  Response.End
END SUB

%>