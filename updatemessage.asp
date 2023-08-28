
<%
 session("orderMessage") = left(Request("message"), 50)
 if len(Request("message")) > 50 then
  Response.Redirect "myOrder.asp?id=0&item=&status=message"
 else
  Response.Redirect "myOrder.asp?id=0&item=&status=view"
 end if 
 %>
