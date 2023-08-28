<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>View Distributor Orders</title>
</head>

<body>
<h1>View Distributor Order Detail</h1>
<%
  Dim objConn, strConnect, strDist, strDatabaseFolder
  strDist = Request.QueryString("dist")
  strOrder = Request.QueryString("order")
  Session("distUser") = strDist
  strDatabaseFolder = "/database/" & Session("distUser") & "/"
  strConnect = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath(strDatabaseFolder & "wsdist_web.mdb") 'This one is for Access 2000/2002  
  Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open strConnect
  Set objCommand = Server.CreateObject("ADODB.Command")
  objCommand.ActiveConnection = objConn
  strSQL = "SELECT * From Order_Detail where order_number =" & strOrder & " order by line_number"
  objCommand.CommandText = strSQL
  objCommand.CommandType = 1
  Set rsView = objCommand.Execute 
  response.write "<TABLE BORDER='1' align='left' cellpadding='0' cellspacing='0'>"
  response.write "<TR>"
  Response.write "<TD>"
  Response.Write "Distributor"
  response.write "</TD>"

  Response.write "<TD align='right'>"
  Response.Write "WEB Order"
  response.write "</TD>"
  Response.write "<TD>"
  Response.Write "Line Number"
  response.write "</TD>"
  Response.write "<TD>"
  Response.Write "Item Number"
  response.write "</TD>"
  Response.write "<TD>"
  Response.Write "Description"
  response.write "</TD>"
  Response.write "<TD>"
  Response.Write "Quantity"
  response.write "</TD>"
  Response.write "<TD>"
  Response.Write "Code"
  response.write "</TD>"

  Response.write "</TR>"	
  
  While Not rsView.EOF 
   response.write "<TR>"
   Response.Write "<TD>" & strDist & "</TD>"
   Response.Write "<TD align='right'>" & rsView("order_number") & "</TD>"
   Response.Write "<TD>" & rsView("line_number") & "</TD>"
   Response.Write "<TD>" & rsView("item_number") & "</TD>"
   Response.Write "<TD>" & rsView("description") & "</TD>"
   Response.Write "<TD>" & rsView("quantity_ordered") & "</TD>"
   Response.Write "<TD>" & rsView("delivery_code") & "</TD>"
   response.write "</TR>"
   rsView.MoveNext
  Wend
  response.write "</TABLE>"
  Set objConn = Nothing
  Set rsView = Nothing

%>

</body>
</html>
