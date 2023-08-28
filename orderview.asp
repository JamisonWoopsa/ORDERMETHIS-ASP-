<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>View Distributor Orders</title>
</head>

<body>
<h1>View Distributor Orders</h1>
<%
  Dim objConn, strConnect, strDist, strDatabaseFolder
  strDist = Request.QueryString("dist")
  Session("distUser") = strDist
  strDatabaseFolder = "/database/" & Session("distUser") & "/"
  strConnect = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath(strDatabaseFolder & "wsdist_web.mdb") 'This one is for Access 2000/2002  
  Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open strConnect
  Set objCommand = Server.CreateObject("ADODB.Command")
  objCommand.ActiveConnection = objConn
  strSQL = "SELECT * From Order_Header order by order_date DESC"
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
  Response.Write "Date"
  response.write "</TD>"
  Response.write "<TD>"
  Response.Write "Time"
  response.write "</TD>"
  Response.write "<TD>"
  Response.Write "Customer"
  response.write "</TD>"
  Response.write "<TD>"
  Response.Write "Submitted"
  response.write "</TD>"
  Response.write "</TR>"	
  
  While Not rsView.EOF 
   response.write "<TR>"
   Response.Write "<TD>" & strDist & "</TD>"
   Response.Write "<TD align='right'>" & rsView("order_number") & "</TD>"
   Response.Write "<TD>" & rsView("order_date") & "</TD>"
   Response.Write "<TD>" & rsView("order_time") & "</TD>"
   Response.Write "<TD>" & rsView("c_number") & "</TD>"
   Response.Write "<TD>" & rsView("submitted") & "</TD>"
   response.write "</TR>"
   rsView.MoveNext
  Wend
  response.write "</TABLE>"
  Set objConn = Nothing
  Set rsView = Nothing

%>

</body>
</html>
