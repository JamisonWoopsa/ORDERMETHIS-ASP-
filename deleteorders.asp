<html>
<head>
<title>Enter Order</title>
</head>
<body background="img/bg.gif">

<center><h2 STYLE="{color : gold}" ALIGN="Center">Orders Deleted</h2></center>
<%
Dim objConn, objRS, objCommand, intNoOfRecords, strConnect

strConnect = "D10100190-db3" 
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open strConnect
Set objCommand = Server.CreateObject("ADODB.Command")
varOrderA = Request.QueryString("sNum")
varOrderB = Request.QueryString("eNum")
objCommand.ActiveConnection = objConn
objCommand.CommandText = "Delete * FROM order_header where order_number >=" & varOrderA & " and order_number <=" & varOrderB
objCommand.CommandType = 1
objCommand.Execute 
objCommand.CommandText = "Delete * FROM order_detail where order_number >=" & varOrderA & " and order_number <=" & varOrderB
objCommand.CommandType = 1
objCommand.Execute 
Set objCommand = Nothing 
Set objConn = Nothing
response.write "Starting order: " & varOrderA & "<br>"
response.write "Ending order: " & varOrderB & "<br>"
%>

</body>

</html>
