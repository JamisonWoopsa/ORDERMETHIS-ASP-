<html>
<head>
<title>Enter Order</title>
</head>
<body background="img/bg.gif">
<!--#include file="jconnection.asp"-->
<center><h2 STYLE="{color : gold}" ALIGN="Center">Welcome to Internet Order Entry</h2></center>
<%
 Function submitVerify()
   response.write "alert('OK to submit this order ?')"
 End Function

  Set objCommand = Server.CreateObject("ADODB.Command")
  objCommand.ActiveConnection = objConn
  strSQL = "SELECT * From Order_Detail order by order_number"
  objCommand.CommandText = strSQL
  objCommand.CommandType = 1
  Set rsView = objCommand.Execute 
  response.write "<TABLE BORDER='1' align='center'>"
  rowColor = 0
  intRecords = 0
  While Not rsView.EOF 
   intRecords = intRecords + 1
   If rowColor = 0 Then
    Response.Write "<TR BGCOLOR='FFFFC0'>" 
   rowColor = 1
   Else
    Response.Write "<TR BGCOLOR='CDD9F7'>" 
    rowColor = 0
   End if 
   response.write "<TD align='right'>"
   Response.Write intRecords
   response.write "</TD>"
   response.write "<TD align='right'>"
   Response.Write rsView("order_number")
   response.write "</TD>"
'   response.write "<TD align='left'>"
'   Response.Write rsView("size")
'   response.write "</TD>"
   response.write "<TD>"
   Response.Write rsView("item_number")
   response.write "</TD>"
   response.write "<TD>"
   Response.Write "Order QTY:"
   response.write "</TD>"
   response.write "<TD>"
   Response.Write rsView("quantity_ordered")
   response.write "</TD>" 
   response.write "</TR>"   
   rsView.MoveNext
  Wend
  response.write "</TABLE>"
  set rsView = nothing


Set objCommand = Nothing 
Set objConn = Nothing
%>

</body>

</html>
