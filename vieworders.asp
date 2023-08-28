<html>
<head>
<title>Enter Order</title>
</head>
<body background="img/bg.gif">
<%
 Session("distID") = "1"
%>
<!--#include file="jconnection.asp"-->
<center><h2 STYLE="{color : gold}" ALIGN="Center">Welcome to Internet Order Entry</h2></center>
<%
 Function submitVerify()
   response.write "alert('OK to submit this order ?')"
 End Function

  Set objCommand = Server.CreateObject("ADODB.Command")
  objCommand.ActiveConnection = objConn
  'strSQL = "SELECT * From Order_Header order by order_number"
  strSQL = "SELECT Order_Number, Order_Header.C_Number, Order_Date, Order_Time, submitted, message, C_Name " _
         & "FROM Order_Header INNER JOIN Customer ON Order_Header.C_Number = Customer.C_Number Order By Order_date, order_time, Order_number"
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
   response.write "<TD>"
   Response.Write "Date: " & rsView("order_date")
   response.write "</TD>"
   response.write "<TD>"
   Response.Write "Time: " & rsView("order_time")
   response.write "</TD>" 
   response.write "<TD>"
   Response.Write rsView("c_number") & ":" & rsView("c_name")
   response.write "</TD>"
   response.write "<TD>"
   Response.Write "Submitted: " & rsView("submitted")
   response.write "</TD>"
   response.write "<TD>"
   Response.Write rsView("message")
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
