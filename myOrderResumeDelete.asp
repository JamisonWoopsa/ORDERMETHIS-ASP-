<!--#include file="jconnection.asp"-->
<%
Const adCmdText = &H0001

'delete order
Set objCommand = Server.CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConn
objCommand.CommandText = "DELETE * FROM Order_Detail WHERE order_number=" & session("orderNumberResume")
objCommand.CommandType = 1
objCommand.Execute 
 
objCommand.CommandText = "DELETE * FROM Order_Header WHERE order_number=" & session("orderNumberResume")
objCommand.CommandType = 1
objCommand.Execute 
 
Set objCommand = Nothing 
Set objConn = Nothing

'redirect
Response.Redirect "login.asp"
%>

