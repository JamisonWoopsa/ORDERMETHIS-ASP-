
<%
Dim strSQL
Dim objConn, objRS, objCommand, strConnect
Response.ContentType = "text/xml"
strSQL = Request.QueryString("sql")
'strSQL = "Update Inventory Set Price1 = 13 where item_number = 27"
strConnect = "D10100190-db3" 
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open strConnect
Set objCommand = Server.CreateObject("ADODB.Command")

objCommand.ActiveConnection = objConn
objCommand.CommandText = server.htmlencode(strSQL)
objCommand.CommandType = 1
objCommand.Execute 
Response.Write "<Status>Success</Status>"
Set objCommand = Nothing 
Set objConn = Nothing


%>

