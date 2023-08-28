
<%
Dim strItem, rowColor, strDist, strPassword, strZip, strContact, strEmail, strPhone, strComments
Dim objConn, objRS, objCommand, intRecords, strConnect, totGames, myDesc
Response.ContentType = "text/xml"
strItem = Request.QueryString("item")

strConnect = "D10100190-db3" 
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open strConnect
Set objCommand = Server.CreateObject("ADODB.Command")

objCommand.ActiveConnection = objConn
if isNumeric(strItem) then
 objCommand.CommandText = "SELECT * FROM Inventory WHERE Item_Number = " & strItem
else
 objCommand.CommandText = "SELECT * FROM Inventory WHERE Description Like '" & strItem & "%' ORDER BY Description"
end if
objCommand.CommandType = 1
Set objRS = objCommand.Execute 
intRecords = 0
Response.Write "<Items>" & vbNewLine
While Not objRS.EOF
 intRecords = intRecords + 1
 Response.Write vbTab & "<item>" & vbNewLine
 Response.Write vbTab & vbTab & "<item_number>" & objRS("Item_Number") & "</item_number>" & vbNewLine
 Response.Write vbTab & vbTab & "<description>" & Server.HTMLEncode(objRS("Description")) & "</description>" & vbNewLine
 Response.Write vbTab & vbTab & "<size>" & Server.HTMLEncode(objRS("Size")) & "</size>" & vbNewLine
 Response.Write vbTab & vbTab & "<pack>" & objRS("Pack") & "</pack>" & vbNewLine
 Response.Write vbTab & vbTab & "<price1>" & objRS("Price1") & "</price1>" & vbNewLine
 Response.Write vbTab & vbTab & "<onhand>" & objRS("Inventory_Onhand") & "</onhand>" & vbNewLine  
 Response.Write vbTab & "</item>" & vbNewLine
 objRS.MoveNext
Wend
Response.Write "</Items>" & vbNewLine

Set objRS = Nothing
Set objCommand = Nothing 
Set objConn = Nothing
%>

