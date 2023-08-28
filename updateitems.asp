<!--#include file="jconnection.asp"-->
<%

Dim line_number()
Dim line_qty()
Dim line_delete()
Dim orig_qty()
aLength = clng(Request.Form("iTotal")) - 1
ReDim line_number(aLength)
ReDim line_qty(aLength)
ReDim line_delete(aLength)
ReDim orig_qty(aLength)

'validate before getting here
'Set up arrays for line number, new qty, delete flag, original qty
'Step thru array - if delete or new qty <> original qty, update the database

'itemCount = Request.Form("iTotal")
'errorCount = 0
'qtyCount = 0

'LOAD LINE NUMBERS
strLine = Request.Form("uptLine")
myLine = ""
iCtr = 0
for i = 1 to len(strLine) + 1 'read thru line data
 myChar = Mid(strLine,i,1)
 if myChar = "," or i = len(strLine) + 1 then 'process item
   'end of current element
   line_number(iCtr) = myLine
    'clear previous line increment item count
   myLine = ""
   iCtr = iCtr + 1
 elseif myChar = " " then
	  'continue
 else
	  myLine = myLine & myChar
 end if  
Next

'LOAD QTY
strLine = Request.Form("qty")
myLine = ""
iCtr = 0
for i = 1 to len(strLine) + 1 'read thru line data
 myChar = Mid(strLine,i,1)
 if myChar = "," or i = len(strLine) + 1 then 'process item
   'end of current element
   if trim(myLine) = "" then
    myLine = "0"
   end if	
   line_qty(iCtr) = myLine
    'clear previous line increment item count
   myLine = ""
   iCtr = iCtr + 1
 elseif myChar = " " then
	  'continue
 else
	  myLine = myLine & myChar
 end if  
Next

'LOAD DELETE
strLine = Request.Form("uptDelete")
myLine = ""
iCtr = 0
for i = 1 to len(strLine) + 1 'read thru line data
 myChar = Mid(strLine,i,1)
 if myChar = "," or i = len(strLine) + 1 then 'process item
   'end of current element
   line_delete(iCtr) = myLine
    'clear previous line increment item count
   myLine = ""
   iCtr = iCtr + 1
 elseif myChar = " " then
	  'continue
 else
	  myLine = myLine & myChar
 end if  
Next

'LOAD ORIG QTY
strLine = Request.Form("origQty")
myLine = ""
iCtr = 0
for i = 1 to len(strLine) + 1 'read thru line data
 myChar = Mid(strLine,i,1)
 if myChar = "," or i = len(strLine) + 1 then 'process item
   'end of current element
   orig_qty(iCtr) = myLine
    'clear previous line increment item count
   myLine = ""
   iCtr = iCtr + 1
 elseif myChar = " " then
	  'continue
 else
	  myLine = myLine & myChar
 end if  
Next

'DEBUG
'for i = 0 to aLength
' if line_qty(i) <> orig_qty(i) then
'  response.write "QTY CHANGE LINE " & line_number(i) & ":" & line_qty(i) & ":" & line_delete(i) & ":" & orig_qty(i) & "<BR>"
' elseif line_delete(i) = "Y" then
'  response.write "DELETE LINE " & line_number(i) & ":" & line_qty(i) & ":" & line_delete(i) & ":" & orig_qty(i) & "<BR>"
' end if
'Next
'END DEBUG

'UPDATE OR DELETE
Set objCommand = Server.CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConn

for i = 0 to aLength
 if line_delete(i) = "Y" then  'delete line
  objCommand.CommandText = "DELETE * FROM Order_Detail WHERE order_number=" & session("OrderNumber") & " AND line_number =" & line_number(i)
  objCommand.CommandType = 1
  objCommand.Execute 
 elseif line_qty(i) <> orig_qty(i) then 'change line
  objCommand.CommandText = "UPDATE Order_Detail SET quantity_ordered=" & line_qty(i) & " WHERE order_number=" & session("OrderNumber") & " AND line_number =" & line_number(i)
  objCommand.CommandType = 1
  objCommand.Execute 
 end if
Next
Set objCommand = Nothing 
Set objConn = Nothing

Response.Redirect "myOrder.asp?id=0&item=&status=view"
 
%>