<html>
<head>
<title>Internet Order Entry for candy, tobacco and grocery distributors</title>
</head>
<body bgcolor=white>
<!--#include file="jconnection.asp"-->
<%
 Const adCmdText = &H0001
 'update order header COMPLETE
 Set objCommand = Server.CreateObject("ADODB.Command")
 objCommand.ActiveConnection = objConn

 strCust1 = session("orderMessage")
 strCust = ""
 'Response.Write "Debug ASC<br>"
 For char = 1 to Len(strCust1)
  myChar = Mid(strCust1, char, 1)
  if Asc(myChar) = 34 or Asc(myChar) = 39 then
   strCust = strCust & "~"
  elseif Asc(myChar) = 13 then ' CR
   strCust = strCust & "."
  elseif Asc(myChar) = 10 then ' LF
   'remove
  else
   strCust = strCust & myChar
  end if
'  response.write Asc(myChar) & "|"
 Next
' Response.Write "<br>"
' For char = 1 to Len(strCust)
'  myChar = Mid(strCust, char, 1)
'  response.write Asc(myChar) & "|"
 'Next
 'Response.Write strCust1
 'Response.Write strCust
 'Response.Write "<br>End Debug ASC<br>"

 'create text file
 mydate = formatdatetime(date(), 2)
 mytime = formatdatetime(time(), 3)
' 're-format date
 strDate = ""
 dTest = 0
 for i = 1 to len(mydate)
  myChar = mid(mydate, i, 1)
  if myChar = "/" then
   dTest = dTest + 1
  else
   if dTest = 0 then 'month
    myMo = myMo & myChar
   elseif dtest = 1 then 'day
    myDa = myDa & myChar
   else
    myYr = myYr & myChar
   end if
  end if
 next
 strDate = myYr
 if len(myMo) < 2 then
  strDate = strDate & "0" & myMo
 else 
  strDate = strDate & myMo 
 end if
 if len(myDa) < 2 then
  strDate = strDate & "0" & myDa
 else 
  strDate = strDate & myDa 
 end if'
' response.write "Date" & strDate
' 're-format time
 strTime = ""
 dTest = 0
 for i = 1 to len(mytime)
  myChar = mid(mytime, i, 1)
  if myChar = ":" then
   dTest = dTest + 1
  else
   if dTest = 0 then 'hour
    myHr = myHr & myChar
   elseif dtest = 1 then 'minute
    myMi = myMi & myChar
   else
    mySe = mySe & myChar
   end if
  end if
 next 

 if len(myHr) < 2 then
  strTime = strTime & "0" & myHr
 else 
  strTime = strTime & myHr 
 end if
 if len(myMi) < 2 then
  strTime = strTime & "0" & myMi
 else 
  strTime = strTime & myMi 
 end if
 if len(mySe) < 2 then
  strTime = strTime & "0" & mySe
 else 
  strTime = strTime & mySe 
 end if
 strTime = Left(strTime, 6) '6 digits only
 'create file name 
 zPad = 6 - len(session("C_Number")) 'zero pad to 6 digits
 sCust = ""
 for i = 1 to zPad
  sCust = sCust & "0" 
 Next
 sCust = sCust & session("C_Number")
 strFile = sCust & ".1." & strDate & "." & strTime & ".WEB" 'distributor now has specific folder
 strFileName = server.MapPath("/orders/" & Session("distUser") & "/" & strFile)
 set objFSO = CreateObject("Scripting.FileSystemObject")
' response.write strFileName
 set objOrderFile = objFSO.CreateTextFile(strFileName)
 'strRecord = "H" & session("C_Number") & "~WEB" & session("orderNumber") & "~" & session("orderMessage")
 strRecord = "H" & session("C_Number") & "~WEB" & session("orderNumber") & "~" & strCust
 objOrderFile.WriteLine strRecord  'write text file header
 
'DO THIS AT THE END
' objCommand.CommandText = "update order_header set submitted = true, message = '" & strCust & "' where order_number=" & session("ordernumber")
' objCommand.CommandType = adCmdtext
' objCommand.Execute intNoOfRecords
' response.write intNoOfRecords
' set objCommand = nothing
'TABLE TO CONTAIN BOTH TABLES
'TABLE FOR HEADER
 response.write "<TABLE width='650' BORDER='0' align='left' cellpadding='0' cellspacing='0'>"
 response.write "<TR><TD>" 'table inside row

 response.write "<TABLE width='650' BORDER='0' align='left' cellpadding='0' cellspacing='0'>"
 response.write "<TR>"
 response.write "<TD width='40%'>Thank you for placing your order with:</TD>"
 Response.write "<TD width='40%'>Account Number: " & session("C_Number") & "</TD>"
 response.write "</TR>"
 response.write "<TR>"
 response.write "<TD><b>" & session("D_Name") & "</b></TD>"
 Response.write "<TD><b>" & session("C_Name") & "</b></TD>"
 response.write "</TR>"
 response.write "<TR>"
 response.write "<TD>" & session("D_Addr1") & "</TD>"
 Response.write "<TD>" & session("C_Address") & "</TD>"
 response.write "</TR>"
 response.write "<TR>"
 response.write "<TD>" & session("D_City") & "  " &  session("D_State") & "  " &  session("D_Zip") & "</TD>"
 Response.write "<TD>" & session("C_City") & "  " &  session("C_State") & "  " &  session("C_Zip") & "</TD>"
 response.write "</TR>"
  response.write "<TR>"
 response.write "<TD>" & session("D_Phone") & "</TD>"
 Response.write "<TD>" & "Order Reference: WEB " & session("orderNumber") & "</TD>"
 response.write "</TR>"

 
 response.write "<TR>" 'blank line
 response.write "<TD>&nbsp</TD><TD>&nbsp</TD>"
 response.write "</TR>"
 response.write "<TR>" 'log off
 response.write "<TD>" & "<A href='logoff.asp'>Click here to Log Off</a>" & "</TD>"
 response.write "</TR>"
 response.write "<TR>" 'blank line
 response.write "<TD>&nbsp</TD><TD>&nbsp</TD>"
 response.write "</TR>"
 
 response.write "</TABLE>"  'end table inside row  
 response.write "</TD></TR>"

' Response.Write "Thank you for placing your order with:<br><br>"
' Response.Write "<b>" & session("D_Name") & "</b><br>"
' Response.Write session("D_Addr1") & "<br>"
' Response.Write session("D_City") & "  " &  session("D_State") & "  " &  session("D_Zip") & "<br>"
' Response.Write session("D_Phone") & "<br>"

' response.write "<BR>You may print this page for your reference.<br><br>"
'   Response.write "<A href='logoff.asp'>Click here to Log Off</a><BR><BR>"
' Response.Write "<BR><BR>Order Reference: WEB " & session("orderNumber") & "<BR><BR>"

' Response.Write "<b>" & session("C_Name") & "<br></b>"
' Response.Write "Account Number: " & session("C_Number") & "<br>"
' Response.Write session("C_Address") & "<br>"
' Response.Write session("C_City") & "  " &  session("C_State") & "  " &  session("C_Zip") & "<br><br>"'

' if session("orderMessage") <> "" then
'   response.write "MESSAGE: <br>"
'   response.write session("orderMessage")
'   Response.Write "<BR><BR>"
' end if
 
 'list items
 'Response.Write "Items ordered:<BR><BR>"
 
  Set objCommand = Server.CreateObject("ADODB.Command")
  objCommand.ActiveConnection = objConn
  'strSQL = "SELECT * From Order_Detail Where Order_number=" & session("orderNumber") & " order by line_number"
  lastCat = -1
  catQty = 0
  catDollars = 0
  totDollars = 0
  strSQL = "SELECT * From Order_Detail Where Order_number=" & session("orderNumber") & " order by sales_category, description"
    objCommand.CommandText = strSQL
  objCommand.CommandType = 1
  Set rsView = objCommand.Execute 
  intRecords = 0

   response.write "<TR><TD>" 'table inside row
   response.write "<TABLE width='650' BORDER='1' align='left' cellpadding='0' cellspacing='0'>"
'  response.write "<TR BGCOLOR='CDD9F7'><TD>&nbsp</td><TD align='center' border='0'><b>Item</b></td><TD align='left' border='0'><b>Size</b></td><TD align='center' border='0'><b>Pack</b></td><TD align='left' border='0'><b>Description</b></td><TD align='left' border='0'><b>Quantity</b></td><TD align='center' border='0'><b>&nbsp&nbsp</b></td></TR>"
   response.write "<TR>"
   Response.write "<TD>"
   Response.Write "Item #"
   response.write "</TD>"
   Response.write "<TD>"
   Response.Write "Description"
   response.write "</TD>"
   Response.write "<TD>"
   Response.Write "Size"
   response.write "</TD>"
   if session("ShowPrice") = 1 then   
    Response.write "<TD align='right'>"
    Response.Write "Price"
    response.write "</TD>"
   end if
   Response.write "<TD align='right'>"
   Response.Write "Quantity"
   response.write "</TD>"
   if session("ShowPrice") = 1 then   
    Response.write "<TD align='right'>"
    Response.Write "Extended"
    response.write "</TD>"
   end if
   Response.write "</TR>"
'**

  While Not rsView.EOF 
   response.write "<TR>"
   if rsView("sales_category") = lastCat then
     'nothing
   else

	if lastCat = -1 then 'just header / no total line
    else
     'TOTAL LINE
     'strName = "cat" & lastCat
 
     catName = ""
	 if session("CategorySelect") = 1 then ' (strCat is index of combo - not necessarily sales_category - catRef + (combo.index) contains category from checkLogin.asp
      'selected categories
      'IF catRef matches category, get matching description
	  for i = 1 to Session("totalCat")
	   strSelect = "catRef" & i
	   selectCat = session(strSelect)
       if lastCat = selectCat then
	    strName = "cat" & i
	    catName = session(strName) 
	   end if
	  next
     else
      'all categories
	  strName = "cat" & lastCat 'rsView("Sales_Category")
	  catName = session(strName)
	 end if
 
     response.write "<TR>"
     Response.write "<TD>&nbsp</TD>"
     Response.write "<TD align='right'>"
     'Response.Write "TOTAL " & session(strName) & " :  "
     Response.Write "TOTAL " & catName & " :  "
	 response.write "</TD>"
     Response.write "<TD>&nbsp</TD>"
     if session("ShowPrice") = 1 then   
      Response.write "<TD>&nbsp</TD>"
     end if
     Response.write "<TD align='right'>"
     Response.Write catQty
     response.write "</TD>"
     if session("ShowPrice") = 1 then   
      Response.write "<TD align='right'>"
      Response.Write formatcurrency(catDollars, 2)
      response.write "</TD>"
     end if
     Response.write "</TR>"
'	 Response.Write "<BR>" & session(strName) & " TOTAL: " & catQty & "<BR>" & "<BR>"
	 catQty = 0
	 catDollars = 0
	end if

    catName = ""
	if session("CategorySelect") = 1 then ' (strCat is index of combo - not necessarily sales_category - catRef + (combo.index) contains category from checkLogin.asp
     'selected categories
     'IF catRef matches category, get matching description
	 for i = 1 to Session("totalCat")
	  strSelect = "catRef" & i
	  selectCat = session(strSelect)
      if rsView("Sales_Category") = selectCat then
	   strName = "cat" & i
	   catName = session(strName) 
	  end if
	 next
    else
     'all categories
	 strName = "cat" & rsView("Sales_Category")
	 catName = session(strName)
	end if
 
 '   strName = "cat" & rsView("Sales_Category")
	'HEADER
'	Response.Write session(strName) & "<BR>" & "<BR>"
     response.write "<TR>"
     Response.write "<TD>&nbsp</TD>"
     Response.write "<TD>"
     Response.Write catName 'session(strName) 
     response.write "</TD>"
     Response.write "<TD>&nbsp</TD>"
     if session("ShowPrice") = 1 then   
      Response.write "<TD>&nbsp</TD>"
      Response.write "<TD>&nbsp</TD>"
     end if

	 Response.write "<TD>&nbsp</TD>"
     Response.write "</TR>"
   end if 

   lastCat = rsView("sales_category")
   uCode = " "
   if rsView("delivery_code") = "U" then
    uCode = "U"
   end if	
'add price to output file
'   strRecord = "D" & rsView("item_number") & "~" & rsView("quantity_ordered") & "~" & uCode & "~0~ "
   strRecord = "D" & rsView("item_number") & "~" & rsView("quantity_ordered") & "~" & uCode & "~" & rsView("price") & "~ "
   objOrderFile.WriteLine strRecord  'write text file line
   intRecords = intRecords + 1
 '  Response.Write rsView("Line_Number") & "   "
   Response.write "<TD>"
   Response.Write rsView("item_number")
   response.write "</TD>"

'   Response.Write rsView("item_number") & "   "
'   Response.Write rsView("Description") & " / "
   Response.write "<TD>"
   Response.Write rsView("Description")
   response.write "</TD>"

    if isnull(rsView("item_size")) or trim(rsView("item_size")) = "" then
     Response.write "<TD align='left'>&nbsp</TD>" 'formatting
    else
     response.write "<TD align='left'>"
     Response.Write rsView("item_size")
     response.write "</TD>"
    end if
'   Response.write "<TD>"
'   Response.Write rsView("item_size")
'   response.write "</TD>"
   if session("ShowPrice") = 1 then   
    Response.write "<TD align='right'>"
    Response.Write formatcurrency(rsView("price"), 2)
    response.write "</TD>"
   end if

'   Response.Write "QTY: " & rsView("quantity_ordered") & " " & uCode & "<BR>"
   catQty = catQty + rsView("quantity_ordered")
   catDollars = catDollars + (rsView("quantity_ordered") * rsView("price"))
   totDollars = totDollars + (rsView("quantity_ordered") * rsView("price"))
   Response.write "<TD align='right'>"
   Response.Write rsView("quantity_ordered")
   response.write "</TD>"
   if session("ShowPrice") = 1 then   
    Response.write "<TD align='right'>"
    Response.Write formatcurrency((rsView("quantity_ordered") * rsView("price")), 2)
    response.write "</TD>"
   end if
   Response.write "</TR>"
   rsView.MoveNext
  Wend
  strRecord = "EOF"
  objOrderFile.WriteLine strRecord  'write text file last line

  'LAST TOTAL LINE
  if intRecords = 0 then

   response.write "<TR>"
   Response.write "<TD>&nbsp</TD>"
   Response.write "<TD>"
   Response.Write "* NO ITEMS WERE ORDERED ... NO ORDER HAS BEEN PLACED *" 
   response.write "</TD>"
   Response.write "<TD>&nbsp</TD>"
   Response.write "<TD>&nbsp</TD>"
   Response.write "</TR>"

  else

   response.write "<TR>"
   Response.write "<TD>&nbsp</TD>"
   Response.write "<TD align='right'>"
   Response.Write "TOTAL " & session(strName) & " :  "
   response.write "</TD>"
   Response.write "<TD>&nbsp</TD>"
   if session("ShowPrice") = 1 then   
    Response.write "<TD>&nbsp</TD>"
   end if
 
   Response.write "<TD align='right'>"
   Response.Write catQty
   response.write "</TD>"
   if session("ShowPrice") = 1 then   
    Response.write "<TD align='right'>"
    Response.Write formatcurrency(catDollars, 2)
    response.write "</TD>"
   end if
   Response.write "</TR>"

   'GRAND TOTAL
   if session("ShowPrice") = 1 then   

    Response.write "<TR>"
    Response.write "<TD colspan='6'>&nbsp</TD>"
    Response.write "</TR>"
	
    Response.write "<TR>"
    Response.write "<TD colspan='2' align='right'>"
    Response.Write "ORDER TOTAL:" 
    response.write "</TD>"
    Response.write "<TD colspan='4' align='right'>"
    Response.Write formatcurrency(totDollars, 2)
    response.write "</TD>"
    Response.write "</TR>"
   end if  
   
   
  end if
  
'  if session("orderMessage") <> "" then
'   response.write "Message: <br>"
'   response.write session("orderMessage")
'  end if
'  Set objCommand = Nothing 
  response.write "</TABLE>"  'end table inside row  
  response.write "</TD></TR>"

 'MESSAGE TABLE
 response.write "<TR><TD>"
 response.write "<TABLE>" 
 response.write "<TR>" 'blank line
 response.write "<TD>&nbsp</TD><TD>&nbsp</TD>"
 response.write "</TR>"
 response.write "<TR>" 'log off
 if trim(session("orderMessage")) = "" then
     Response.write "<TD align='left'>&nbsp</TD>" 'formatting
 else
     response.write "<TD align='left'>"
     response.write "Message: " & session("orderMessage")
     response.write "</TD>"
 end if
 response.write "</TR>" '
 
 if session("ShowPrice") = 1 then   
  response.write "<TR>" 'blank line
  response.write "<TD>&nbsp</TD><TD>&nbsp</TD>"
  response.write "</TR>" 
 
  response.write "<TR>" 'blank line
  response.write "<TD><b>PRICES SUBJECT TO CHANGE WITHOUT NOTICE</b></TD"
  response.write "</TR>"
  
  response.write "<TR>" 'blank line
  response.write "<TD>&nbsp</TD><TD>&nbsp</TD>"
  response.write "</TR>"
 end if	
 response.write "<TR>" 'log off
 response.write "<TD>" & "<A href='logoff.asp'>Click here to Log Off</a>" & "</TD>"
 response.write "</TR>" '
 
 response.write "</TABLE>"  'end table inside row  
 response.write "</TD></TR>"
 
  response.write "</TABLE>"  'end container table

  'UPDATE HEADER : SUBMITTED
  objCommand.CommandText = "update order_header set submitted = true, message = '" & strCust & "' where order_number=" & session("ordernumber")
  objCommand.CommandType = adCmdtext
  objCommand.Execute intNoOfRecords
  set objCommand = nothing

  set rsView = nothing
  Set objConn = Nothing
  set objFSO = Nothing
  set objOrderFile = Nothing
  session.Abandon
'  Response.write "<BR><A href='logoff.asp'>Click here to Log Off</a>"
%>
</BODY>
</HTML>
