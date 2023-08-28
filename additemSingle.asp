
<!--#include file="jconnection.asp"-->
<%
  Set objCommand = Server.CreateObject("ADODB.Command")
 'add item number w/qty if item found 
 'validate before getting here
 'if ordernumber = 0 then create new order before adding first item
   itemCount = Request.Form("iTotal")
   errorCount = 0
   qtyCount = 0

   'if session("orderNumber") = 0 then 'start new order
   'PAC 12/17/2002 - Also trap for a blank order number being stored in the session state
   if (session("orderNumber") = 0) OR (NOT IsNumeric(session("orderNumber"))) then 'start new order   
     Set rsOrder = Server.CreateObject("ADODB.Recordset")
     
     '12/18/2002 PAC - Set Cursor location to client so I could use the Bookmark property later on
     rsOrder.CursorLocation = 3 'adUseClient
     
     rsOrder.Open "order_header", objConn, 0, 3, 2
     session("lineNumber") = 0
     rsOrder.AddNew
     rsOrder("C_Number") = session("c_number")
	 myDate = Date
     myTime = Time
	 rsOrder("Order_Date") = myDate
     rsOrder("Order_Time") = myTime
     rsOrder.Update
     
     'PAC 12/17/2002 - Added this to get the order number back from the header just created
     rsOrder.Resync 1,2 'adAffectCurrent,adResyncAllValues
     session("orderNumber") = rsOrder("Order_Number")
     
     rsOrder.Close
	 Set rsOrder = nothing
   end if
   
'    Set rsOrderD = Server.CreateObject("ADODB.Recordset")
'    rsOrderD.Open "order_detail", objConn, 0, 3, 2
    strQty = Request.Form("qty")
    myQty = ""
    iCtr = 1
	for i = 1 to len(strQty) + 1
	 myChar = Mid(strQty,i,1)
     if myChar = "," or i = len(strQty) + 1 then 'process item
       strItem = "itemID" & iCtr
 	   myItem = Request.Form(strItem)
	   if Session("distID") = "1001" then
  	    strDesc = "desc30" & iCtr
	   else
  	    strDesc = "desc" & iCtr
	   end if
	   myDesc = Request.Form(strDesc)

  	   strSize = "size" & iCtr
	   mySize = Request.Form(strSize)

   	   strBreak = "CanBreak" & iCtr
	   myBreak = Request.Form(strBreak)

   	   strPack = "pack" & iCtr
	   myPack = Request.Form(strPack)

   	   strCat = "scat" & iCtr
	   myCat = Request.Form(strCat)
 '      msgbox(myCat)
 
   	   strPrice = "price" & iCtr
	   myPrice = Request.Form(strPrice)

 
       myCode = ""
  	   if myBreak = "Y" then
	    strCode = Request.Form("unitFlag")
   	    if ucase(strCode) = "ON" then
    	   myCode = "U" 
           if myPack <> 0 then
		    myPrice = myPrice / myPack
		   end if
		end if
       end if
	   if myQty <> "" then
        session("lineNumber") = session("lineNumber") + 1
        objCommand.ActiveConnection = objConn
        strSQL =  "INSERT INTO order_detail (order_number, line_number, item_number, description, quantity_ordered, item_size, item_pack, delivery_code, sales_category, price) VALUES (" & session("orderNumber") & ", " & session("lineNumber") & ", " & myItem & ", '" & myDesc & "', " & CDbl(myQty) & ", '" & mySize & "', " & Clng(myPack) & ", '" & myCode & "', " & Clng(myCat) & ", " & Ccur(myPrice) & ")"
        objCommand.CommandText = strSQL
        objCommand.CommandType = 1
        objCommand.Execute intNoOfRecords
	   end if
       'clear previous qty increment item count
       myQty = ""
	   iCtr = iCtr + 1
	 elseif myChar = " " then
	  'continue
	 else
	  myQty = myQty & myChar
	 end if  
	Next

    set objCommand = nothing
    Response.Redirect "myOrder.asp?id=0&item=&status=view&itemadd=" & Request.Form("selectCat")

%>