

<%
  Dim strDID, strCID, strPassword
  strCID = Request("cid")
  strDID = Request("did")
  Session("custID") = strCID
  Session("distID") = strDID
'check for valid distributors
 '907 = Ford Distributing
 '103 = Sabin Wholesale
 '1 = Jamison Computer Demo
 '3015 = Christopher Wholesalers
 '1001 = SPC with special processing (vendors instead of categories)
 'DON THINK I NEED THE FOLLOWING
  'if ucase(Session("distID")) = "JAMISON" or Session("distID") = "907" or Session("distID") = "103" or Session("distID") = "3015" or Session("distID") = "1001" or Session("distID") = "476" or Session("distID") = "2113" then
  ' 'proceed with customer validation
  'else
  ' 'invalid distributor
  ' Response.Redirect "login.asp?SecondTry=True&Dist=False"               ' - must register
  'end if
%>
<!--#include file="jconnection.asp"-->
<%
  strPassword = Request("Password")
  if strCID = "" or strDID = "" then
   Response.Redirect "login.asp?SecondTry=True"
  end if
 
  Dim rsUsers
  set rsUsers = Server.CreateObject("ADODB.Recordset")
  strSQL = "SELECT * FROM Customer WHERE C_Number = " & strCID
  rsUsers.Open strSQL, objConn_In
  If rsUsers.EOF Then                                               ' User not found
      Response.Redirect "login.asp?SecondTry=True"               ' - must register
  Else                                     'One or more users found - check password
    While Not rsUsers.EOF
      If UCase(rsUsers("Web_Password")) = UCase(strPassword) Then     ' password matched
        Dim strName, strValue
        For Each strField in rsUsers.Fields
          strName = strField.Name                       'Customer populate session variables
          strValue = strField.value
          Session(strName) = strValue 
        Next
 
        Session("CategorySelect") = 0  'default in case field not found in distrubtor table
        Dim rsDist
        set rsDist = Server.CreateObject("ADODB.Recordset")
        strSQL = "SELECT * FROM Distributor"
        rsDist.Open strSQL, objConn_In
        For Each strField in rsDist.Fields
          strName = strField.Name                       ' Distributor populate session variables
          strValue = strField.value
          Session(strName) = strValue 
        Next

        Dim rsCat
        set rsCat = Server.CreateObject("ADODB.Recordset")
        strSQL = "SELECT * FROM Sales_Categories Order By Sales_Category"
        rsCat.Open strSQL, objConn_In
        Dim cnt, okCat, cExclude
		cnt = 0
        comboCnt = 0
        Session("custCatCount") = 12  'default
		While Not rsCat.EOF
         if session("CategorySelect") = 1 then 'load selected categories / set catRef for combo cross reference
		  okCat = True
		  Session("catRef0") = 0  'combo index (all categories selection)
		 ' 'verify category_allowed
		  if not rsCat("Category_Allowed") then
		   okCat = false
		  end if
		  
		  'verify customer exclusions
		  Session("cExclude") = rsUsers("Category_Exclusion")
	      Session("custCatCount") = 0
		  if mid(Session("cExclude"), rsCat("Sales_Category"), 1) <> "1" then
 		   okCat = false
		  else
 	       Session("custCatCount") = Session("custCatCount") + 1 
		  end if
		  
		  'set cross reference session variables = where (catRef & combo.index) returns category 
          if okCat then
		   comboCnt = comboCnt + 1
	       strCatRef = "catRef" & comboCnt
		   Session(strCatRef) = rsCat("Sales_Category")
		   cnt = cnt + 1
           strName = "cat" & cnt
		   Session(strName) = rsCat("Category_Desc")
		  end if
	
		 else  'load all categories
		  cnt = cnt + 1
          strName = "cat" & cnt
		  Session(strName) = rsCat("Category_Desc")
		 end if 
'         cnt = cnt + 1
		 rsCat.MoveNext
        Wend
		Session("totalCat") = cnt
		
		Session("blnValidUser") = True
        Session("nextDesc") = ""
		Session("nextItem") = 0
'		Session("ShowPrice") = 1  'debug
		
        Session("orderNumber") = 0

        session("reset100") = 0
        session("orderMessage") = ""

		'*** check for existing order for this customer to resume if non submitted in last 7 days
		'*** if one or more, select on next screen
	    Dim rsResume, resumeOrder, myResp, resumeDate, testDate
        resumeOrder = 0
	    testDate = Date - 7
		set rsResume = Server.CreateObject("ADODB.Recordset")
        strSQL = "SELECT * FROM Order_Header where c_number = " & Session("custID") & " And Not Submitted And order_date >= #" & testDate & "# Order By Order_Date"
        rsResume.Open strSQL, objConn
 		cnt = 0
        While Not rsResume.EOF
         cnt = cnt + 1
	     resumeOrder = rsResume("Order_Number")
	     resumeDate = rsResume("Order_Date")
		 rsResume.MoveNext
        Wend
        Session("orderNumberResume") = resumeOrder
        Session("orderDateResume") = resumeDate '
       ' ***	
        set rsResume = nothing
		
        Response.Cookies("wsLogin") ("did") = Session("distID")
        Response.Cookies("wsLogin") ("cid") = Session("custID")
        Response.Cookies("wsLogin").Expires = Date + 90
        if cnt = 0 then
		 Response.Redirect "myOrder.asp?id=-1&item=-1" ''&resumeOrder=" & resumeOrder   ' successful login
        else
		 Response.Redirect "myOrderResume.asp"   ' successful login
		end if
	  Else 
        rsUsers.MoveNext
      End If
    Wend
    Response.Redirect "login.asp?SecondTry=True&WrongPW=True" ' Username right; password wrong
  End If
%>
