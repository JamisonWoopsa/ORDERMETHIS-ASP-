

<%
  Set objXML = Server.CreateObject("microsoft.XMLDOM")
  objXML.load(Server.MapPath("update907.xml"))
  if objXML.parseError.errorCode = 0 then

'ITEMS
    response.write "Items" & "<br>"
    Set objNodeList = objXML.getElementsByTagName("Items")
    For Each objChild In objNodeList.Item(0).childNodes
        For each objField in objChild.childNodes
           if objField.NodeName = "itemNumber" then
	        strItem = objField.text
           elseif objField.NodeName = "description" then
            strDesc = objField.text
		   end if
        next
        response.write strItem & " : " & strDesc & "<br>"
    Next
 
 'CUSTOMERS
  response.write "Customers" & "<br>"
    Set objNodeList = objXML.getElementsByTagName("customers")
    For Each objChild In objNodeList.Item(0).childNodes
        For each objField in objChild.childNodes
           if objField.NodeName = "cNumber" then
	        strItem = objField.text
           elseif objField.NodeName = "name" then
            strDesc = objField.text
		   end if
        next
        response.write strItem & " : " & strDesc & "<br>"
    Next	

'CATEGORIES
    response.write "Categories" & "<br>"
    Set objNodeList = objXML.getElementsByTagName("categories")
    For Each objChild In objNodeList.Item(0).childNodes
        For each objField in objChild.childNodes
           if objField.NodeName = "cNumber" then
	        strItem = objField.text
           elseif objField.NodeName = "description" then
            strDesc = objField.text
		   end if
        next
        response.write strItem & " : " & strDesc & "<br>"
    Next	
	
  response.write "Vendors" & "<br>"

'VENDORS
    Set objNodeList = objXML.getElementsByTagName("vendors")
    For Each objChild In objNodeList.Item(0).childNodes
        For each objField in objChild.childNodes
           if objField.NodeName = "vNumber" then
	        strItem = objField.text
           elseif objField.NodeName = "description" then
            strDesc = objField.text
		   end if
        next
        response.write strItem & " : " & strDesc & "<br>"
    Next	

  response.write "Distributor" & "<br>"
'DISTRIBUTOR
    Set objNodeList = objXML.getElementsByTagName("distributor")
    For Each objChild In objNodeList.Item(0).childNodes
       if objChild.NodeName = "name" then
        strItem = objChild.text
       elseif objChild.NodeName = "phone" then
        strDesc = objChild.text
	   end if
    Next	
    response.write strItem & " : " & strDesc & "<br>"
  else
   response.write "<BR>Error: " & objXML.parseError.errorCode
   response.write "<BR>File: " & objXML.parseError.filepos
   response.write "<BR>LineNum: " & objXML.parseError.line
   response.write "<BR>LinePos: " & objXML.parseError.linepos
   response.write "<BR>Reason: " & objXML.parseError.reason
   response.write "<BR>Source: " & objXML.parseError.srcText
  end if
  set objXML = nothing
%>