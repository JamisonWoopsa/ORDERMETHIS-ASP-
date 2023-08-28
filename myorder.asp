<html>
<head>
<title>Enter Order</title>

<STYLE>
.maintext {color: black; font-family: "ms sans serif", "arial"; font-size: 10pt}
.lineitems {color: black; font-family: "ms sans serif", "arial"; font-size: 10pt}
.inputText {text-align: right;}
</style> 


</head>

<Script language="Javascript">
function setFormFocus()
{
 //tLines = document.InputForm.iTotal.value - 1;
 //tLines = document.InputForm.iTotal.value;
 //alert("tLines=" + tLines + " " + document.InputForm.iTotal.value);
 //if (tLines = 0) {
	//	alert("tLines=" + tLines);
	//	document.forms[0].item.focus();
    //    }
 //else if (tLines = 1) {
//	alert("OneLineA=1");
//		document.forms[1].qty.focus();
//		}
 //else {
//		alert("tLines=" + tLines);
//        document.forms[1].qty[1].value;
	//	document.forms[1].qty1.focus();
//	  }

//NEW
 tLines = document.InputForm.iTotal.value - 1;
 //alert("Set Focus " + tLines);
 if (tLines == 0) // ONE ITEM
 {
	//alert("OneLineB=1");
	document.forms[1].qty.focus();
 }
 else if (tLines == -1) // NO ITEMS
 {
	document.forms[0].item.focus();
 }
 else
 { 
  	//alert("MoreLines");
	document.forms[1].qty[0].focus();
 } 
//NEW
}

</script>

<%
strItem = Request.QueryString("item") 
'onLoad = "document.forms[0].item.focus();"
if trim(strItem) = "" or trim(strItem) = "-1" then
 'set focus on search
 response.write "<body onLoad = document.forms[0].item.focus(); background='img/bg.gif' LINK='#7690B2' ALINK='#7690B2' VLINK='#7690B2'>"
else
 response.write "<body onLoad = setFormFocus(); background='img/bg.gif' LINK='#7690B2' ALINK='#7690B2' VLINK='#7690B2'>"
' response.write "<body  background='img/bg.gif' LINK='#7690B2' ALINK='#7690B2' VLINK='#7690B2'>"
 'response.write "<body  background="img/bg.gif" LINK="#7690B2" ALINK="#7690B2" VLINK="#7690B2">"
end if
%>

<!--#include file="jconnection.asp"-->


<Script language="Javascript">


function LeftS(str, n){
    if (n <= 0)
        return "";
    else if (n > String(str).length)
        return str;
    else
        return String(str).substring(0,n);
}

function RightS(str, n){

    if (n <= 0)
       return "";
    else if (n > String(str).length)
       return str;
    else {
       var iLen = String(str).length;
       return String(str).substring(iLen, iLen - n);
    }
} 
//window.open("http://www.w3schools.com", "_blank", "toolbar=yes, scrollbars=yes, resizable=yes, top=500, left=500, width=400, height=400");
function openHelp() {
 var newWindow = window.open("WebOrderEntryHelp.pdf", "_blank", "toolbar=no, scrollbars=yes, resizable=yes");
 newWindow.focus();
}

function openImage(img) {
 var newWindow = window.open(img, "_blank", "toolbar=no, scrollbars=yes, resizable=yes");
 newWindow.focus();
}

image1 = new Image();
image1.src = "img/vc_on.gif";

image2 = new Image();
image2.src = "img/vc_off.gif";

image3 = new Image();
image3.src = "img/logoff_on.gif";

image4 = new Image();
image4.src = "img/logoff_off.gif";

image5 = new Image();
image5.src = "img/sub_on.gif";

image6 = new Image();
image6.src = "img/sub_off.gif";

image7 = new Image();
image7.src = "img/am_on.gif";

image8 = new Image();
image8.src = "img/am_off.gif";

function swapimage(img_name,img_src) {

  if (img_name == 1){
   if (img_src == 0) document["message"].src="img/am_on.gif";
   else document["message"].src="img/am_off.gif";
  }

  if (img_name == 2){
   if (img_src == 0) document["viewchange"].src="img/vc_on.gif";
   else document["viewchange"].src="img/vc_off.gif";
  }

  if (img_name == 3){
   if (img_src == 0) document["submitorder"].src="img/sub_on.gif";
   else document["submitorder"].src="img/sub_off.gif";
  }
  
  if (img_name == 4){
   if (img_src == 0) document["logoff"].src="img/logoff_on.gif";
   else document["logoff"].src="img/logoff_off.gif";
  }

}


function imgError(image, isrc) {
   // if image not found, use onerror.jpg in user folder
   image.onerror = "";
   image.src = isrc;
   return true;
}

function imgVerify(myReference) {

 var imgSource = document.getElementById(myReference).src; //this works
 var imgTest = imgSource.substring(imgSource.length - 11);
 
 //if right 11 of imgSource = 'onerror.jpg' return false else return true
 
 if (imgTest == "onerror.jpg") { return false; }
   else { return true; }

}

function checkQTY() {
//if only 1 line, no array 
 var badItems, strQty, strCode, strU, strUU, nLength, ictr = 0;
 badItems = 0;
 tLines = document.InputForm.iTotal.value - 1;
 if (tLines == 0) //no array
 {
//   strQty = document.InputForm.qty.value
//   nLength = strQty.length
//   strU = RightS(strQty, 1)
//   strUU = strU.toUpperCase()
//   if (strUU == "U" && nLength > 1) {
// if breakable then set delivery code - additems
//      document.InputForm.qty.value = LeftS(strQty, nLength - 1)
//      document.InputForm.DeliveryCode.value = "U"
//      if (isNaN(document.InputForm.qty.value) || document.InputForm.CanBreak.value == "N") {
//         document.InputForm.DeliveryCode.value = ""
//         alert("Line " + (ictr + 1) + "- Item Cannot Be Broken: "  + document.InputForm.qty.value);
//         badItems = badItems + 1;
//      }
//    }

    if (isNaN(document.InputForm.qty.value)) {
     alert("Line " + (ictr + 1) + "- Invalid quantity: " + document.InputForm.qty.value);
     badItems = badItems + 1;
    }
 }
 else
 { 
  for (var ictr = 0; ictr <= tLines; ++ictr) {
   if (isNaN(document.InputForm.qty[ictr].value)) {
     alert("Line " + (ictr + 1) + "- Invalid quantity: "  + document.InputForm.qty[ictr].value);
     badItems = badItems + 1;
    }
   }
 } 

 if (badItems > 0) 
   return false; 
 else
   return true;
}

function checkQTY2() {
//if only 1 line, no array 
 var badItems, ictr = 0;

 badItems = 0;
 tLines = document.updateForm.iTotal.value - 1;
 if (tLines == 0) //no array
 {
   if (document.updateForm.delFlag.checked == true) {
      document.updateForm.uptDelete.value = 'Y';
     }
    if (isNaN(document.updateForm.qty.value)) {
      alert("Line " + (ictr + 1) + "- Invalid quantity: "  + document.updateForm.qty.value);
      badItems = badItems + 1;
     }
 }
 else
 {
  for (var ictr = 0; ictr <= tLines; ++ictr) {
    if (document.updateForm.delFlag[ictr].checked == true) {
      document.updateForm.uptDelete[ictr].value = 'Y';
     }
    if (isNaN(document.updateForm.qty[ictr].value)) {
      alert("Line " + (ictr + 1) + "- Invalid quantity: "  + document.updateForm.qty[ictr].value);
      badItems = badItems + 1;
     }
   }
 }
 if (badItems > 0) 
   return false; 
 else
   return true;
}

function checkQTY3() {
//check for quantities if 'more items' is clicked don't proceed until added
//if only 1 line, no array 
 var badItems, okContinue, myQty, ictr = 0;
// alert("here");
 badItems = 0;
 okContinue = true;
 tLines = document.InputForm.iTotal.value - 1;
  for (var ictr = 0; ictr <= tLines; ++ictr) {
//   alert(document.InputForm.qty[ictr].value); //debug
   if (isNaN(document.InputForm.qty[ictr].value)) {
     alert("Line " + (ictr + 1) + "- Invalid quantity: "  + document.InputForm.qty[ictr].value);
     badItems = badItems + 1;
    }
   if (document.InputForm.qty[ictr].value != "") {
//     alert("Line " + (ictr + 1) + "- Quantity: "  + document.InputForm.qty[ictr].value);
  	  okContinue = false;
    }

 //valid qty
//   else  {
//     alert(document.InputForm.qty[ictr].value);
//     if document.InputForm.qty[ictr].value <> "0" {
//      alert("No continue");
//	  okContinue = false;
//    }
//    }

   } //end for

   if (okContinue) { return true; }
   else { 
      alert("Quantities have been entered on this screen.\r\rPlease click 'Add items to order' and then 'Resume Next Items' to continue");
      return false;
	  }
}

function submitVerify(){
 return confirm("OK to submit this order ?");
}
function logoffVerify(){
 return confirm("Are you sure that you want to logoff and end this session ?");
}

</script>
<%


'***
Dim strSQL, strSQL2, lastLineNumberResume, myPos, myQuickQty

If Not session("blnValidUser") then
 Response.redirect "login.asp"
else
 Dim strCat, strItem, rowColor, strStatus, newQueryString
 strItem = Request.QueryString("item") 
 'strItem = "99" 'debug
 strItemString = cstr(Request.QueryString("item")) 'convert to string - decode %2B = +
 
 'strip + from strItem if found
 myPos = instr(strItemString, "+")
 myQuickQty = 0
 if myPos <> 0 then
    strItem = left(strItemString, myPos - 1)
    if myPos = len(strItemString) then
	 'set qty = 1 if + is last character
	 strQuickQty = "1"
	else
	 strQuickQty = right(strItemString, len(strItemString) - myPos)
	end if
    if isnumeric(strQuickQty) then
	 myQuickQty = strQuickQty
	end if
 end if

 'set variable to process addItemSingle.asp
' session("quickQty") = myQuickQty
 
 strCat = Request.QueryString("id") 'which category to display 0 resets unless itemadd = something
 strStatus = Request.QueryString("status")
 strReset = Request.QueryString("resetID")
 strItemAdd = Request.QueryString("itemadd") 'stores category to select after add itemss

'***
 strResume = Request.QueryString("resumeOrder") 'order number to resume
 If strResume = ""  then
 else
  session("orderNumber") =   session("orderNumberResume")
 end if
'***
 
 If strReset = "true" then
  session("nextDesc") = ""
  session("nextItem") = 0
 end if

 response.write "<TABLE BACKGROUND='img/ordr_frm.gif' border='0' width='725' height='211' cellpadding='0' align='center' CLASS='maintext'>"
 response.write "<TR><TD>&nbsp</td><TD>&nbsp</td><TD align='center'>&nbsp</td></TR>"

 response.write "<TR>"
 response.write "<TD width='30%' align='center'>"

  response.write "<TABLE class='maintext'>"
  Response.Write "<FORM ACTION='myOrder.asp' METHOD='GET'>"
  response.write "<TR>"
  Response.write "<TD align='center'>"
  Response.write "&nbsp"
  response.write "</TD>"
  Response.write "</TR>"
  response.write "<TR>"
  Response.write "<TD>"
  Response.Write "<SELECT NAME='id'>"
  Response.Write "<OPTION VALUE='0'>"
  if Session("distID") = "1001" then
   Response.Write "Select a vendor"
  else
   Response.Write "Select a category"
  end if
  For varCnt = 1 to session("totalCat") 'allow for more categories (or vendors)
'  For varCnt = 1 to 12
   if varCnt = Clng(strCat) or Clng(strItemAdd) = varCnt then
    strName = "<OPTION VALUE='" & varCnt & "' SELECTED>"
   else
    strName = "<OPTION VALUE='" & varCnt & "'>"
   end if
   Response.Write strName
   strName = "cat" & varCnt
   Response.Write Session(strName)
  Next
  Response.write "</select>"
  response.write "<INPUT TYPE='Hidden' NAME='resetID' VALUE='true'"

  response.write "</TD>"
  Response.write "</TR>"
  response.write "<TR>"
  Response.write "<TD>"
  Response.write "Search:&nbsp<INPUT TYPE='Text' NAME='item' SIZE='18' MAXLENGTH='20'><br>"
  response.write "</TD>"
  Response.write "</TR>"
  Response.write "<TR><TD>(Description or Item Number)</td></tr>"
  response.write "<TR>"
  Response.write "<TD align='center'>"
  Response.write "<INPUT Type='SUBMIT' Value='Get Item(s)'>"
  response.write "</TD>"
  Response.write "</TR>"
  response.write "<TR>"
  Response.write "<TD align='center'>"
  Response.write "&nbsp"
  'Response.write "<a href='' STYLE='color : #FFE779' onClick='openHelp(); return false'>Click here for help</a>"
  'Response.write "<a href='WebOrderEntryHelp.pdf' STYLE='color : #FFE779'>Click here for help</a>"
  response.write "</TD>"
  Response.write "</TR>"

  Response.Write "</FORM>"
  response.write "</TABLE>"
  response.write "</TD>"
  response.write "<TD width='30%' align='center'>" 
  Response.Write "Your Distributor:<br>"
  Response.Write "<b>" & session("D_Name") & "</b><br>"
  Response.Write session("D_Addr1") & "<br>"
  Response.Write session("D_City") & "  " &  session("D_State") & "  " &  session("D_Zip") & "<br>"
  strPhoneNumber = "(" & left(session("D_Phone"), 3) & ") " & mid(session("D_Phone"), 5, 3) & "-" & right(session("D_Phone"), 4) 
  Response.Write strPhoneNumber & "<br>" 

'  Response.Write session("D_Phone") & "<br>"
  response.write "</TD>"
  response.write "<TD width='30%' align='center'>"
  Response.Write "Processing order for:<br>"
  Response.Write "<b>" & session("C_Name") & "<br></b>"
  Response.Write "Account Number: " & session("C_Number") & "<br>"
  Response.Write session("C_Address") & "<br>"
  Response.Write session("C_City") & "  " &  session("C_State") & "  " &  session("C_Zip") 
  response.write "</TD>"

  response.write "</TR>"

  response.write "<TR>"
  Response.write "<TD align='center' COLSPAN='3' HEIGHT='40' VALIGN='middle'>"

  Response.write "<A HREF='myOrder.asp?id=0&item=&status=message' onmouseover='swapimage(1,0)' onmouseout='swapimage(1,1)'><IMG NAME='message' SRC='img/am_off.gif' BORDER=0></A><IMG NAME='spacer' SRC='img/spacer.gif' BORDER=0>"
  Response.write "<A HREF='myOrder.asp?id=0&item=&status=viewchange' onmouseover='swapimage(2,0)' onmouseout='swapimage(2,1)'> <IMG NAME='viewchange' SRC='img/vc_off.gif' BORDER=0></A><IMG NAME='spacer' SRC='img/spacer.gif' BORDER=0>"
  Response.write "<A HREF='completedOrder.asp' onmouseover='swapimage(3,0)' onmouseout='swapimage(3,1)'  onclick='return submitVerify()'> <IMG NAME='submitorder' SRC='img/sub_off.gif' BORDER=0></A><IMG NAME='spacer' SRC='img/spacer.gif' BORDER=0>"
  Response.write "<A HREF='logoff.asp' onmouseover='swapimage(4,0)' onmouseout='swapimage(4,1)' onclick='return logoffVerify()'> <IMG NAME='logoff' SRC='img/logoff_off.gif' BORDER=0></A>"
  response.write "</TD>"
 
  
  Response.write "</TR>"
  response.write "</TABLE>"

'start of full page table
 response.write "<TABLE width='640' border='0' cellspacing='0' cellpadding='0' name='HEADER_TABLE' align='center'>" & _
  "<TR><TD>&nbsp</td></tr>" 
 
' if ((strCat >= 1 and strCat <= 12) or (strCat = 0 and Trim(strItem) <> "")) then 'first time cat=-1
 if ((strCat >= 1) or (strCat = 0 and Trim(strItem) <> "")) then 'first time cat=-1  
  Set objCommand = Server.CreateObject("ADODB.Command")
  objCommand.ActiveConnection = objConn_In

  if Trim(strItem) <> "" Then
   'new check for + Quick Add With Qty after + if item found
   'myPos = instr(strItem, "+")
   'myQuickQty = 0
   'if myPos <> 0 then
   ' strItem = left(strItem, myPos - 1)
   ' if myPos = len(strItem) then
'	 'set qty = 1 if + is last character
'	 strQuickQty = "1"
'	else
'	 strQuickQty = right(strItem, len(strItem) - myPos)
'	end if
'    if isnumeric(strQuickQty) then
'	 myQuickQty = strQuickQty
'	end if
'   end if

'   'set variable to process addItemSingle.asp
'   session("quickQty") = myQuickQty
   '

   if isnumeric(strItem) then
    strSQL2 = " item_number =" & strItem 
   else
    strSQL2 = " Description Like '" & strItem & "%'"
   end if
 
  else
   strSQL2 = ""
  end if
 
  if session("nextItem") <> 0 then
   if strSQL2 = "" then
    'strSQL2 = " item_number >=" & session("nextItem") & " and Description >= '" & replace(session("nextDesc"),"'","''") & "'"
	strSQL2 = " item_number <>" & session("lastItem") & " and Description >= '" & replace(session("nextDesc"),"'","''") & "'"
   else
    'strSQL2 = strSQL2 & " and item_number >=" & session("nextItem") & " and Description >= '" & session("nextDesc") & "'"
    strSQL2 = strSQL2 & " and item_number <>" & session("lastItem") & " and Description >= '" & session("nextDesc") & "'"
   end if
 end if
 
 if strCat <> 0 then
   selectCat = strCat
   'get cat select from array to allow selected categories only
   if session("CategorySelect") = 1 then ' (strCat is index of combo - not necessarily sales_category - catRef + (combo.index) contains category from checkLogin.asp
   '  only show/select customer/distributor categories allowed
     strSelect = "catRef" & strCat
     selectCat = session(strSelect) 
   end if
   
    if Trim(strItem) <> "" Then
     strSQL = "SELECT * FROM Inventory WHERE sales_category = " & selectCat & " And " & strSQL2 & " Order by Description"
    else
     if strSQL2 <> "" then
      strSQL = "SELECT * FROM Inventory WHERE sales_category = " & selectCat & " And " & strSQL2 & " Order by Description"
     else
      strSQL = "SELECT * FROM Inventory WHERE sales_category = " & selectCat & " Order by Description"
	 end if
	end if
 else
    'all categories

    if session("CategorySelect") = 1 then ' (strCat is index of combo - not necessarily sales_category - catRef + (combo.index) contains category from checkLogin.asp
   '  only show/select customer/distributor categories allowed
     'strSelect = "catRef" & strCat
     'selectCat = session(strSelect) 
     if session("custCatCount") < 12 then
	  'still check for allowed customer categories
	  'sales_category = " & selectCat &
	  whereClause = "("
	  for i = 1 to 12
	   if mid(Session("cExclude"), i, 1) = "1" then
 	    if whereClause = "(" then
		 whereClause = whereClause & "Sales_Category=" & i
		else
		 whereClause = whereClause & " OR Sales_Category=" & i
		end if
	   end if
	  next
	  whereClause = whereClause & ")"
	 
	  strSQL = "SELECT * FROM Inventory WHERE " & whereClause &  " And " & strSQL2 & " Order by Description" 
	 else
	  strSQL = "SELECT * FROM Inventory WHERE " & strSQL2 & " Order by Description"  
	 end if
	else 
     strSQL = "SELECT * FROM Inventory WHERE " & strSQL2 & " Order by Description"  
    end if
 end if
 objCommand.CommandText = strSQL
 objCommand.CommandType = 1
 Set objRS = objCommand.Execute 
 intRecords = 0
 'objRS.MoveLast
 'response.write "C " & objRS.Recordcount
 'NEW
 '	 if + redirect to addItemSingle.asp with item
 '
 'if no items, don't display the form
 if not objRS.EOF then
  response.write "<FORM ACTION='addItems.asp' NAME='InputForm' METHOD='POST' onsubmit='return checkQTY()'>"  
  response.write "<TABLE width='470' BORDER='0' align='center' cellpadding='0' cellspacing='0'>"
  response.write "<TR><TD align='center'>"
  response.write "<INPUT Type='SUBMIT' Value='Add items to order'>"
  response.write "</TD></TR>"
  response.write "</TABLE><BR>" '
   
  'response.write "<TABLE width='650' BORDER='1' align='center' cellpadding='0' cellspacing='0' CLASS='lineitems'>"
  'if session("ShowPrice") = 1 then
  ' response.write "<TR BGCOLOR='CDD9F7'><TD>&nbsp</td><TD align='center' border='0'><b>Item</b></td><TD align='left' border='0'><b>Size</b></td><TD align='center' border='0'><b>Pack</b></td><TD align='left' border='0'><b>Description</b></td><TD align='center' border='0'><b>Price</b></td><TD align='center' border='0'><b>Quantity</b></td><TD align='center' border='0'><b>&nbsp&nbsp</b></td>" '</TR>"
  'else
  ' response.write "<TR BGCOLOR='CDD9F7'><TD>&nbsp</td><TD align='center' border='0'><b>Item</b></td><TD align='left' border='0'><b>Size</b></td><TD align='center' border='0'><b>Pack</b></td><TD align='left' border='0'><b>Description</b></td><TD align='center' border='0'><b>Quantity</b></td><TD align='center' border='0'><b>&nbsp&nbsp</b></td>" '</TR>"
  'end if
  'if session("ShowImage") = 1  then  'column for camera image
  '  response.write "<TD align='center'>&nbsp</TD>"
  'end if 
  'response.write "/<TR>"
  response.write "<TABLE width='900' BORDER='0' align='center' cellpadding='5' cellspacing='0' CLASS='lineitems'>"
  
  rowColor = 0
  colCount = 1

  dim fs
  set fs=Server.CreateObject("Scripting.FileSystemObject")
  
  While (Not objRS.EOF) and intRecords < 100
   session("lastItem") = objRS("item_number")
   intRecords = intRecords + 1
   'don't start row until 2 columns

   if colCount = 1 then  'NEW ROW
    If rowColor = 0 Then
     Response.Write "<TR BGCOLOR='#dcdcba'>"
     rowColor = 1
    Else
     Response.Write "<TR BGCOLOR='CDD9F7'>" 
     rowColor = 0
    End if 
   end if
  
   'If rowColor = 0 Then
   ' Response.Write "<TR BGCOLOR='#dcdcba'>"
   ' rowColor = 1
   'Else
   ' Response.Write "<TR BGCOLOR='CDD9F7'>" 
   ' rowColor = 0
   'End if 
   '2 column

   response.write "<TD width='300' align='left' >"
   If objRS("breakable")  Then
    response.write "<INPUT TYPE='Hidden' NAME='CanBreak" & intRecords & "' VALUE='Y'>"
   else
    response.write "<INPUT TYPE='Hidden' NAME='CanBreak" & intRecords & "' VALUE='N'>"
   end if
   Response.Write "<b>" & objRS("Description") & "</b>"
   Response.write "<br><br>Item Number: " & objRS("Item_Number")
   Response.write "<br><br>Size: " & objRS("Item_Size") & "&nbsp" & "Pack: " & objRS("Pack")
 
   if session("ShowPrice") = 1 then
    Response.write "<br><br>Price: " & formatcurrency(objRS("price1"), 2)
   end if
   
   Response.Write "<br><br>Quantity: " & "&nbsp" & "<INPUT CLASS='inputText' TYPE='Text' NAME='qty' SIZE='5' MAXLENGTH='5' align='right'>"
  
   'additional lines (add to cell with line break)
   'add additional informational lines if not blank
   'ITEM MESSAGE
   if trim(objRS("item_message")) <> "" then 
	response.write "<br>" & objRS("Item_Message")
   end if 
   'ALT DESC
   if trim(objRS("altdesc")) <> "" then 
	response.write "<br>" & objRS("altdesc")
   end if 
   'VENDOR
   if trim(objRS("v_description")) <> "" then 
	Response.Write "<br>" & "Mfg: " & objRS("v_description")
   end if 
  'VENDOR ITEM
   if trim(objRS("vendor_itemnumberalpha")) <> "" then 
    Response.Write "<br>" & "Mfg Item # " & objRS("vendor_itemnumberalpha")
   end if    

   response.write "</TD>"  'end description cell

    
  'PRICE
  ' if session("ShowPrice") = 1 then
  '  response.write "<TD width='40' align='right'>"
'	Response.Write formatcurrency(objRS("price1"), 2)
'    response.write "</TD>"
'   end if

   'response.write "<TD width='80' align='left'>"
   'Response.Write "Order QTY:"
   'response.write "</TD>"
 '  response.write "<TD width='40' align='right'>"
   'QTY CELL
   'if myQuickQty = 0 then
   ' Response.Write "<INPUT CLASS='inputText' TYPE='Text' NAME='qty' SIZE='5' MAXLENGTH='5' align='right'>"
   'else
   '  Response.Write "<INPUT CLASS='inputText' TYPE='Text' NAME='qty' SIZE='5' MAXLENGTH='5' align='right' VALUE='" & myQuickQty & "'>"  
   'end if
 '  response.write "</TD>" 
 '  If objRS("breakable")  Then
 '   response.write "<TD width='10' align='center'>"
 '   Response.Write "<INPUT TYPE='checkbox' NAME='unitFlag'>Units"
 '   response.write "</TD>" 
 '  else
 '   response.write "<TD width='10' align='center'>"
 '   Response.Write "&nbsp"
 '   response.write "</TD>"
 '  end if

'*** IMAGE *********************************************************************************

 '  if session("ShowImage") = 1 then

     response.write "<TD width = 120 align='center' border='0'>"  
     strImage = "orders/" & Session("distUser") & "/images/" & objRS("item_number") & ".png"
     'if no .jpg, set = .png
	 'if no .png, onerror.jpg
	 if fs.FileExists(Server.MapPath(strImage)) then
     else
      strImage = "orders/" & Session("distUser") & "/images/" & objRS("item_number") & ".jpg"     
	 end if
     strImageID = objRS("item_number") & "_jpg"
	 strImageNotFound  = "orders/" & Session("distUser") & "/images/onerror.jpg"
     'these work
     'Response.write "<img name='itemIMG' src='" & strImage & "' onerror='imgError(this);' width='120' height='144' BORDER='0'/>"
	 'Response.write "<img name='itemIMG' src='" & strImage & "' onerror='imgError(this, " & "&quot;" & strImageNotFound & "&quot;" & ");' width='120' height='144' BORDER='0'/>"
 
     'test
	 Response.write "<a href='" & strImage & "' id='imgURL' onclick='return imgVerify(" & "&quot;" & strImageID & "&quot;" & ")'> <img name='imageName' id='" & strImageID & "' src='" & strImage & "' onerror='imgError(this, " & "&quot;" & strImageNotFound & "&quot;" & ");' width='120' height='144' BORDER='0'/></a>" 
     response.write "</TD>"
 
 '*** IMAGE *********************************************************************************
  
     if colCount = 1 then
     'divider column
	  response.write "<TD width='10' align='left' BGCOLOR='#4682B4'>"
      Response.Write "&nbsp&nbsp"
      response.write "</TD>"
     end if
  
   '** end 2 col

   response.write "<INPUT TYPE='Hidden' NAME='ItemID" & intRecords & "' VALUE='" & objRS("item_number") & "'>"
  ' response.write "<INPUT TYPE='Hidden' NAME='QtyOrdered" & intRecords & "' VALUE='0'>"

   myNewDesc = ""
   for i = 1 to len(objRS("description"))
    myChar = mid(objRS("description"), i, 1)
    if myChar <> "'" then
	 myNewDesc = myNewDesc & myChar
	end if 
   next
 '  response.write "<INPUT TYPE='Hidden' NAME='desc" & intRecords & "' VALUE='" & objRS("description") & "'>"
   response.write "<INPUT TYPE='Hidden' NAME='desc" & intRecords & "' VALUE='" & myNewDesc & "'>"

 '  if Session("distID") = "1001" then
 '     response.write "<INPUT TYPE='Hidden' NAME='desc30" & intRecords & "' VALUE='" & objRS("description30") & "'>"
 '  end if
   response.write "<INPUT TYPE='Hidden' NAME='size" & intRecords & "' VALUE='" & objRS("item_size") & "'>"
   response.write "<INPUT TYPE='Hidden' NAME='pack" & intRecords & "' VALUE='" & objRS("pack") & "'>"
   response.write "<INPUT TYPE='Hidden' NAME='scat" & intRecords & "' VALUE='" & objRS("sales_category") & "'>"
   response.write "<INPUT TYPE='Hidden' NAME='price" & intRecords & "' VALUE='" & objRS("price1") & "'>"
   'response.write "</TR>"

    'don't end row until 2 columns
   if colCount = 2 then
    response.write "</TR>"
    colCount = 1
   else
    colCount = colCount + 1
   end if
   
   objRS.MoveNext
  Wend

  set fs=nothing


  response.write "<INPUT TYPE='Hidden' NAME='iTotal' VALUE='" & intRecords & "'>"
  response.write "<INPUT TYPE='Hidden' NAME='selectCat' VALUE='" & strCat & "'>"
'  if intRecords = 0 then
'   response.write "<TABLE>"
'   response.write "<TR BGCOLOR='FFFFC0'>"
'   response.write "<TD>"
'   Response.Write "<b>No items found. Please select a category or search by item number or description.</b>"
'   response.write "</TD>"
'   response.write "</TR>"
'   response.write "</TABLE>"
'  end if
  response.write "</table><BR>"
  response.write "<TABLE width='650' BORDER='0' align='center' cellpadding='0' cellspacing='0'>"
  response.write "<TR><TD align='center'>"
  response.write "<INPUT Type='SUBMIT' Value='Add items to order'>"
  response.write "</TD></TR>"
  response.write "</TABLE>"  
 ' if intRecords = 1 then '1 item only
 ' response.write "<script  language='javascript' type='text/javascript'> document.getElementById('qty').focus();}</script>"
  'response.write "<script  language='javascript' type='text/javascript'> alert("here");</script>"
    ' end if
  response.write "</form>"
 else 'no items
  
  'response.write "<TABLE border='0' width='600' cellpadding='0' align='center'>"
  'response.write "<TR><TD>&nbsp</td></tr>"
  'Response.write "<TR align='center'><TD><font color='#FFFFCO'><b>" & "Entered: " & strItem & " - " & " No items found. Please select a category or search by item number or description.</b></font></td></tr>"
  'response.write "</TABLE>"
  
  '**NEW  - CREATE FORM TO ALLOW Enter / Set Focus to Input (iTotal = 0)
  response.write "<FORM ACTION='' NAME='InputForm' METHOD='POST'>"  
  'response.write "<BR>"
  response.write "<TABLE width='650' BORDER='1' align='center' cellpadding='0' cellspacing='0' CLASS='lineitems'>"
  response.write "<TR BGCOLOR='CDD9F7'><TD align='center' border='0'><b>Search Text= " & strItem & "</b></td></TR>"
  response.write "<TR BGCOLOR='#dcdcba'><TD align='center' border='0'><b>* No Matching Item(s) Found *</b></td></TR>"
  response.write "<TR BGCOLOR='CDD9F7'><TD align='center' border='0'><b>Please select a category or search by item number or description</b></td></TR>"
  response.write "<INPUT TYPE='Hidden' NAME='iTotal' VALUE='0'>"
  response.write "</TABLE>"  
  response.write "</form>"
  '**NEW
 
 end if 
 if not objRs.EOF then

   session("nextItem") = objRS("item_number")
   session("nextDesc") = Server.HTMLEncode(objRS("description"))
 '  session("reset100") = 1
   response.write "<TABLE width='650' BORDER='0' align='center' cellpadding='0' cellspacing='0'>"
   response.write "<TR><TD align='center'>"

   newQueryString = "item=" & strItem & "&id=" & strCat & "&resetID=false" & "&status="
'   response.write "<a href='myOrder.asp?" & Request.QueryString & "' STYLE='{color : gold}'>Next 100 Items</a>"
   session("resumeString") = newQueryString
'   Response.write "<a href='' STYLE='color : #FFE779' onClick='openHelp(); return false'>Click here for help</a>"

   response.write "<a href='myOrder.asp?" & newQueryString & "' STYLE='{color : gold}' onClick='return checkQTY3();'>More Items</a>"

   response.write "</TD></TR>"
   response.write "</TABLE>"
   response.write  "</table>" 'end of full page table
 else
'   session("reset100") = 0
   session("nextItem") = 0
   session("lastItem") = 0
   session("nextDesc") = ""
   session("resumeString") = ""
 end if
 Set objRS = Nothing
end if

'VIEW/CHANGE
 if strStatus = "viewchange" then 'view current order
  Set objCommand = Server.CreateObject("ADODB.Command")
  objCommand.ActiveConnection = objConn
  strSQL = "SELECT * From Order_Detail Where Order_number=" & session("orderNumber") & " order by line_number"
  objCommand.CommandText = strSQL
  objCommand.CommandType = 1
  Set rsView = objCommand.Execute 
 'if no items, don't display the form
  if not rsView.EOF then
   response.write "<FORM ACTION='updateItems.asp' NAME='updateForm' METHOD='POST' onsubmit='return checkQTY2()'>"  
   response.write "<TABLE width='650' BORDER='0' align='center' cellpadding='0' cellspacing='0'>"
   response.write "<TR><TD align='center'>"
   response.write "<INPUT Type='SUBMIT' Value='Update Changes'>"
   response.write "</TD></TR>"
   response.write "</TABLE><BR>"
   response.write "<TABLE width='650' BORDER='1' align='center' cellpadding='0' cellspacing='0' CLASS='lineitems'>"

  if session("ShowPrice") = 1 then
   response.write "<TR BGCOLOR='CDD9F7'><TD>&nbsp</td><TD align='center' border='0'><b>Item</b></td><TD align='left' border='0'><b>Size</b></td><TD align='center' border='0'><b>Pack</b></td><TD align='left' border='0'><b>Description</b></td><TD align='center' border='0'><b>Price</b></td><TD align='center' border='0'><b>Quantity</b></td><TD>&nbsp</td></TR>"
  else
   response.write "<TR BGCOLOR='CDD9F7'><TD>&nbsp</td><TD align='center' border='0'><b>Item</b></td><TD align='left' border='0'><b>Size</b></td><TD align='center' border='0'><b>Pack</b></td><TD align='left' border='0'><b>Description</b></td><TD align='center' border='0'><b>Quantity</b></td><TD>&nbsp</td></TR>"
  end if
 
   rowColor = 0
   intRecords = 0
   orderTotal = 0
   While Not rsView.EOF 
    intRecords = intRecords + 1
    If rowColor = 0 Then
     Response.Write "<TR BGCOLOR='#dcdcba'>" 
    rowColor = 1
    Else
     Response.Write "<TR BGCOLOR='CDD9F7'>" 
     rowColor = 0
    End if 
    response.write "<TD align='right'>"
    Response.Write rsView("Line_Number")
    response.write "</TD>"
    response.write "<TD width='60' align='right'>"
    Response.Write rsView("item_number")
    response.write "</TD>"

    if isnull(rsView("item_size")) or trim(rsView("item_size")) = "" then
     Response.write "<TD align='left'>&nbsp</TD>" 'formatting
    else
     response.write "<TD align='left'>"
     Response.Write rsView("item_size")
     response.write "</TD>"
    end if
 '   response.write "<TD align='left'>"
 '   Response.Write rsView("item_size")
 '   response.write "</TD>"
    response.write "<TD width='40' align='right'>"
    Response.Write rsView("item_pack")
    response.write "</TD>"

    response.write "<TD>"
    Response.Write rsView("Description")
    response.write "</TD>"

	if session("ShowPrice") = 1 then
     response.write "<TD width='40' align='right'>"
  	 Response.Write formatcurrency(rsView("price"), 2)
     response.write "</TD>"
     orderTotal = orderTotal + (rsView("quantity_ordered")*rsView("price"))
    end if

    'response.write "<TD width='40' align='right'>"
	response.write "<TD align='center'>"
    myQty = rsView("quantity_ordered") ' & " " & rsView("delivery_code")
    response.write "<INPUT CLASS='inputText' TYPE='Text' NAME='qty' SIZE='5' MAXLENGTH='5' VALUE='" & myQty & "'>&nbsp" & rsView("delivery_code") & "&nbsp"
    response.write "</TD>" 

    response.write "<TD align='center'>"
    Response.Write "<INPUT TYPE='checkbox' NAME='delFlag'>Delete"
    Response.Write "<INPUT TYPE='Hidden' NAME='uptDelete' VALUE='N'>"
    Response.Write "<INPUT TYPE='Hidden' NAME='uptLine' VALUE='" & rsView("Line_Number") & "'>"
    response.write "<INPUT TYPE='Hidden' NAME='origQty' VALUE='" & myQty & "'>"
    response.write "</TD>" 
    response.write "</TR>"   
    rsView.MoveNext
   Wend

   if session("ShowPrice") = 1 then   
    'DISPLAY TOTAL LINE
    If rowColor = 0 Then
     Response.Write "<TR BGCOLOR='#dcdcba'>" 
    Else
     Response.Write "<TR BGCOLOR='CDD9F7'>" 
    End if 
    response.write "<TD COLSPAN='8' align='center'>" 
	response.write "&nbsp"
	response.write "</TD>" 
    response.write "</TR>"      

    If rowColor = 0 Then
     Response.Write "<TR BGCOLOR='#dcdcba'>" 
    Else
     Response.Write "<TR BGCOLOR='CDD9F7'>" 
    End if   
    response.write "<TD COLSPAN='8' align='center'>" 
	response.write "<b>ORDER TOTAL: " & formatcurrency(orderTotal, 2) & "</b>"
	response.write "</TD>" 
    response.write "</TR>"     
   end if
   
   response.write "</table><BR>"
   response.write "<TABLE width='650' BORDER='0' align='center' cellpadding='0' cellspacing='0'>"
   response.write "<TR><TD align='center'>"
   response.write "<INPUT Type='SUBMIT' Value='Update Changes'>"
   response.write "</TD></TR>"
   response.write "<INPUT TYPE='Hidden' NAME='iTotal' VALUE='" & intRecords & "'>"
   response.write "</TABLE>"
   response.write "</TABLE></FORM>"
  else
   'response.write "<TABLE border='0' width='600' cellpadding='0' align='center'>"
   'response.write "<TR><TD>&nbsp</td></tr>"
   'Response.write "<TR align='center'><TD><font color='#FFFFCO'><b>No items have been added to your order.</b></font></td></tr>"
   'response.write "</TABLE>"
   response.write "<TABLE width='600' BORDER='1' align='center' cellpadding='0' cellspacing='0' CLASS='lineitems'>"
   response.write "<TR BGCOLOR='CDD9F7'><TD align='center' border='0'><b>No items have been added to your order.</b></td></TR>"
   response.write "</TABLE>"  

  end if
  set rsView = nothing
 end if

'VIEW ONLY
 if strStatus = "view" then 'view current order
  Set objCommand = Server.CreateObject("ADODB.Command")
  objCommand.ActiveConnection = objConn
  strSQL = "SELECT * From Order_Detail Where Order_number=" & session("orderNumber") & " order by line_number"
  objCommand.CommandText = strSQL
  objCommand.CommandType = 1
  Set rsView = objCommand.Execute 
 'if no items, don't display the form

'***
  lastLineNumberResume = 0
'***
  orderTotal = 0
  if not rsView.EOF then
  
   response.write "<TABLE width='650' BORDER='1' align='center' cellpadding='0' cellspacing='0' CLASS='lineitems'>"
   if session("ShowPrice") = 1 then
    response.write "<TR BGCOLOR='CDD9F7'><TD>&nbsp</td><TD align='center' border='0'><b>Item</b></td><TD align='left' border='0'><b>Size</b></td><TD align='center' border='0'><b>Pack</b></td><TD align='left' border='0'><b>Description</b></td><TD align='center' border='0'><b>Price</b></td><TD align='center' border='0'><b>Quantity</b></td></TR>"
   else
    response.write "<TR BGCOLOR='CDD9F7'><TD>&nbsp</td><TD align='center' border='0'><b>Item</b></td><TD align='left' border='0'><b>Size</b></td><TD align='center' border='0'><b>Pack</b></td><TD align='left' border='0'><b>Description</b></td><TD align='center' border='0'><b>Quantity</b></td></TR>"
   end if
   rowColor = 0
   intRecords = 0
   While Not rsView.EOF 
    intRecords = intRecords + 1
    If rowColor = 0 Then
     Response.Write "<TR BGCOLOR='#dcdcba'>" 
    rowColor = 1
    Else
     Response.Write "<TR BGCOLOR='CDD9F7'>" 
     rowColor = 0
    End if 
    response.write "<TD align='right'>"

    Response.Write rsView("Line_Number")

'***
    lastLineNumberResume = rsView("Line_Number") 'set last line number at end if order resumed
'***

    response.write "</TD>"
    response.write "<TD width='60' align='right'>"
    Response.Write rsView("item_number")
    response.write "</TD>"
    if isnull(rsView("item_size")) or trim(rsView("item_size")) = "" then
     Response.write "<TD align='left'>&nbsp</TD>" 'formatting
    else
     response.write "<TD align='left'>"
     Response.Write rsView("item_size")
     response.write "</TD>"
    end if
'    response.write "<TD align='left'>"
'    Response.Write rsView("item_size")
'    response.write "</TD>"

    response.write "<TD width='40' align='right'>"
    Response.Write rsView("item_pack")
    response.write "</TD>"

    response.write "<TD>"
    Response.Write rsView("Description")
    response.write "</TD>"

	if session("ShowPrice") = 1 then
     response.write "<TD width='40' align='right'>"
  	 Response.Write formatcurrency(rsView("price"), 2)
     response.write "</TD>"
     orderTotal = orderTotal + (rsView("quantity_ordered")*rsView("price"))
	end if
	 
    response.write "<TD CLASS='inputText' width='40' align='right' COLSPAN='2'>"
    response.write rsView("quantity_ordered") & " " & rsView("delivery_code")
    response.write "</TD>" 
    response.write "</TR>"   
    

    rsView.MoveNext
   Wend
   
   if session("ShowPrice") = 1 then   
    'DISPLAY TOTAL LINE
    If rowColor = 0 Then
     Response.Write "<TR BGCOLOR='#dcdcba'>" 
    Else
     Response.Write "<TR BGCOLOR='CDD9F7'>" 
    End if 
    response.write "<TD COLSPAN='7' align='center'>" 
	response.write "&nbsp"
	response.write "</TD>" 
    response.write "</TR>"      

    If rowColor = 0 Then
     Response.Write "<TR BGCOLOR='#dcdcba'>" 
    Else
     Response.Write "<TR BGCOLOR='CDD9F7'>" 
    End if   
    response.write "<TD COLSPAN='7' align='center'>" 
	response.write "<b>ORDER TOTAL: " & formatcurrency(orderTotal, 2) & "</b>"
	response.write "</TD>" 
    response.write "</TR>"     
   end if
   
   
'***
   If strResume = ""  then
   else
    session("lineNumber") =  lastLineNumberResume
   end if
'***

   response.write "</TABLE>"
   if session("resumeString") <> "" then
    response.write "<TABLE width='650' BORDER='0' align='center' cellpadding='0' cellspacing='0'>"
    response.write "<TR><TD align='center'>&nbsp</TD></TR>"

    response.write "<TR><TD align='center'>"
    response.write "<a href='myOrder.asp?" & session("resumeString") & "' STYLE='{color : gold}'>Resume Next Items</a>"
    response.write "</TD></TR>"
    response.write "</TABLE>"
   end if
  else
   'response.write "<TABLE border='0' width='600' cellpadding='0' align='center'>"
   'response.write "<TR><TD>&nbsp</td></tr>"
   'Response.write "<TR align='center'><TD><font color='#FFFFCO'><b>There are currently no items in your order.</b></font></td></tr>"
   'response.write "</TABLE>"
   response.write "<TABLE width='600' BORDER='1' align='center' cellpadding='0' cellspacing='0' CLASS='lineitems'>"
   response.write "<TR BGCOLOR='CDD9F7'><TD align='center' border='0'><b>There are currently no items in your order.</b></td></TR>"
   response.write "</TABLE>"  
  ' response.write "<TR BGCOLOR='#dcdcba'><TD align='center' border='0'><b>There are currently no items in your order.</b></td></TR>"
  ' response.write "<TR BGCOLOR='CDD9F7'><TD align='center' border='0'>&nbsp</td></TR>"


  end if
  set rsView = nothing
 end if
End if

'ATTACH MESSAGE
 if strStatus = "message" then

  response.write "<TABLE border='0' width='500' height='211' cellpadding='0' align='center'>"
  Response.Write "<FORM ACTION='updateMessage.asp' METHOD='GET'>"

  'Response.write "<TR align='center'><TD><font color='#FFFFCO'><b>Enter a message to attach to your order (maximum 50 characters)</b></font></td></tr>"
  response.write "<TR BGCOLOR='CDD9F7'><TD align='center' border='0'><b>Enter a message to attach to your order (maximum 50 characters)</b></td></TR>"

  Response.write "<TR align='center'><TD><TEXTAREA NAME='message' COLS='30' ROWS='3'>" & session("orderMessage") & "</textarea></td></tr>"
  response.write "<TR>"
  Response.write "<TD align='center'>"
  Response.write "<INPUT Type='SUBMIT' Value='Update Message'>"
  response.write "</TD>"
  Response.write "</TR>"
  Response.Write "</FORM>"
  response.write "</TABLE>"

 end if
 Set objCommand = Nothing 
 Set objConn = Nothing
 Set objConn_In = Nothing
%>

</body>

</html>
