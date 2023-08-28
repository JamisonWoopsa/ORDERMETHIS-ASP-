<html>
<head>
<title>Internet Order Entry for candy, tobacco and grocery distributors</title>

<STYLE>
.maintext {color: black; font-family: "ms sans serif", "arial"; font-size: 10pt}
.lineitems {color: black; font-family: "ms sans serif", "arial"; font-size: 10pt}
</style> 
</head>
<body background="img/bg.gif" LINK="#7690B2" ALINK="#7690B2" VLINK="#7690B2">

<!--#include file="jconnection.asp"-->
<Script language="Javascript">

function submitVerify(){
	//from web sample
    var val;
    // get list of radio buttons with specified name
    form = document.getElementById('InputForm')
	name = "selResume"
	var radios = form.elements[name];
    
    // loop through list of radio buttons
    for (var i=0, len=radios.length; i<len; i++) {
        if ( radios[i].checked ) { // radio checked?
            val = radios[i].value; // if so, hold its value in val
            break; // and break out of for loop
        }
    }

    if (val == "RESUME ORDER") { return confirm("OK to RESUME selected order ?"); }
	else { return confirm("OK to DELETE selected order ?"); }
}

</script>

<%
 Const adCmdText = &H0001
 'update order header COMPLETE
 Set objCommand = Server.CreateObject("ADODB.Command")
 objCommand.ActiveConnection = objConn
 
'TABLE TO CONTAIN BOTH TABLES
'TABLE FOR HEADER
 
' response.write "The following orders were previously started within the last 7 days but not submitted.<br><br>"

   response.write "<TABLE width='400' BORDER='1' align='center' cellpadding='0' cellspacing='0'>" 'item list table

   Response.Write "<FORM ACTION='myOrderResumeLaunch.asp' NAME='InputForm' METHOD='POST'>" 
   response.write "<TR BGCOLOR='#CDD9F7'>"
   Response.write "<TD COLSPAN='3'>"
   Response.Write "<b>Orders previously started but not submitted (Last 7 Days).</b>"
   response.write "</TD>"
   Response.write "</TR>"
 
   response.write "<TR BGCOLOR='#CDD9F7'>"
   Response.write "<TD COLSPAN='3'>"
   Response.Write "&nbsp"
   response.write "</TD>"
   Response.write "</TR>"
      
'   Response.Write "<FORM ACTION='myOrderResumeLaunch.asp' NAME='InputForm' METHOD='POST'>" 
   response.write "<TR BGCOLOR='#CDD9F7'>"
   Response.write "<TD>" '<b>Size</b>
   Response.Write "<b>Order Reference</b>"
   response.write "</TD>"
   Response.write "<TD>"
   Response.Write "<b>Order Date</b>"
   response.write "</TD>"
   Response.write "<TD>"
   Response.Write "<b>Select</b>"
   response.write "</TD>"
   Response.write "</TR>"

  testDate = Date - 7
  set rsResume = Server.CreateObject("ADODB.Recordset")
  strSQL = "SELECT * FROM Order_Header where c_number = " & Session("custID") & " And Not Submitted And order_date >= #" & testDate & "# Order By Order_Date DESC"
  rsResume.Open strSQL, objConn
  ct = 0
  rowColor = 1

  While Not rsResume.EOF

   If rowColor = 0 Then
    Response.Write "<TR BGCOLOR='#CDD9F7'>" 
    rowColor = 1
   Else
    Response.Write "<TR BGCOLOR='#dcdcba'>" 
    rowColor = 0
   End if 

   Response.write "<TD>"
   Response.Write rsResume("order_number")
   response.write "</TD>"
   
   Response.write "<TD>"
   Response.Write rsResume("order_date") & " (" & rsResume("order_time") & ")"
   response.write "</TD>"
   
   response.write "<TD align='center'>"
   if ct = 0 then '1st is default
    Response.Write "<INPUT TYPE='RADIO' NAME='selOption' VALUE='" & rsResume("order_number") & "' checked>"
   else
    Response.Write "<INPUT TYPE='RADIO' NAME='selOption' VALUE='" & rsResume("order_number") & "'>"
   end if
   response.write "</TD>" 


 'debug
 'Response.write "<TD>&nbsp</TD>"
 
 
   response.write "/<TR>"
   ct = ct + 1
   rsResume.MoveNext
  Wend

  




response.write "<TR><TD BGCOLOR='CDD9F7' COLSPAN='3' align='center'>&nbsp" 'table inside row
'Response.Write "<INPUT TYPE='Hidden' NAME='oTotal' VALUE='" & ct & "'>"
response.write "</TD></TR>"

response.write "<TR>"
Response.write "<TD BGCOLOR='CDD9F7' align='center' COLSPAN='2'>"
Response.Write "<b>RESUME Selected Order</b>"
response.write "</TD>"
Response.write "<TD BGCOLOR='CDD9F7' align='center'>"
Response.Write "<INPUT ID='RESUME' TYPE='RADIO' NAME='selResume' VALUE='RESUME ORDER' checked>"
response.write "</TD></TR>"

response.write "<TR>"
Response.write "<TD BGCOLOR='CDD9F7' align='center' COLSPAN='2'>"
Response.Write "<b>DELETE Selected Order</b>"
response.write "</TD>"
Response.write "<TD BGCOLOR='CDD9F7' align='center'>"
Response.Write "<INPUT ID='DELETE' TYPE='RADIO' NAME='selResume' VALUE='DELETE ORDER'>"
response.write "</TD></TR>"

response.write "<TR><TD BGCOLOR='CDD9F7' COLSPAN='3' align='center'>&nbsp" 'table inside row
'Response.Write "<INPUT TYPE='Hidden' NAME='oTotal' VALUE='" & ct & "'>"
response.write "</TD></TR>"

response.write "<TR>"
Response.write "<TD BGCOLOR='CDD9F7' COLSPAN='3' align='center'>"
Response.write "<INPUT Type='SUBMIT' SIZE='20' Value='Continue'>"
response.write "</TD>"
response.write "</TR>"

'response.write "</FORM>"

response.write "<TR><TD BGCOLOR='CDD9F7' COLSPAN='3' align='center'>&nbsp" 'table inside row
'Response.Write "<INPUT TYPE='Hidden' NAME='oTotal' VALUE='" & ct & "'>"
response.write "</TD></TR>"

response.write "</FORM>"

Response.Write "<FORM ACTION='myOrder.asp?id=-1&item=-1' METHOD='POST'>"
response.write "<TR>"
Response.write "<TD BGCOLOR='CDD9F7' COLSPAN='3' align='center'>"
Response.write "<INPUT Type='SUBMIT' SIZE='20' Value='Start New Order'>"
response.write "</TD>"
response.write "</TR>"
response.write "</FORM>"

response.write "</TABLE>"  '


' response.write "</TD>"
' response.write "</TR>"

'response.write "<TR>"
'Response.write "<TD>"

'response.write "<TABLE>" 

'response.write "</TABLE>"  '

'response.write "</TD>"
'response.write "</TR>"

'response.write "</TD>"
'response.write "</TR>"
'response.write "</TABLE>"  'end container table

set objCommand = nothing
set rsView = nothing
'Set objConn = Nothing
'set objFSO = Nothing
'set objOrderFile = Nothing
'  Response.write "<BR><A href='logoff.asp'>Click here to Log Off</a>"
%>
</BODY>
</HTML>
