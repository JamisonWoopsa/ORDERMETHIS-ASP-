<html>
<head>
<title>Software for candy, tobacco and grocery distributors</title>

</head>
<body bgcolor=white>
<%
 ' Dim objConn, strConnect
 ' strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\inetpub\wwwroot\data\db3.mdb" 
 ' Set objConn = Server.CreateObject("ADODB.Connection")
 ' objConn.Open strConnect
%>
<%
  Dim objConn, strConnect
  strConnect = "D10100190-db3" 
  Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open strConnect
%>
<%
Dim strCust, strAddr, strCity, strSt, strZip, strContact, strEmail, strPhone, strComments

strCust1 = request.form("cust")
strAddr1 = request.form("addr")
strCity1 = request.form("city")
strSt1 = request.form("st")
strZip1 = request.form("zip")
strContact1 = request.form("contact")
strPassword = request.form("password")
strPhone1 = request.form("phone")
strAccount1 = request.form("account")

Set rsOrder = Server.CreateObject("ADODB.Recordset")
 rsOrder.Open "customer", objConn, 0, 3, 2
 rsOrder.AddNew
 rsOrder("C_Number") = strAccount1
 rsOrder("C_Name") = strCust1
 rsOrder("C_CoName") = strContact1
 rsOrder("C_Address") = strAddr1
 rsOrder("C_City") = strCity1
 rsOrder("C_Zip") = strZip1
 rsOrder("Password") = strPassword
 rsOrder("C_Phone") = strPhone1
 rsOrder.Update
 rsOrder.Close
 Set rsOrder = nothing
 set objConn = nothing
%>
<h1>Customer added.</h1>
<%=strCust1%><br>
<%=strAddr1%><br>
<%response.write strCity1 & " " & strSt1 & " " & strZip1 & "<br>"%>
Contact: <%=strContact1%><br>
<%=strPhone1%><br>

</body>
</html>
