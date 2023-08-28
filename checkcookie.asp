<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
Get cookie info
<%
strDist = request.Cookies("wsLogin") ("did")
strCust = request.Cookies("wsLogin") ("cid")
'strExp = request.cookies("wsLogin").Expires
response.write "Distributor " &  strDist & "<br>"
response.write "Customer " & strCust & "<br>"
'response.write "Expires " & strExp
%>


</body>
</html>
