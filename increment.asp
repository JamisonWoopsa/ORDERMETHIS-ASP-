<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
Order number increment +1
<%
response.write "Before increment: " & Application("order_number") & "<BR>"
Application("order_number") = Application("order_number") + 1
response.write "After increment: " & Application("order_number")
%>


</body>
</html>
