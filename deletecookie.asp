<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
Delete cookie info
<%
response.Cookies("wsLogin").expires = date - 1000
%>


</body>
</html>
