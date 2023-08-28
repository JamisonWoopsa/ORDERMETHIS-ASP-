
<!--#include file="jconnection.asp"-->
<%

session("orderNumberResume") = Request.Form("selOption")

if Request.Form("selResume") = "DELETE ORDER" then
 'redirect to confirm delete
 Response.Redirect "myOrderResumeDeleteConfirm.asp?resumeOrder=" & session("orderNumberResume")
else
 Response.Redirect "myOrder.asp?id=-1&item=-1&status=view&resumeOrder=" & session("orderNumberResume")
end if
%>