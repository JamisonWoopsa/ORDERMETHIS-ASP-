<html>
<head>
<title>Internet Order Entry for candy, tobacco and grocery distributors</title>

</head>
<body onLoad = "document.forms[0].Password.focus();" background="img/bg.gif">
<script language="Javascript">
function openHelp() {
 var newWindow = window.open("WebOrderEntryHelp.pdf", "_blank", "toolbar=no, scrollbars=yes, resizable=yes");
 newWindow.focus();
}

function checkDist() {
 var badDist;
 badDist = 0;
 //if (isNaN(document.signon.did.value)) {
 //    alert("Invalid Distributor ID ... Numeric Entry Required"); allow char alias
 //    badDist = 1;
 //  }
 if (isNaN(document.signon.cid.value)) {
     alert("Invalid Customer ID ... Numeric Entry Required");
     badDist = 1;
   } 
 
 if (badDist > 0) 
   return false; 
 else
   return true;
}


//browserType = navigator.appName
//browserVer = navigator.appVersion
//redirect to a different web page if browser not Microsoft Internet Explorer 4.0 or better
//if (browserType != "Microsoft Internet Explorer") {
//  strMessage = "Microsoft Internet Explorer Version 4.0 or better is required to utilize 'Web Order Entry'. Your current browser is: ";
//  alert(strMessage + browserType);
//  location.href = "http://www.ordermethis.com/invalidBrowser.html";
//  }
</script>
<%

 if Request("SecondTry") <> "True" then  'get info from cookie if first time
  session("distID") = request.Cookies("wsLogin") ("did")
  session("custID") = request.Cookies("wsLogin") ("cid")
  if session("distID") <> "" then
   session("strLogo") = "img/" & session("distID") & ".jpg"
  else
   session("strLogo") = "img/logo0.jpg"  'default logo  
  end if
 elseif session("strLogo") = "" then
  session("strLogo") = "img/logo0.jpg"  'default logo
 end if 
%>

<table width="640" border="0" cellspacing="0" cellpadding="0" name="HEADER_TABLE" align="center">
  <tr align="center"> 
    <td ><img src="<%= Session("strLogo")%>" width="544" height="120" border="0">

	</td>
  </tr> 

  <tr align="center"> 
    <td >&nbsp</td>
  </tr> 
</table>
<P>

<table width="60%" border="0" cellspacing="0" cellpadding="5"  align="center">
    <tr align="left"><td colspan="3">
	</td>
	</tr>

   <tr align="center"><td> 
   <FORM ACTION="CheckLogin.asp" METHOD="POST" NAME="signon" onsubmit="return checkDist()">
   <TABLE BACKGROUND="img/login.gif" width="358" height="246" border="0" cellspacing="0" cellpadding="0">
    <TR align='left'>
      <TD width="45%">&nbsp</TD>
    </TR>
    <TR align='left'>
      <TD width="45%">&nbsp</TD>
    </TR>
    <TR align='left'>
      <TD width="45%" align="center"><IMG src="img/loginA.gif"></TD>
      <TD><INPUT TYPE="Text" NAME="did" 

            VALUE="<%= Session("distID") %>"

          SIZE="15" MAXLENGTH="15"></TD>
    </TR>
    <TR align="left">
      <TD width="45%" align="center"><IMG src="img/loginB.gif"></TD>
      <TD><INPUT TYPE="Text" NAME="cid" 

            VALUE="<%= Session("custID") %>"

          SIZE="6" MAXLENGTH="6"></TD>
    </TR>	
    <TR align="left">
      <TD width="45%" align="center"><IMG src="img/loginC.gif"></TD>
      <TD><INPUT TYPE="Password" NAME="Password" SIZE="10" MAXLENGTH="10"></TD>
    </TR>
    <TR align="left">
	  <TD width="45%" align="center"></TD>
      <TD><INPUT SRC="img/login1.gif" TYPE="Image" NAME="loginbutton" VALUE="Login" ></TD>
    </TR>
   </FORM>
   </td></tr>
   
   <TR align="left">
      <TD width="45%" align="center" colspan="2"><a align="center" href="WebOrderEntryHelp.pdf" STYLE="color : #FFE779" >Web Order Entry - Users Guide</a><br><br></TD>
   </TR> 

   </TABLE>

   <TABLE>
   <tr><td>&nbsp
   <% 
     If Request("SecondTry") = "True" Then        ' User's second attempt at logging in
       If Request("Dist") = "False" Then 
         Response.Write "<tr><td><font color='#FFFFCO'><b>Invalid Distributor ID. Please try again.</b></font><BR></td></tr>"
       elseIf Request("WrongPW") = "True" Then 
         Response.Write "<tr><td><font color='#FFFFCO'><b>Invalid Password. Please try again.</b></font><BR></td></tr>"
       Else
         Response.Write "<tr><td><font color='#FFFFCO'><b>Customer ID not found. Please try again.</b></font><BR></td></tr>"
       End If
     End If
   %>
   </td></tr>
   </table>
</TABLE>


</BODY>
<script language="Javascript">
newImage = new Image()
newImage.src = "img/ordr_frm.gif"
</script>

</HTML>
