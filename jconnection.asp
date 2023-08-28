<%
  Dim objConn, strConnect, objConn_In, strConnect_In, strDatabaseFolder
  'set user name 
  'database path = database folder + distUser
  'orders database file name = wsdist_web.mdb
  'create separate connection for data to read = wsdist_webU.mdb

  if ucase(Session("distID")) = "1" or ucase(Session("distID")) = "JAMISON" then 'JAMISON COMPUTER 
   Session("distUser") = "JAM201105"
  elseif ucase(Session("distID")) = "EAS201309" or ucase(Session("distID")) = "EW"  then 'EAST WEST DISTRIBUTORS
   Session("distUser") = "EAS201309"  
  elseif ucase(Session("distID")) = "JJD" or ucase(Session("distID")) = "JJD201211"  then 'JJ DISTRIBUTORS
   Session("distUser") = "JJD201211"  
  elseif ucase(Session("distID")) = "JOHNINC" or ucase(Session("distID")) = "JOH202006"  then 'JOHN INC
   Session("distUser") = "JOH202006"  
  elseif ucase(Session("distID")) = "SOMERSET" or ucase(Session("distID")) = "SOM201108"  then 'SOMERSET
   Session("distUser") = "SOM201108"  
  elseif ucase(Session("distID")) = "ARD" or ucase(Session("distID")) = "ARD201905"  then 'AR DISTRIBUTORS
   Session("distUser") = "ARD201905"  
  elseif ucase(Session("distID")) = "LOU" or ucase(Session("distID")) = "LOU202101"  then 'LOU
   Session("distUser") = "LOU202101"  
  elseif ucase(Session("distID")) = "BEARTRACKS" or ucase(Session("distID")) = "BEA202103"  then 'BEAR TRACKS
   Session("distUser") = "BEA202103"  
  elseif ucase(Session("distID")) = "UNIQUE" or ucase(Session("distID")) = "SNE201907"  then 'S&N / Unique
   Session("distUser") = "SNE201907"  
  elseif ucase(Session("distID")) = "AMS" or ucase(Session("distID")) = "AMS202103"  then 'AMS
   Session("distUser") = "AMS202103"  
  elseif ucase(Session("distID")) = "SATYA" or ucase(Session("distID")) = "SAT202105"  then 'SATYA
   Session("distUser") = "SAT202105"  
  elseif ucase(Session("distID")) = "JES" or ucase(Session("distID")) = "JES202108"  then 'JES
   Session("distUser") = "JES202108"  
  elseif ucase(Session("distID")) = "BLAZE" or ucase(Session("distID")) = "BLZ202210"  then 'BLAZE ORLANDO
   Session("distUser") = "BLZ202210"  
  elseif ucase(Session("distID")) = "SUNSET" or ucase(Session("distID")) = "SUN202302"  then 'SUNSET
   Session("distUser") = "SUN202302"  
  end if

  strDatabaseFolder = "/_database/" & Session("distUser") & "/"
  strConnect = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath(strDatabaseFolder & "wsdist_web.mdb") 'This one is for Access 2000/2002  
  strConnect_In = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath(strDatabaseFolder & "wsdist_webU.mdb") 'This one is for Access 2000/2002  
  Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open strConnect
  Set objConn_In = Server.CreateObject("ADODB.Connection")
  objConn_In.Open strConnect_In

%>
