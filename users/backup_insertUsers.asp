<%
'---------------------------- Using DLL instead of this-------------------------------------------------
'dim objConn, strSql
'set objConn = Server.CreateObject("ADODB.Connection")
'objConn.ConnectionString = "Provider=SQLOLEDB;Data Source=192.168.2.45;UID=ptfm;PWD=#$@ptfm10;database=shishir;"
'objConn.Open

'strSql = "sp_ptfmUsers_insert"
'strSql = strSql &"('"&Request.Form("userid")&"','"&Request.Form("password")&"','"&Request.Form("fname")&"','"&Request.Form("mname")&"','"&Request.Form("lname")&"','"&Request.Form("email")&"')"
'objConn.Execute(strSql)

'Response.Redirect ("listUser.asp")

'objConn.Close()
'set objConn = Nothing
'---------------------------- Using DLL instead of this-------------------------------------------------
%>