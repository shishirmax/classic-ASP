<%@Language = VBScript %>
<% option explicit %>
<%
dim objConn
set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = "Provider=SQLOLEDB;Data Source=192.168.2.45;UID=ptfm;PWD=#$@ptfm10;database=shishir;"
objConn.Open

dim strSql
strSql = "sp_update_user"
strSql = strSql &"('"&Request.Form("id")&"','"&Request.Form("userid")&"','"&Request.Form("fname")&"','"&Request.Form("mname")&"','"&Request.Form("lname")&"','"&Request.Form("email")&"')"
objConn.Execute(strSql)
response.redirect ("listUser.asp")

objConn.Close
set objConn = Nothing
%>