<%@Language = VBScript %>
<% option explicit %>
<%
dim objConn
set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = "Provider=SQLOLEDB;Data Source=192.168.2.45;UID=ptfm;PWD=#$@ptfm10;database=shishir;"
objConn.Open

dim strSql
strSql = "delete from ptfm_users where id = "&"'"&Request.QueryString("id")&"'"
objConn.Execute(strSql)

Response.Redirect("listuser.asp")
objConn.Close()
set objConn = Nothing
%>