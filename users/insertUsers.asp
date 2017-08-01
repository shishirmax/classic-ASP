<%@Language = VBScript %>
<%
Dim userid, password, fname, mname, lname, email

userid = Request.Form("userid")
password = Request.Form("password")
fname = Request.Form("fname")
mname = Request.Form("mname")
lname = Request.Form("lname")
email = Request.Form("email")
%>
<%
Dim objConn
Set objConn = Server.CreateObject("SampleApplication.App")
Call objConn.InsertUser(userid,password,fname,mname,lname,email)

Response.Redirect("listUser.asp")
Set objConn = Nothing
%>
