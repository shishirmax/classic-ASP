<%@Language = VBScript %>
<% Option Explicit %>
<%
'--> for debugging
dim fs,f
set fs=Server.CreateObject("Scripting.FileSystemObject")
set f=fs.CreateTextFile("C:\Inetpub\wwwroot\ptfm_debug_report\Report_Test.txt",true)
'--> for debugging
%>
<%
dim objConn
set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = "Provider=SQLOLEDB;Data Source=192.168.2.45;UID=ptfm;PWD=#$@ptfm10;database=shishir;"
objConn.Open

dim objRs
set objRs = Server.CreateObject("ADODB.Recordset")
dim strSql
strSql = "select * from ptfm_users where id ="&"'"&Request.QueryString("id")&"'"

objRs.Open strSql, objConn
%>
<html>
<head>
	<title>
		User Detail
	</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>
<body>
	<div id="logo" align="left">
		<a href="index.asp"><img src="img/Logo1.png" height="50" width="180"></a>
	</div>
	<hr>
<form>
	<div id="tbl" align="center">
	<table border="1" width="800px">
		<tr>
			<td colspan="2" id="tr_even">User Detail</td>
		</tr>

		<tr>
			<td colspan="2">&nbsp</td>
		</tr>

		<tr>
			<td id="tr_even">User Id</td>
			<td id="tvalue"><%=objRs.Fields("userid").Value%></td>
		</tr>
		<tr>
				<td id="tr_even">Contact Person(First/Middle/Last Name)</td>
				<td id="tvalue"><%=objRs.Fields("fname").Value%>/<%=objRs.Fields("mname").Value%>/<%=objRs.Fields("lname").Value%></td>
			</tr>
		<tr>
			<td id="tr_even">Email</td>
			<td id="tvalue"><%=objRs.Fields("email").Value%></td>
		</tr>
		<tr>
			<td colspan="2">&nbsp</td>
		</tr>
		<tr>
			<td colspan="2"><div align="center">
				<a href="edituser.asp?id=<%=objRs.Fields("id").Value%>" id="a_edit">Edit</a>
				&nbsp
				<a href="deleteUser.asp?id=<%=objRs.Fields("id").Value%>" onclick = "return confirm('Confirm delete?')" id="a_edit">Delete</a>
				&nbsp
				<a href="listuser.asp" id="a_edit">Cancel</a></div>
			</td>
		</tr>
	</table>
</div>
<%  'f.WriteLine("userid: "&objRs.Fields("userid").Value)
	'f.WriteLine("fname: "&objRs.Fields("fname").Value)
	'f.WriteLine("mname: "&objRs.Fields("mname").Value)
	'f.WriteLine("lname: "&objRs.Fields("lname").Value)
	'f.WriteLine("email: "&objRs.Fields("email").Value)
%>
	<%
	objRs.Close
	objConn.Close

	%>
</form>
</body>
</html>

