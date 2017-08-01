<%@Language=VBScript%>
<% Option Explicit %>
<!--#include virtual="/adovbs.inc"-->
<html>
<head>
	<title>
		User List
	</title>
	<link rel="stylesheet" type="text/css" href="css/style.css">
	<script type="text/javascript">
	function clearfilter()
	{
		window.document.form.userid.value="";
	}
	</script>
</head>
<body>
	<div id="logo" align="left">
		<a href="index.asp"><img src="img/Logo1.png" height="50" width="180"></a>
	</div>
	<hr>
	<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="get" name="form">
		<div id="tbl" align="center">
		<table border="1" width="800px">
			<tr id="tr_even">
				<td colspan="4"><a href="listuser.asp" id="a_edit">User List</a></td>
			</tr>

			<tr>
				<td colspan="4">&nbsp</td>
			</tr>

			<tr id="tr_even">
				<td colspan="3"><div align="left">Filter</div></td>
				<td colspan="1"><div align="right"><a HREF="javascript:clearfilter();" id="a_edit">Clear Filter</a></div></td>
			</tr>

			<tr>
				<td><div align="left" id="tr_even">User ID</div></td>
				<td colspan="2"><input type="text" name="userid"></td>
				<td><div align="right"><input type="submit" name="go" value="Go" id="btn"></td>
			</tr>

			<tr>
				<td colspan="4">&nbsp</td>
			</tr>


			<tr>
				<td colspan="4"><div align="right"><button type="submit" formaction="AddUsers.asp" id="btn">Add</button>&nbsp<button type="submit" formaction="export_csv.asp" id="btn">Export</button></div></td>
			</tr>
		</table>
	</div>
	</form>

			<%
			'--> for debugging
			dim fs,f
			set fs=Server.CreateObject("Scripting.FileSystemObject")
			set f=fs.CreateTextFile("C:\Inetpub\wwwroot\ptfm_test\users\Report_Test.txt",true)
			'--> for debugging

			dim objConn
			set objConn = Server.CreateObject("ADODB.Connection")
			objConn.ConnectionString = "Provider=SQLOLEDB;Data Source=SHISHIR\SQLEXPRESS;UID=sa;PWD=sysadmin;database=shishir;"
			objConn.Open

			dim objRs
			set objRs = Server.CreateObject("ADODB.Recordset")
			objRs.CursorLocation = adUseClient

			dim strSql
			strSql = "select * from ptfm_users where userid like '%"&request.querystring("userid")&"%'"
			objRs.Open strSql, objConn
			objRs.Sort = "userid"
			'f.WriteLine("objConn"&objConn)
			'f.WriteLine("objRs: "&objRs)
			%>
			<div id="tbl" align="center">
			<table border="1" width="800px">
			<tr id="tr-head">
				<th width="100">UserId</th>
				<th>First Name</th>
				<th>Email</th>
				<th>Action</th>
			</tr>
			<% do while not objRs.EOF %>
			<tr>
				<td><a href="viewUser.asp?id=<%=objRs.Fields("id").Value%>"><%=objRs.Fields("userid").Value%></a></td>
				<td><%=objRs.Fields("fname").Value%></td>
				<td><%=objRs.Fields("email").Value%></td>
				<td><div align="center"><a href="edituser.asp?id=<%=objRs.Fields("id").Value%>" id="a_edit">Edit</a></div></td>
				<!--<td><a href="edituser.asp?id=<%=objRs.Fields("id").Value%>">
					<div align="center"><img src="img/view_edit.gif"></div>
				</a></td>-->
			</tr>
			<%
			objRs.MoveNext
			loop
			%>
			<% if objRs.BOF and objRs.EOF = true then %>
			<tr>
				<td colspan="4"><div align="center"><b>No Record Found</b></div></td>
			</tr>
			<% end if %>
			<%
			'f.WriteLine("strSql : "&strSql)
			objRs.Close
			set objRs = Nothing
			objConn.Close
			set objConn = Nothing
			%>

		</table>
	</div>
</body>
</html>
