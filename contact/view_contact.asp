<%@ Language = VBScript %>
<html>
<head>
	<title>
		Contact View
	</title>
	<link rel="stylesheet" type="text/css" href="css/style.css">
</head>
<body>
	<div id="logo" align="left">
		<a href="contact.asp"><img src="img/Logo1.png" height="50" width="180"></a>
	</div>
	<hr>
	<%
	dim objConn
	set objConn = Server.CreateObject("ADODB.Connection")
	objConn.ConnectionString = "Provider=SQLOLEDB;Data Source=192.168.2.45;UID=ptfm;PWD=#$@ptfm10;database=shishir;"
	objConn.Open

	dim objRs
	set objRs = Server.CreateObject("ADODB.Recordset")
	dim strSql
	strSql = "select * from contact"
	objRs.Open strSql, objConn
	%>
	<form name="form">
		<div id="tbl" align="center">
		<table id="table_view">
			<tr>
				<td>
					<div align="right">
						<button type="submit" formaction="contact.asp" id="btn">Add</button>
					</div>
				</td>
			</tr>
		</table>
		</div>
		<div id="tbl" align="center">
		<table border="1" id="table_view">
			<tr id="tr-head">
				<th width="50">ID</th>
				<th>Name</th>
				<th>Email</th>
				<th>Contact</th>
				<th>Adress</th>
				<th>Action</th>
			</tr>
			<% do while not objRs.EOF %>
			<tr>
				<td align="center"><%=objRs.Fields("id").Value%></td>
				<td><%=objRS.Fields("name").Value%></td>
				<td><%=objRs.Fields("email").Value%></td>
				<td><%=objRs.Fields("contact").Value%></td>
				<td>
					<%=objRs.Fields("address").Value%><br>
					<%=objRs.Fields("state").Value%><br>
					<%=objRs.Fields("city").Value%><br>
					<%=objRs.Fields("pincode").Value%>
				</td>
				<td align="center"><a href="print_contact.asp?id=<%=objRs.Fields("id").Value%>" target="_blank">Print</a></td>
			</tr>
			<%
			objRs.MoveNext
			loop
			%>
			<%
			objRs.Close
			set objRs = Nothing
			objConn.Close
			set objConn = Nothing
			%>
		</table>
	</div>
	</form>
</body>
</html>