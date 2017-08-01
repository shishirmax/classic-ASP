<%@Language = VBScript %>
<% option explicit %>
<%
dim objConn
set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = "Provider=SQLOLEDB;Data Source=192.168.2.45;UID=ptfm;PWD=#$@ptfm10;database=shishir;"
objConn.Open

dim objRs
set objRs = Server.CreateObject("ADODB.Recordset")
dim strSql
strSql = "select * from ptfm_users where id ="&"'"&Request.QueryString("id")&"'"
objRs.Open strSql,objConn
%>
<html>
<head>
	<title>
		Edit User
	</title>
<link rel="stylesheet" type="text/css" href="css/style.css">

	<script type="text/javascript">
	function validate()
	{
		if(document.edituser.userid.value=="")
		{
			alert("Enter User ID");
			document.edituser.userid.focus();
			return false;
		}

		if(document.edituser.fname.value=="")
		{
			alert("Enter First Name");
			document.edituser.fname.focus();
			return false;
		}

		if(document.edituser.lname.value=="")
		{
			alert("Enter Last Name");
			document.edituser.fname.focus();
			return false;
		}

		var mailformat = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/; 
		if(document.edituser.email.value.match(mailformat))
		{
		//continue
		}
			else
		{
		alert("Enter valid Email id like(example@domain.com)");
		return false;
		}
		return true;
	}

	function redirect(){
		window.location.assign("listUser.asp")
	}

	</script>
</head>
<body>
	<div id="logo" align="left">
		<a href="index.asp"><img src="img/Logo1.png" height="50" width="180"></a>
	</div>
	<hr>
	<form action="editcmd.asp?id=<%=Request.QueryString("id")%>" name="edituser" method="post" onsubmit="return(validate());">
		<div if="tbl" align="center">
	<table border="1" width="800px">

		<tr>
			<td id="tr_even">ID</td>
			<td><input type="hidden" name="id" value="<%=objRs.Fields("id").Value%>"></td>
		</tr>
		<tr>
				<td id="tr_even">User ID</td>
				<td id="tvalue"><input type="text" name="userid" value="<%=objRs.Fields("userid").Value%>"></td>
			</tr>

			<tr>
				<td id="tr_even">Contact Person(First/Middle/Last Name)</td>
				<td id="tvalue"><input type="text" name="fname" value="<%=objRs.Fields("fname").Value%>">/<input type="text" name="mname" value="<%=objRs.Fields("mname").Value%>">/<input type="text" name="lname" value="<%=objRs.Fields("lname").Value%>"></td>
			</tr>
			<tr>
				<td id="tr_even">Email</td>
				<td id="tvalue"><input type="email" name="email" value="<%=objRs.Fields("email").Value%>"></td>
			</tr>
			<tr>
				<td colspan="2">&nbsp</td>
			</tr>
			<tr>
				<td colspan="2" id="tbtn">
					<input type="submit" name="submit" value="Submit" id="btn">
					&nbsp
					<input type="button" name="cancle" value="Cancel" onclick="redirect()" id="btn">
				</td>
			</tr>

		</table>
	</div>

<%
objRs.Close
set objRs = Nothing

objConn.Close
set objConn = Nothing
%>
	</form>
</body>
</html>
