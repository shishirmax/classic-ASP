<html>
<head>
	<title>
		Add User
	</title>
	<!-- CSS Code-->
<link rel="stylesheet" type="text/css" href="css/style.css">
	<!-- CSS Code ends-->

	<!-- JavaScript Code-->
	<script type="text/javascript">
		function validate(){
			if(document.adduser.userid.value=="")
			{
				alert("Enter User ID")
				document.adduser.userid.focus();
				return false;
			}

			if(document.adduser.password.value=="")
				{
					alert("Enter Password")
					document.adduser.password.focus();
					return false;
				}
			if(document.adduser.fname.value=="")
			{
				alert("Enter First Name")
				document.adduser.fname.focus();
				return false;
			}

			if(document.adduser.lname.value=="")
			{
				alert("Enter Last Name")
				document.adduser.lname.focus();
				return false;
			}

			var mailformat = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/; 
	if(document.adduser.email.value.match(mailformat))
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

	function redirect()
	{
		window.location.assign("listUser.asp")
	}
	</script>
<!-- JS Code ends-->
</head>
<body>
	<div id="logo" align="left">
		<a href="index.asp"><img src="img/Logo1.png" height="50" width="180"></a>
	</div>
	<hr>
	<form name="adduser" action="insertUsers.asp" method="POST" onsubmit="return(validate());">
		<div id="tbl" align="center">
		<table border="1">
			<tr>
				<td id="tr_even">User ID</td>
				<td ><input type="text" name="userid"></td>
			</tr>

			<tr>
				<td id="tr_even">Password</td>
				<td id="tvalue"><input type="password" name="password"></td>
			</tr>
			<tr>
				<td id="tr_even">Contact Person(First/Middle/Last Name)</td>
				<td id="tvalue"><input type="text" name="fname">/<input type="text" name="mname">/<input type="text" name="lname"></td>
			</tr>
			<tr>
				<td id="tr_even">Email</td>
				<td id="tvalue"><input type="email" name="email"></td>
			</tr>
			<tr>
				<td colspan="2">&nbsp</td>
			</tr>
			<tr>
				<td colspan="2" id="tbtn">
					<input type="submit" name="submit" value="Submit" id="btn">
					&nbsp
					<input type="button" name="cancel" value="Cancel" onClick="redirect()" id="btn">
				</td>
			</tr>
		</table>
	</div>
	</form>
</body>
</html>
