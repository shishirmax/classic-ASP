<html>
<head>
	<title>
		Add Contact Detail
	</title>
	<link rel="stylesheet" type="text/css" href="css/style.css">
	<link href='http://fonts.googleapis.com/css?family=PT+Serif' rel='stylesheet' type='text/css'>
	<script src="js/script.js"></script>
</head>
<body>
	<div id="header">
		<div id="logo" align="left">
		<a href="contact.asp"><img src="img/Logo1.png" height="50" width="180"></a>
	</div>
	<hr>
	</div>
	<div id="body">
	<!--<form name="form" action="contact_submit.asp" method="post" onsubmit="return(validate());">-->
	<form name="form" action="contact_submit.asp" method="post">
		<div id="tbl" align="center">
			<h4>Contact Detail</h4>
			<table border="1" id="table_contact">
				<tr>
					<td id="tr_even">Name</td>
					<td><input type="text" name="name" id="inpt"></td>
				</tr>

				<tr>
					<td id="tr_even">Email</td>
					<td><input type="text" name="email" id="inpt"></td>
				</tr>

				<tr>
					<td id="tr_even">Contact</td>
					<td><input type="text" name="contact" id="inpt"></td>
				</tr>

				<tr>
					<td id="tr_even">Address</td>
					<td><textarea name="address" rows="5" cols="17" id="inpt"></textarea></td>
				</tr>

				<tr>
					<td id="tr_even">State</td>
					<td><input type="text" name="state" id="inpt"></td>
				</tr>

				<tr>
					<td id="tr_even">City</td>
					<td><input type="text" name="city" id="inpt"></td>
				</tr>

				<tr>
					<td id="tr_even">Pincode</td>
					<td><input type="text" name="pincode" id="inpt"></td>
				</tr>

				<tr>
					<td colspan="2">
						<div align="center">
						<input type="submit" name="submit" value="Submit"  id="btn" onclick="return(validate());">
						&nbsp
						<button type="submit" name="cancle"  id="btn" formaction="contact.asp">Cancle</button>
						&nbsp
						<button name="view" id="btn" formaction="view_contact.asp">View</button>
						<div>
					</td>
				</tr>
			</table>

	</form>
	</div>

	<div id="footer">
	</div>
</body>
</html>
