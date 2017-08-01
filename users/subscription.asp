<html>
<head>
	<title>
		Subscription Form
	</title>
</head>
<body>
	<form name="form" method="post" action="subscription_submit.asp">
		<table border="1">
			<tr>
				<td colspan="2"><div align="center">Subscription Form</div></td>
			</tr>

			<tr>
				<td>Name</td>
				<td><input type="text" name="name" required></td>
			</tr>

			<tr>
				<td>Email</td>
				<td><input type="email" name="email" required></td>
			</tr>

			<tr>
				<td>Contact</td>
				<td><input type="text" name="contact" required></td>
			</tr>

			<tr>
				<td colspan="2">
					<div align="center"><input type="submit" name="submit" value="Subscribe"></div>
				</td>
			</tr>
		</table>
	</form>
</body>
</html>