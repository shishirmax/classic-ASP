<html>
<head>
	<title>
		Button Test
	</title>
	<script type="text/javascript">
	function redirect(){
		window.location.assign("listUser.asp")
	}
	</script>
</head>
<body>
	<p> Test of JS on Button </p>
	<br>
	<form name="form" method="post">
	<input type ="button" name="cancel" onClick="redirect()" value="Cancel">

</form>
</body>
</html>
