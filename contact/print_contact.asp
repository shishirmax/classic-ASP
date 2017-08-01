<%@ Language = VBScript %>
<%
dim objConn
set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = "Provider=SQLOLEDB;Data Source=192.168.2.45;UID=ptfm;PWD=#$@ptfm10;database=shishir;"
objConn.Open

dim objRs
set objRs = Server.CreateObject("ADODB.Recordset")
dim strSql
strSql = "select * from contact where id = "&"'"&Request.QueryString("id")&"'"
objRs.Open strSql,objConn
%>
<%
Function GetFormattedDate
  strDate = CDate(Date)
  'Response.Write(strDate)
  strDay = DatePart("d", strDate)
  strMonth = DatePart("m", strDate)
  strYear = DatePart("yyyy", strDate)
  If strDay < 10 Then
    strDay = "0" & strDay
  End If
  If strMonth < 10 Then
    strMonth = "0" & strMonth
  End If
  GetFormattedDate = strDay & "-" & strMonth & "-" & strYear
End Function
%>
<html>
<head>
	<title>
		Print Contact
	</title>
	<script type="text/javascript">
	function PrintPage()
	{
		window.print();
	}
	</script>
	<link rel="stylesheet" type="text/css" href="css/style.css">
	<style type="text/css">
	@page {
  	size: A4 landscape;
	}
	</style>
</head>
<body>
	<div id="logo" align="left">
		<a href="contact.asp"><img src="img/Logo1.png" height="50" width="180"></a>
	</div>
	<hr>
	<div id="date" align="right">
		<b>Printed on:</b> <%=GetFormattedDate%>
	<form>
		<div id="tbl" align="center">
		<table>
			<tr>
				<td id="tr_even" width="100">Name</td>
				<td><%=objRs.Fields("name").Value%></td>
			</tr>
			<tr>
				<td id="tr_even" width="100">Email</td>
				<td><%=objRs.Fields("email").Value%></td>
			</tr>
			<tr>
				<td id="tr_even" width="100">Contact</td>
				<td><%=objRs.Fields("contact").Value%></td>
			</tr>
			<tr>
				<td id="tr_even" width="100">Address</td>
				<td>
					<%=objRs.Fields("address").Value%><br>
					<%=objRs.Fields("state").Value%><br>
					<%=objRs.Fields("city").Value%><br>
					<%=objRs.Fields("pincode").Value%>
				</td>
			</tr>
		</table>
	</div>
		<div id="img" align="center">
			<br>
			<img src="img/print.png" height="20" width="20" onclick="javascript:PrintPage()">
		</div>
	</form>
</body>
</html>