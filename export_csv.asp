<%@ Language = VBScript %>
<% Option Explicit %>
<html>
<head>
	<title>
		Export CSV
	</title>
</head>
<body>
	<%
	dim objFSO, oFiles
	dim sFileName

	sFileName = "employee.csv"

	set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	set oFiles = objFSO.CreateTextFile(Server.MapPath(sFileName),2)
	oFiles.WriteLine("EMP012,John Doe,john@infocom.com")
	oFiles.WriteLine("EMP013,Shishir Max,shishirm@infocom.com")
	oFiles.WriteLine("EMP014,Barry Adams, barry@infocom.com")

	oFiles.Close()
	set oFiles = Nothing
	set objFSO = Nothing

	Response.Write"Generate CSV Completed..<br><a href="&sFileName&">Download</a>"
	%>
</body>
</html>