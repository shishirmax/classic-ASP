<%@Language = VBScript %>
<% Option Explicit %>
<html>
<head>
	<title>
		CSV
	</title>
</head>
<body>
	<%
	dim sFileName,objFso, oFiles
	dim objConn
	set objConn = Server.CreateObject("ADODB.Connection")
	objConn.ConnectionString = "Provider=SQLOLEDB;Data Source=192.168.2.45;UID=ptfm;PWD=#$@ptfm10;database=shishir;"
	objConn.Open

	dim objRs
	set objRs = Server.CreateObject("ADODB.Recordset")

	dim strSql
	strSql = "select * from ptfm_users"
	objRs.Open strSql, objConn

	if not objRs.EOF then
	sFileName = "csv/userdetail.csv"

	set objFso = Server.CreateObject("Scripting.FileSystemObject")
	set oFiles = objFso.CreateTextFile(Server.MapPath(sFileName))
	'set oFiles = objFso.CreateTextFile("csv/userdetail.csv")
	oFiles.WriteLine("User Id" & "," & "First Name" & "," & "Middle Name" & "," & "Last Name" & "," & "Email Id")

	while not objRs.EOF
	oFiles.WriteLine(objRs.Fields("userid").Value&","&objRs.Fields("fname").Value&","&objRs.Fields("mname").Value&","&objRs.Fields("lname").Value&","&objRs.Fields("email").Value)
	objRs.MoveNext()
	wend

	oFiles.Close
	set oFiles = Nothing
	set objFso = Nothing

	Response.Write "CSV Completed.<br><a href="&sFileName&">Download</a>"
	end if

	objRs.Close
	objConn.Close

	set objRs = Nothing
	set objConn = Nothing
	%>
</body>
</html>