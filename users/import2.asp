<% Option Explicit %>
<html>
<head>
	<title>
	</title>
</head>
<body>
	<%
	dim objFso, oInStream, sRows, arrRows
	dim objConn, strSql, objExec
	dim sFileName
	dim objUpload

	set objUpload = Server.CreateObject("aspsmartupload.SmartUpload")
	objUpload.Upload
	sFileName = objUpload.Files("file1").FileName 
	if sFileName <> "" then
	objUpload.Files("file1").SaveAs(Server.MapPath("files/" & sFileName))

	set objFso = Server.CreateObject("Scripting.FileSystemObject")
	if not objFso.FileExists(Server.MapPath("files/" & sFileName)) then
	Response.Write("File Not Found.")
	else

	set oInStream = objFso.OpenTextFile(Server.MapPath("files/"&sFileName),1,False)

	set objConn = Server.CreateObject("ADODB.Connection")
	objConn.ConnectionString = "Provider=SQLOLEDB;Data Source=192.168.2.45;UID=ptfm;PWD=#$@ptfm10;database=shishir;"
	objConn.Open

	do until oInStream.AtEndOfStream
	sRows = oInStream.readLine
	arrRows = Split(sRows,",")

	strSql = ""
	strSql = strSql & "insert into ptfm_users"
	strSql = strSql & "(userid,password,fname,mname,lname,email"
	strSql = strSql & "Values"
	strSql = strSql & "('"&arrRows(0)& "','"&arrRows(1)&"','"&arrRows(2)&"','"&arrRows(3)&"','"&arrRows(4)&"','"&arrRows(5)&"')"
	set objExec = objConn.Execute(strSql)

	set objExec = Nothing
	Loop

	oInStream.Close
	objConn.Close
	set oInStream = Nothing
	set objConn = Nothing

	end if

	Response.Write("CSV import sucess.")
	end if
	%>
</body>
</html>
