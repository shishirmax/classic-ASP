<%@ Language = VBScript %>
<% Option Explicit %>
<html>
<head>
	<title>
		Export DB CSV
	</title>
</head>
<bodu>
	<%
	dim objFso, oFiles
	dim objConn, objRs, strSql
	dim sFileName

	set objConn = Server.CreateObject("ADODB.Connection")
	objConn.ConnectionString = "Provider=SQLOLEDB;Data Source=192.168.2.45;UID=ptfm;PWD=#$@ptfm10;database=shishir;"
	objConn.Open

	set objRs = Server.CreateObject("ADODB.Recordset")
	strSql = "select book_name,book_author,book_isbn,book_publisher,book_entry_date from book"
	objRs.Open strSql, objConn

	if not objRs.EOF then

	sFileName = "book_detail.csv"

	set objFso = Server.CreateObject("Scripting.FileSystemObject")

	set oFiles = objFso.CreateTextFile(Server.MapPath(sFileName))

	while not objRs.EOF
	oFiles.WriteLine(objRs.Fields("book_name").Value&","&objRs.Fields("book_author").Value&","&objRs.Fields("book_isbn").Value&","&objRs.Fields("book_publisher").Value&","&objRs.Fields("book_entry_date").Value)
	objRs.MoveNext()
	wend

	oFiles.Close()
	set oFiles = Nothing
	set objFso = Nothing

	Response.write "Generate CSV Completed..<br><a href="&sFileName&">Download</a>"
	end if

	objRs.Close()
	objConn.Close()

	set objRs = Nothing
	set objConn = Nothing
	%>
</body>
</html>