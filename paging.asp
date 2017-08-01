<%@ Language = VBScript %>
<% Option Explicit %>
<!--#include virtual="/ADOVBS.inc"-->
<%
Const NumPerPage = 10
Dim CurPage
If Request.QueryString("CurPage") = "" then
	CurPage = 1
	Else
	CurPage = Request.QueryString("CurPage")
End If

Dim Conn
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionString = "Provider=SQLOLEDB;Data Source=192.168.2.45;UID=ptfm;PWD=#$@ptfm10;database=PTFMInvoicing;"
Conn.Open

Dim rs 
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = adUseClient
rs.CacheSize = NumPerPage

Dim strSQL
strSQL = "select cl_clientnumber,cl_name from client where cl_firmcode='wgsb' order by cl_id"
rs.Open strSQL,Conn

rs.MoveFirst
rs.PageSize = NumPerPage

Dim TotalPages
TotalPages = rs.PageCount
rs.AbsolutePage = CurPage

Dim Count
%>
<html>
<head>
	<!--<title>Paging</title>-->
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<link rel="stylesheet" type="text/css" href="css/bootstrap.min.css">
	<script type="text/javascript" src="js/jquery-1.11.3.min.js"></script>
	<script type="text/javascript" src="js/bootstrap.min.js"></script>
	
</head>
<body>
	<div class="container">
		<br>
	<!--<b>Client No  -  Client Name</b><br>-->
	<table class="table table-bordered table-striped table-condensed " style="width: auto;"> 
		<thead>
			<tr>
				<th>Client Number</th>
				<th>Client Name</th>
			</tr>
		</thead>
		<tbody>
	<%
	Count = 0
	Do While Not rs.EOF and Count < rs.PageSize
	'Response.Write(rs("cl_clientnumber") & " - " & rs("cl_name") & "<br>")
	%>
	<tr>
		<td><%= rs("cl_clientnumber")%></td>
		<td><%= rs("cl_name")%></td>
	</tr>
	<%
	Count = Count + 1
	rs.MoveNext
	Loop
	%>
</tbody>
</table>
	<%
	Response.Write("Page " & CurPage & " of " & TotalPages & "<p>")
	if CurPage > 1 then
	'Show the prev button
	Response.Write("<input type=button value=prev onclick=""document.location.href='paging.asp?curpage="&curpage-1&"';"">")
	End if
	'Response.Write("Page " & CurPage & " of " & TotalPages & "<p>")
	if CInt(CurPage)<> CInt(TotalPages) then
	'Show the next button
	Response.Write("<input type=button value=next onclick=""document.location.href='paging.asp?curpage=" & curpage + 1 &"';"">")
	end If
	%>
</div>
</body>
</html>