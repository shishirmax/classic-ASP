<%@ Language = VBScript %>

<%
dim fs,f
set fs = Server.CreateObject("Scripting.FileSystemObject")
set f = fs.CreateTextFile("C:\Inetpub\wwwroot\ptfm_debug_report\debug.txt",true)
%>

<%
dim objConn, strSql
set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = "Provider=SQLOLEDB;Data Source=192.168.2.45;UID=ptfm;PWD=#$@ptfm10;database=shishir;"
objConn.Open

strSql = "sp_InsSubscription"
strSql = strSql & "('"&Request.Form("name")&"','"&Request.Form("email")&"','"&Request.Form("contact")&"')"
objConn.Execute(strSql)

dim objMail,strMsg
set objMail = Server.CreateObject("CDONTS.NewMail")


strTo = Request.Form("email")
strMsg = "Hi, "
strMsg = strMsg & "Thank you for subscribing to our website."

objMail.From = "Webmaster <webmaster@samplemail.com>"
objMail.Value("Reply-To") = "reply@samplemail.com"
objMail.To = strTo
objMail.Bcc = "shishir@contata.co.in"
objMail.Subject = "Subscription Confirmation"
objMail.Importance = 2
objMail.Body = strMsg

objMail.Send
Response.Write("Mail Sent....")
Response.Write("<br>")

objConn.Close
set objConn = Nothing
set objMail = Nothing
%>
<a href="subscription.asp">Click</a>&nbsp Here