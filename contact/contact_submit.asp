<%@ Language = VBscript %>

<%
dim objConn,strSql
set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = "Provider=SQLOLEDB;Data Source=192.168.2.45;UID=ptfm;PWD=#$@ptfm10;database=shishir;"
objConn.Open

strSql = "sp_InsertContact"
strSql = strSql & "('"&Request.Form("name")&"','"&Request.Form("email")&"','"&Request.Form("contact")&"','"&Request.Form("address")&"','"&Request.Form("state")&"','"&Request.Form("city")&"','"&Request.Form("pincode")&"')"
objConn.Execute(strSql)

dim objMail, strMsg
set objMail = Server.CreateObject("CDONTS.NewMail")

strName = Request.Form("name")
strTo = Request.Form("email")

strMsg = "<!DOCTYPE HTML PUBLIC ""-//IETF//DTD HTML//EN"">"
strMsg = strMsg & "<html>"
strMsg = strMsg & "<head>"
strMsg = strMsg & "<meta http-equiv=""Content-Type"""
strMsg = strMsg & "content=""text/html; charset=iso-8859-1"">"
strMsg = strMsg & "<meta name=""GENERATOR"" content=""Microsoft FrontPage 2.0"">"
strMsg = strMsg & "</head>"
strMsg = strMsg & "<body>"
strMsg = "Hi, " & strName
strMsg = strMsg & "<br>Your Contact has been added to our database."
strMsg = strMsg & "</body>"
strMsg = strMsg & "</html>"

objMail.From = "Webmaster <webmaster@contact.com>"
objMail.Value("Reply-To") = "care@contact.com"
objMail.To = strTo
objMail.Bcc = "shishir@contata.co.in"
objMail.Subject = "Contact Added"
objMail.MailFormat = 0
objMail.BodyFormat = 0
objMail.Importance = 2
objMail.Body = strMsg

objMail.Send

Response.Redirect("view_contact.asp")
objConn.Close
set objConn = Nothing
set objMail = Nothing

%>