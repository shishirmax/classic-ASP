<%@ Language = VBScript %>
<% Option Explicit %>
<%
dim objUpload
dim intCount

set objUpload = Server.CreateObject("aspsmartupload.SmartUpload")
objUpload.Upload 
intCount = objUpload.Save(Server.MapPath("files/"))
Response.Write(intCount & " File uploaded.")
Response.Write("<br>")
Response.Write("<a href='upload.asp'>Upload More</a>")

set objUpload = Nothing
%>