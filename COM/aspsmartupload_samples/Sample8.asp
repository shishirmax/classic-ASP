<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=utf-8">
</HEAD>
<BODY BGCOLOR="white">
<H1>aspSmartUpload : Sample 8</H1>
<HR>

<%
'  Variables
'  *********
   Dim mySmartUpload
   Dim intCount
        
'  Object creation
'  ***************
   Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

'  Upload
'  ******
   mySmartUpload.CodePage = "utf-8"
   mySmartUpload.Upload

   Response.Write "Text : "
   Response.Write mySmartUpload.Form("TEXT1").values & "<br>"

   Response.Write "File : "
   Response.Write mySmartUpload.Files("FILE1").Filename & "<br>"
   intCount = mySmartUpload.Save("/aspSmartUpload/Upload")

%>
</BODY>
</HTML>