<HTML>
<BODY BGCOLOR="white">

<H1>aspSmartUpload : Sample 5</H1>
<HR>

<%
'  Variables
'  *********
   Dim mySmartUpload
   Dim item
   Dim value
   Dim file
        
'  Object creation
'  ***************
   Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

'  Upload
'  ******
   mySmartUpload.Upload

'  FILES Collection
'  ****************
   Response.Write("<BR><STRONG>Files Collection</STRONG><BR>")

'  Informations about files
'  ************************
   Response.Write("Number of files =" & mySmartUpload.Files.count &"<BR>")
   Response.Write("Total bytes of files =" & mySmartUpload.Files.TotalBytes &"<BR>")

'  Select each file
'  ****************
   For each file In mySmartUpload.Files
      Response.Write(file.FileName & " (" & file.Size & "bytes)<BR>")
   Next

'  FORM Collection
'  ***************
   Response.Write("<BR><STRONG>Form Collection</STRONG><BR>")

'  Select each item
'  ****************
   For each item In mySmartUpload.Form
   '  Select each value of the current item
   '  *************************************
      For each value In mySmartUpload.Form(item)
         Response.Write(item & " = " & value & "<BR>")
      Next
   Next
%>
</BODY>
</HTML>
