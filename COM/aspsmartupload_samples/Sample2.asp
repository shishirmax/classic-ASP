<HTML>
<BODY BGCOLOR="white">

<H1>aspSmartUpload : Sample 2</H1>
<HR>

<%
'  Variables
'  *********
   Dim mySmartUpload
   Dim file
   Dim intCount
   intCount=0
        
'  Object creation
'  ***************
   Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

'  Upload
'  ******
   mySmartUpload.Upload

'  Select each file
'  ****************
   For each file In mySmartUpload.Files
   '  Only if the file exist
   '  **********************
      If not file.IsMissing Then
      '  Save the files with his original names in a virtual path of the web server
      '  ****************************************************************************
         file.SaveAs("/aspSmartUpload/Upload/" & file.FileName)
         ' sample with a physical path 
         ' file.SaveAs("c:\temp\" & file.FileName)

      '  Display the properties of the current file
      '  ******************************************
         Response.Write("Name = " & file.Name & "<BR>")
         Response.Write("Size = " & file.Size & "<BR>")
         Response.Write("FileName = " & file.FileName & "<BR>")
         Response.Write("FileExt = " & file.FileExt & "<BR>")
         Response.Write("FilePathName = " & file.FilePathName & "<BR>")
         Response.Write("ContentType = " & file.ContentType & "<BR>")
         Response.Write("ContentDisp = " & file.ContentDisp & "<BR>")
         Response.Write("TypeMIME = " & file.TypeMIME & "<BR>")
         Response.Write("SubTypeMIME = " & file.SubTypeMIME & "<BR>")
         intCount = intCount + 1
      End If
   Next

'  Display the number of files which could be uploaded
'  ***************************************************
   Response.Write("<BR>" & mySmartUpload.Files.Count & " files could be uploaded.<BR>")

'  Display the number of files uploaded
'  ************************************
   Response.Write(intCount & " file(s) uploaded.<BR>")
%>
</BODY>
</HTML>