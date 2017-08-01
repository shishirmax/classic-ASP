<HTML>
<BODY BGCOLOR="white">

<H1>aspSmartUpload : Sample 3</H1>
<HR>

<%
   On Error Resume Next

'  Variables
'  *********
   Dim mySmartUpload
   Dim intCount
        
'  Object creation
'  ***************
   Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

'  Only allow txt or htm files
'  ***************************
   mySmartUpload.AllowedFilesList = "htm,txt"

'  DeniedFilesList can also be used :
   ' Allow all files except exe, bat and asp
   ' ***************************************
   ' mySmartUpload.DeniedFilesList = "exe,bat,asp"

'  Deny physical path
'  *******************
   mySmartUpload.DenyPhysicalPath = True

'  Only allow files smaller than 50000 bytes
'  *****************************************
   mySmartUpload.MaxFileSize = 50000

'  Deny upload if the total fila size is greater than 200000 bytes
'  ***************************************************************
   mySmartUpload.TotalMaxFileSize = 200000

'  Upload
'  ******
   mySmartUpload.Upload

'  Save the files with their original names in a virtual path of the web server
'  ****************************************************************************
   intCount = mySmartUpload.Save("/aspSmartUpload/Upload")
   ' sample with a physical path 
   ' intCount = mySmartUpload.Save("c:\temp\")

'  Trap errors
'  ***********
   If Err Then
      Response.Write("<b>Wrong selection : </b>" & Err.description)
   Else
   '  Display the number of files uploaded
   '  ************************************
      Response.Write(intCount & " file(s) uploaded.")
   End If
%>
</BODY>
</HTML>
