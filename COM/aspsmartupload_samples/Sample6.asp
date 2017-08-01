<%
'  Variables
'  *********
   Dim mySmartUpload
        
'  Object creation
'  ***************
   Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

'  Download
'  ********
   mySmartUpload.DownloadFile("/aspSmartUpload/Upload/sample.zip")
   ' sample with a physical path
   ' mySmartUpload.DownloadFile("c:\temp\sample.zip")
   ' sample with optionnals
   ' Call mySmartUpload.DownloadFile("/aspSmartUpload/Upload/sample.zip","application/x-zip-compressed","downloaded.zip")
%>