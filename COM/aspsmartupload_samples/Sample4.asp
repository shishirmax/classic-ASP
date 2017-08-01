<HTML>
<BODY BGCOLOR="white">

<H1>aspSmartUpload : Sample 4</H1>
<HR>

<%
'  Variables
'  *********
   Dim mySmartUpload
   Dim file
   Dim oConn
   Dim oRs
   Dim intCount
   intCount=0
        
'  Object creation
'  ***************
   Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

'  Upload
'  ******
   mySmartUpload.Upload

'  Connect to the DB
'  *****************
   Set oConn = Server.CreateObject("ADODB.Connection")
   curDir = Server.MapPath("\scripts\aspSmartUpload\Sample.mdb")
   oConn.Open "DBQ="& curDir &";Driver={Microsoft Access Driver (*.mdb)};DriverId=25;FIL=MS Access;"

'  Open a recordset
'  ****************
   strSQL = "SELECT FILENAME,FILE FROM TFILES"

   Set oRs = Server.CreateObject("ADODB.recordset")
   Set oRs.ActiveConnection = oConn
   oRs.Source = strSQL
   oRs.LockType = 3
   oRs.Open

'  Select each file
'  ****************
   For each file In mySmartUpload.Files
   '  Only if the file exist
   '  **********************
      If not file.IsMissing Then

      '  Add the current file in a DB field
      '  **********************************
         oRs.AddNew
         file.FileToField oRs.Fields("FILE")
		 oRs("FILENAME") = file.FileName
         oRs.Update
         intCount = intCount + 1
      End If
   Next

'  Display the number of files uploaded
'  ************************************
   Response.Write(intCount & " file(s) uploaded.<BR>")

'  Destruction
'  ***********
   oRs.Close
   oConn.Close
   Set oRs = Nothing 
   Set oConn = Nothing 
%>
</BODY>
</HTML>
