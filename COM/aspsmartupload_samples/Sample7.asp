<%
'  Variables
'  *********
   Dim mySmartUpload
   Dim oConn
   Dim strSQL
   Dim oRs
        
'  Object creation
'  ***************
   Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

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
   oRs.Open


'  DownloadField
'  *************
   mySmartUpload.DownloadField(oRs("FILE"))
   ' samples with optionnals
   ' Call mySmartUpload.DownloadFile(oRs("FILE"), "application/x-zip-compressed", "download.zip")
   ' Call mySmartUpload.DownloadFile(oRs("FILE"), "application/x-zip-compressed", oRs("FILENAME"))

'  Destruction
'  ***********
   oRs.Close
   oConn.Close
   Set oRs = Nothing 
   Set oConn = Nothing 
%>