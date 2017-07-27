<%@ Language = VBScript %>
<%server.ScriptTimeout = 800%>
<!--METADATA TYPE="typelib" uuid="00000205-0000-0010-8000-00AA006D2EA4" -->
<!--#include file="../Common/mainHeader.asp" -->
<!--#include file="../Shared/SharedPost.asp" -->
<%
Call CheckSession()
If Request.QueryString("tabIndex") <> "" Then
		Session("tabIndex") = Request.QueryString("tabIndex")
End If

firmcode=RTRIM(LTRIM(Session("FirmCode")))
Loginid = Session("LoginID")

Dim batch
batch = 43'Request.QueryString("BATCH")
%>
<html>
	<head>
		<title>
			PTFM Foreign Associates
		</title>
		<meta charset="urf-8">
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<link href="../intel/styles.css" rel="stylesheet" type="text/css">
		<link href="../Styles/styles.css" rel="stylesheet" type="text/css">
		<script language="JavaScript" src="../script/preloadImages.js"> </script>
		<script language="JavaScript" src="../Script/mm_menu.js"></script> 
		<script language="JavaScript" src="../Script/params.js"></script> 
		<script language="JavaScript" src="../script/InvoicenavBarRenderer.js"> </script>
	</head>
	<body style="padding-top: 0px;" leftmargin="0" topmargin="0">
	 <% 	  	   
	  if UCase(trim(Session("firmCode"))) = "WGSB" then
			strReturnPath = "dashboard.asp"
				call OrgHeader()						 
	  End if		
	 %>
	 <%
	 'Dim ObjRs,Conn
	 'Set ObjRpt = Server.CreateObject("PTFMInvoicing.DBHelper")
	 'Set Conn = Server.CreateObject("ADODB.CONNECTION")
	 'ConnStr = "Provider=SQLOLEDB;Data Source=SHISHIR\SQLEXPRESS;database=PTFMInvoicingAMEX;UID=sa;PASSWORD=sysadmin;"
	 'Conn.ConnectionString = ConnStr
	 'Conn.Open
	 'Set ObjRs = Server.CreateObject("ADODB.Recordset")
	 'sql = "exec SP_MasterCardUpdate '"&firmcode&"'"
	 'ObjRs.Open sql, Conn


	 'Set ObjRs = Conn.execute(sql)

'**** SP Connection String ******
	 'Dim ObjRpt,Conn,ObjRs,ObjFso,OFiles
	 'Set ObjRpt = Server.CreateObject("PTFMInvoicing.DBHelper")

	 'Set Conn = Server.CreateObject("ADODB.CONNECTION")
	 'ConnStr = "Provider=SQLOLEDB;Data Source=SHISHIR\SQLEXPRESS;database=PTFMInvoicingAMEX;UID=sa;PASSWORD=sysadmin;"
	 'Conn.ConnectionString = ConnStr
	 'Conn.Open
	 'flag = 1
	 'Set ObjRs = Server.CreateObject("ADODB.Recordset")	 
	 'Set ObjCmd = Server.CreateObject("ADODB.Command")
	 'ObjCmd.ActiveConnection = Conn
	 'ObjCmd.CommandText = "sp_MasterCardUpdate"
	 'ObjCmd.CommandType = adCmdStoredProc

	 'sql = "exec SP_MasterCardUpdate '"&firmcode&"'"
	 'sql = "exec SP_GETMCDATA 23"
	 
	 'set ObjRs = Conn.Execute(sql)
	 'ObjRs.Open ObjCmd
	 'ObjRs.Open sql, Conn

	 '******************************************* Connection*************************
	 Dim ObjRpt
	 Set ObjRpt = Server.CreateObject("PTFMInvoicing.DBHelper")
	 'Set ObjRs = Server.CreateObject("ADODB.Recordset")

	 	Dim arrParam(0)
	 	Dim arrParam1(3)
	 	Dim ObjRs

	 	arrParam1(0) = "FirmCode"
	 	arrParam1(1) = 129
	 	arrParam1(2) = 50
	 	arrParam1(3) = firmcode
	 	arrParam(0) = arrParam1

	 	Set ObjRs = ObjRpt.RunSPReturnRS("SP_MasterCardUpdate",arrParam,"")

	 If Not ObjRs.EOF Then
	 	strGroupBy = "MasterCard-"
	 	FileName = strGroupBy & cSTR(Month(DATE)) & CSTR(Day(DATE))& CSTR(YEAR(date))
	 	FileID = Server.MapPath("../Amex/file/" & FileName & ".csv")

	 	Set ObjFso = CreateObject("Scripting.FileSystemObject")
	 	Set OFiles = ObjFso.OpenTextFile(FileID, 2, true, 0)

	 	'-------------------Writing Headers
	 	OFiles.WriteLine("REQUESTING_CONTROL_ACCOUNT"&","&"BASIC_CONTROL_ACCOUNT"&","&"CARDMEMBER_ACCOUNT_NUMBER"&","&"SE_NUMBER"&","&"ROC_ID"&","&"DB\CR_INDICATOR"&","&"TRANSACTION_TYPE_CODE"&","&"FINANCIAL_CATEGORY"&","&"BATCH_NUMBER"&","&"DATE_OF_CHARGE"&","&"LOCAL_CURRENCY_AMOUNT"&","&"CURRENCY_CODE"&","&"CAPTURE_DATE"&","&"PROCESS_DATE"&","&"BILLING_DATE"&","&"BILLING_AMOUNT"&","&"SALES_TAX_AMOUNT"&","&"TIP_AMOUNT"&","&"CARDMEMBER_NAME"&","&"SPECIAL_BILL_IND"&","&"ORIGINATING_BCA"&","&"ORIGINATING_ACCOUNT_NUMBER"&","&"CM_REFERENCE_NUMBER"&","&"SUPPLIER_REFERENCE_NUMBER"&","&"SHIP_TO_ZIP"&","&"SIC_CODE"&","&"COST_CENTER"&","&"EMPLOYEE_ID"&","&"SOCIAL_SECURITY_#"&","&"UNIVERSAL#"&","&"STREET"&","&"CITY"&","&"STATE"&","&"ZIP+4"&","&"TRANS_LIMIT"&","&"MONTHLY_LIMIT"&","&"EXPOSURE_LIMIT"&","&"REV_CODE"&","&"COMPANY_NAME"&","&"CHARGE_DESCRIPTION_LINE1"&","&"CHARGE_DESCRIPTION_LINE2"&","&"CHARGE_DESCRIPTION_LINE3"&","&"CHARGE_DESCRIPTION_LINE4"&","&"CAR_RENTAL_CUSTOMER_NAME"&","&"CAR_RENTAL_CITY"&","&"CAR_RENTAL_STATE"&","&"CAR_RENTAL_DATE"&","&"CAR_RETURN_CITY"&","&"CAR_RETURN_STATE"&","&"CAR_RETURN_DATE"&","&"CAR_RENTAL_DAYS"&","&"HOTEL_ARRIVAL_DATE"&","&"HOTEL_CITY"&","&"HOTEL_STATE"&","&"HOTEL_DEPART_DATE"&","&"HOTEL_STAY_DURATION"&","&"HOTEL_ROOM_RATE"&","&"AIR_AGENCY_NUMBER"&","&"AIR_TICKET_ISSUER"&","&"AIR_CLASS_OF_SERVICE"&","&"AIR_CARRIER_CODE"&","&"AIR_ROUTING"&","&"AIR_DEPARTURE_DATE"&","&"AIR_PASSENGER_NAME"&","&"TELE_DATE_OF_CALL"&","&"TELE_FROM_CITY"&","&"TELE_FROM_STATE"&","&"TELE_CALL_LENGTH"&","&"TELE_REFERENCE_NUMBER"&","&"TELE_TIME_OF_CALL"&","&"TELE_TO_NUMBER"&","&"INDUSTRY_CODE"&","&"SEQUENCE_NUMBER"&","&"MERCATOR_KEY"&","&"FEE_ALLOCATOR_IND"&","&"") 	

	 	While Not ObjRs.EOF
	 		
	 		OFiles.WriteLine(ObjRs("REQUESTING_CONTROL_ACCOUNT")&","&ObjRs.Fields("BASIC_CONTROL_ACCOUNT").Value&","&ObjRs.Fields("CARDMEMBER_ACCOUNT_NUMBER").Value&","&ObjRs.Fields("SE_NUMBER").Value&","&ObjRs.Fields("ROC_ID").Value&","&ObjRs.Fields("DB_CR_INDICATOR").Value&","&ObjRs.Fields("TRANSACTION_TYPE_CODE").Value&","&ObjRs.Fields("FINANCIAL_CATEGORY").Value&","&ObjRs.Fields("BATCH_NUMBER").Value&","&ObjRs.Fields("DATE_OF_CHARGE").Value&","&ObjRs.Fields("LOCAL_CURRENCY_AMOUNT").Value&","&ObjRs.Fields("CURRENCY_CODE").Value&","&ObjRs.Fields("CAPTURE_DATE").Value&","&ObjRs.Fields("PROCESS_DATE").Value&","&ObjRs.Fields("BILLING_DATE").Value&","&ObjRs.Fields("BILLING_AMOUNT").Value&","&ObjRs.Fields("SALES_TAX_AMOUNT").Value&","&ObjRs.Fields("TIP_AMOUNT").Value&","&ObjRs.Fields("CARDMEMBER_NAME").Value&","&ObjRs.Fields("SPECIAL_BILL_IND").Value&","&ObjRs.Fields("ORIGINATING_BCA").Value&","&ObjRs.Fields("ORIGINATING_ACCOUNT_NUMBER").Value&","&ObjRs.Fields("CM_REFERENCE_NUMBER").Value&","&ObjRs.Fields("SUPPLIER_REFERENCE_NUMBER").Value&","&ObjRs.Fields("SHIP_TO_ZIP").Value&","&ObjRs.Fields("SIC_CODE").Value&","&ObjRs.Fields("COST_CENTER").Value&","&ObjRs.Fields("EMPLOYEE_ID").Value&","&ObjRs.Fields("SOCIAL_SECURITY_HASH_CODE").Value&","&ObjRs.Fields("UNIVERSALHASH_CODE").Value&","&ObjRs.Fields("STREET").Value&","&ObjRs.Fields("CITY").Value&","&ObjRs.Fields("STATE").Value&","&ObjRs.Fields("ZIP_PLUS__4").Value&","&ObjRs.Fields("TRANS_LIMIT").Value&","&ObjRs.Fields("MONTHLY_LIMIT").Value&","&ObjRs.Fields("EXPOSURE_LIMIT").Value&","&ObjRs.Fields("REV_CODE").Value&","&ObjRs.Fields("COMPANY_NAME").Value&","&ObjRs.Fields("CHARGE_DESCRIPTION_LINE1").Value&","&ObjRs.Fields("CHARGE_DESCRIPTION_LINE2").Value&","&ObjRs.Fields("CHARGE_DESCRIPTION_LINE3").Value&","&ObjRs.Fields("CHARGE_DESCRIPTION_LINE4").Value&","&ObjRs.Fields("CAR_RENTAL_CUSTOMER_NAME").Value&","&ObjRs.Fields("CAR_RENTAL_CITY").Value&","&ObjRs.Fields("CAR_RENTAL_STATE").Value&","&ObjRs.Fields("CAR_RENTAL_DATE").Value&","&ObjRs.Fields("CAR_RETURN_CITY").Value&","&ObjRs.Fields("CAR_RETURN_STATE").Value&","&ObjRs.Fields("CAR_RETURN_DATE").Value&","&ObjRs.Fields("CAR_RENTAL_DAYS").Value&","&ObjRs.Fields("HOTEL_ARRIVAL_DATE").Value&","&ObjRs.Fields("HOTEL_CITY").Value&","&ObjRs.Fields("HOTEL_STATE").Value&","&ObjRs.Fields("HOTEL_DEPART_DATE").Value&","&ObjRs.Fields("HOTEL_STAY_DURATION").Value&","&ObjRs.Fields("HOTEL_ROOM_RATE").Value&","&ObjRs.Fields("AIR_AGENCY_NUMBER").Value&","&ObjRs.Fields("AIR_TICKET_ISSUER").Value&","&ObjRs.Fields("AIR_CLASS_OF_SERVICE").Value&","&ObjRs.Fields("AIR_CARRIER_CODE").Value&","&ObjRs.Fields("AIR_ROUTING").Value&","&ObjRs.Fields("AIR_DEPARTURE_DATE").Value&","&ObjRs.Fields("AIR_PASSENGER_NAME").Value&","&ObjRs.Fields("TELE_DATE_OF_CALL").Value&","&ObjRs.Fields("TELE_FROM_CITY").Value&","&ObjRs.Fields("TELE_FROM_STATE").Value&","&ObjRs.Fields("TELE_CALL_LENGTH").Value&","&ObjRs.Fields("TELE_REFERENCE_NUMBER").Value&","&ObjRs.Fields("TELE_TIME_OF_CALL").Value&","&ObjRs.Fields("TELE_TO_NUMBER").Value&","&ObjRs.Fields("INDUSTRY_CODE").Value&","&ObjRs.Fields("SEQUENCE_NUMBER").Value&","&ObjRs.Fields("MERCATOR_KEY").Value&","&ObjRs.Fields("FEE_ALLOCATOR_IND").Value&","&"")
	 		ObjRs.MoveNext()
	 		'Response.Write(ObjRs("REQUESTING_CONTROL_ACCOUNT"))
	 	Wend

	 	OFiles.Close()
	 	Set OFiles = Nothing

	 	call DownloadFileAsCsv(strGroupBy)
	 	'Response.write"MasterCard File Transformed.<br><a href="&FileID&">Download File</a>"  
	 End If
	 %>
	 <%
	 	Private Sub DownloadFileAsCsv(strGroupBy)
	 		Dim Filename,OFiles,ExportObj

	 		Filename=strGroupBy& CSTR(Month(DATE)) & CSTR(Day(DATE))& CSTR(YEAR(date))
	 		FileID=server.MapPath("../Amex/file/"&Filename&".csv")
	 		FileAlias = Filename&".csv"

	 		Set ExportObj = Server.CreateObject("Scripting.FileSystemObject")
	 		Set OFiles = ExportObj.OpenTextFile(FileID,1,True)
	 		
	 		Response.CharSet = "ASCII"

	 		Response.AddHeader "Content-Disposition","attachment;filename="&FileAlias
	 		Response.ContentType = "text/csv;charset=US-ASCII"
	 		Response.Clear()

	 		Response.Write(OFiles.readAll)	 		
	 		Response.Flush

	 		OFiles.Close
	 		Set ExportObj = Nothing
	 		Set OFiles = Nothing
	 	End sub
	 %>
	 <%
	 	ObjRs.Close()
	 	'Conn.Close()

	 	Set ObjRs = Nothing
	 	'Set Conn = Nothing
	 %>
