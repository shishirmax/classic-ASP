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
%>

	 <% 	  	   
	  if UCase(trim(Session("firmCode"))) = "WGSB" then
			strReturnPath = "dashboard.asp"
				call OrgHeader()						 
	  End if		
	 '******************************************* Connection*************************
	 Dim ObjRpt
	 Set ObjRpt = Server.CreateObject("PTFMInvoicing.DBHelper")
	 'Set ObjRs = Server.CreateObject("ADODB.Recordset")
	 	'Dim arrParam(0)
	 	'Dim arrParam1(3)
	 	'Dim ObjRs
	 	'arrParam1(0) = "FirmCode"
	 	'arrParam1(1) = 129
	 	'arrParam1(2) = 50
	 	'arrParam1(3) = firmcode
	 	'arrParam(0) = arrParam1
	 	'Set ObjRs = ObjRpt.RunSPReturnRS("sp_GetTransformedPTOData",arrParam,"")



	 	'Get Sale Information
	 	Dim arrParamSI(0)
	 	Dim arrParamSI1(3)
	 	Dim ObjRsSI
	 	arrParamSI1(0) = "FirmCode"
	 	arrParamSI1(1) = 129
	 	arrParamSI1(2) = 50
	 	arrParamSI1(3) = firmcode
	 	arrParamSI(0) = arrParamSI1
	 	Set ObjRsSI = ObjRpt.RunSPReturnRS("sp_GetSaleInformation",arrParamSI,"")

	 	'Get Sale Detail
	 	Dim arrParamSD(0)
	 	Dim arrParamSD1(3)
	 	Dim ObjRsSD
	 	arrParamSD1(0) = "FirmCode"
	 	arrParamSD1(1) = 129
	 	arrParamSD1(2) = 50
	 	arrParamSD1(3) = firmcode
	 	arrParamSD(0) = arrParamSD1
	 	Set ObjRsSD = ObjRpt.RunSPReturnRS("sp_GetSaleDetail",arrParamSD,"")

	 	'Get Payment Detail
	 	Dim arrParamPD(0)
	 	Dim arrParamPD1(3)
	 	Dim ObjRsPD
	 	arrParamPD1(0) = "FirmCode"
	 	arrParamPD1(1) = 129
	 	arrParamPD1(2) = 50
	 	arrParamPD1(3) = firmcode
	 	arrParamPD(0) = arrParamPD1
	 	Set ObjRsPD = ObjRpt.RunSPReturnRS("sp_GetPaymentDetail",arrParamPD,"")

	 	'Get Refund Detail
	 	Dim arrParamRD(0)
	 	Dim arrParamRD1(3)
	 	Dim ObjRsRD
	 	arrParamRD1(0) = "FirmCode"
	 	arrParamRD1(1) = 129
	 	arrParamRD1(2) = 50
	 	arrParamRD1(3) = firmcode
	 	arrParamRD(0) = arrParamRD1
	 	Set ObjRsRD = ObjRpt.RunSPReturnRS("sp_GetRefundDetail",arrParamRD,"")


	'If Not ObjRsSI.EOF Then
	 		strGroupBy = "PTO-"
	 		FileName = strGroupBy & cSTR(Month(DATE)) & CSTR(Day(DATE))& CSTR(YEAR(date))& CSTR(Hour(Time))& CSTR(minute(Time)) & CSTR(second(Time))
	 		FileID = Server.MapPath("../Amex/file/" & FileName & ".csv")

	 		Set ObjFso = CreateObject("Scripting.FileSystemObject")
	 		Set OFiles = ObjFso.OpenTextFile(FileID, 2, true, 0)

'-------------------Writing Headers
			OFiles.WriteLine("United States Patent and Trademark Office")
			OFiles.WriteLine("Date Printed: "& Date)
			OFiles.WriteLine()
			OFiles.WriteLine()
			'Writting Sale Information Header
			OFiles.WriteLine("Sale Information")
			OFiles.WriteLine("CSV Reference Number"&","&"Type"&","&"Transaction Status"&","&"Accounting Date"&","&"Name/Number"&","&"Attorney Docket Number")
			'Writing Sale Information Data
			While Not ObjRsSI.EOF
				OFiles.WriteLine(ObjRsSI("SI_csv")&","&ObjRsSI("SI_Type")&","&ObjRsSI("SI_TransactionStatus")&","&ObjRsSI("SI_AccountingDate")&","&ObjRsSI("SI_NameNumber")&","&ObjRsSI("SI_AttorneyDocket"))
				ObjRsSI.MoveNext()
			Wend

			OFiles.WriteLine()
			OFiles.WriteLine()
			'Writing Sale Details Header
			OFiles.WriteLine("Sale Details")
			OFiles.WriteLine("CSV Reference Number"&","&"Name/Number"&","&"Attorney Docket Number"&","&"Transaction Status"&","&"Quantity"&","&"Item Total"&","&"Payment Amount"&","&"Fee Code"&","&"Description")
			'Writing Sale Details Data
			While Not ObjRsSD.EOF
			OFiles.WriteLine(ObjRsSD("SD_csv")&","&ObjRsSD("SD_Name_Number")&","&ObjRsSD("SD_AttorneyDocketNumber")&","&ObjRsSD("SD_TransactionStatus")&","&ObjRsSD("SD_Quantity")&","&ObjRsSD("SD_ItemPrice")&","&ObjRsSD("SD_ItemTotal")&","&ObjRsSD("SD_FeeCode")&","&ObjRsSD("SD_FeeCodeDescription"))
			ObjRsSD.MoveNext()
			Wend

			OFiles.WriteLine()
			OFiles.WriteLine()
			OFiles.WriteLine("Payment Details")
			OFiles.WriteLine("CSV Reference Number"&","&"Payment Type"&","&"Total Payment Amount"&","&"Payment Date"&","&"Payment Amount (this sale)")
			While Not ObjRsPD.EOF
			OFiles.WriteLine(ObjRsPD("PD_csv")&","&ObjRsPD("PD_PaymentType")&","&ObjRsPD("PD_TotalPaymentAmount")&","&ObjRsPD("PD_PaymentDate")&","&ObjRsPD("PD_PaymentAmount"))
			ObjRsPD.MoveNext()
			Wend

			OFiles.WriteLine()
			OFiles.WriteLine()
			OFiles.WriteLine("Refund Details")
			OFiles.WriteLine("CSV Reference Number"&","&"Refund ID"&","&"Accounting Date"&","&"Refund Amount"&","&"Name/Number"&","&"Payment Method"&","&"Payee Name")
			While Not ObjRsRD.EOF
			OFiles.WriteLine(ObjRsRD("RD_csv")&","&ObjRsRD("RD_TransactionID")&","&ObjRsRD("RD_AccountingDate")&","&ObjRsRD("RD_RefundAmount")&","&ObjRsRD("RD_TransactionRef")&","&ObjRsRD("RD_PaymentMethod")&","&ObjRsRD("RD_PayeeName"))
			ObjRsRD.MoveNext()
			wend
		OFiles.Close()
		Set OFiles = Nothing

		call DownloadFileAsCsv(strGroupBy)
	'End If
	 	%>
	 	<%
	 	Private Sub DownloadFileAsCsv(strGroupBy)
	 		Dim Filename,OFiles,ExportObj

	 		Filename=strGroupBy& cSTR(Month(DATE)) & CSTR(Day(DATE))& CSTR(YEAR(date))& CSTR(Hour(Time))& CSTR(minute(Time)) & CSTR(second(Time))
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
	 	ObjRsSI.Close()
	 	ObjRsSD.Close()
	 	ObjRsPD.Close()
	 	ObjRsRD.Close()
	 	Set ObjRsSI = Nothing
	 	Set ObjRsSD = Nothing
	 	Set ObjRsPD = Nothing
	 	Set ObjRsRD = Nothing
	 %>