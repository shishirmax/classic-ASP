<%@ Language = VBScript %>
<%server.ScriptTimeout = 1800%>
<%ON ERROR RESUME NEXT%>
<!--#include file="../Common/mainHeader.asp" -->
<!--#include file="../Shared/SharedPost.asp" -->
<%
Call CheckSession()
If Request.QueryString("tabIndex") <> "" Then
		Session("tabIndex") = Request.QueryString("tabIndex")
End If
'Response.Write(Session("tabIndex"))
firmcode=Session("FirmCode")
'Response.Write(firmcode)
Loginid = Session("LoginID")
'Response.Write(Loginid)
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
	<style type="text/css">
  a{
    text-decoration: none;
    font-weight: bold;
  }
  a:hover{
    color: #DB5375;
  }
  </style>
</head>
<body style="padding-top: 0px;" leftmargin="0" topmargin="0">
<% 	  	   
	  if UCase(trim(Session("firmCode"))) = "SLWK" then
			strReturnPath = "dashboard.asp"
				call OrgHeader()						 
	  End if		
	 %> 
	<%
	Dim objFso, objFile, sRows, arrRows
	Dim objConn, strSql, ObjExec
	Dim sFileName
	Dim ObjUpload
	Dim completeRow 
	Dim objPTOUpload, objPTOUploadFile,rsobj
    Dim ObjRpt,splitLine
	
	'New Lines For Upload(Start)-------------------------------------
	Set objUpload = Server.CreateObject("PTFMInvoicing.cUpload")
    Set ObjRpt = Server.CreateObject("PTFMInvoicing.DBHelper")
	strUploadPath = server.MapPath("file")
	strFileExtensions = ".csv;"	
	
	fullFilePath = strUploadPath & "/" & objUpload.Form("file").Value
	ImportedFileName = objUpload.Form("file").Value
	
	dtArr=Split(Date(),"/")
	yyyymmdd=dtArr(2) & dtArr(0) & dtArr(1)
	FullFileName = "MC"&yyyymmdd&hour(now)&minute(now)&second(now)&".txt"
	
	objUpload.Form("file").SaveFile strUploadPath, FullFileName, strFileExtensions
	
	
	Set objFso = CreateObject("Scripting.FileSystemObject")

	If Not objFso.FileExists(Server.MapPath("file/"& FullFileName)) Then
	Response.Write("File Not Found..")
	Else

	Set objFile = objFso.OpenTextFile(Server.MapPath("file/"&FullFileName),1,True)

	fName = objFso.GetBaseName(Server.MapPath("file/" & FullFileName))
	LoginID = Session("LoginID")

	'Connection to save file info in tabel
	fPath = objFso.GetAbsolutePathName(Server.MapPath("file/" & FullFileName))

	'-----Parameter for Master Card File-----
	Dim arrParams(74)'For Number of parameters
	'-----Actual parameters-----------------------------
        Dim arrParam1(3)
        Dim arrParam2(3)
        Dim arrParam3(3)
        Dim arrParam4(3)
        Dim arrParam5(3)
        Dim arrParam6(3)
        Dim arrParam7(3)
        Dim arrParam8(3)
        Dim arrParam9(3)
        Dim arrParam10(3)
    
        Dim arrParam11(3)
        Dim arrParam12(3)
        Dim arrParam13(3)
        Dim arrParam14(3)
        Dim arrParam15(3)
        Dim arrParam16(3)
        Dim arrParam17(3)
        Dim arrParam18(3)
        Dim arrParam19(3)
        Dim arrParam20(3)
    
        Dim arrParam21(3)
        Dim arrParam22(3)
        Dim arrParam23(3)
        Dim arrParam24(3)
        Dim arrParam25(3)
        Dim arrParam26(3)
        Dim arrParam27(3)
        Dim arrParam28(3)
        Dim arrParam29(3)
        Dim arrParam30(3)   
    
        Dim arrParam31(3)
        Dim arrParam32(3)
        Dim arrParam33(3)
        Dim arrParam34(3)
        Dim arrParam35(3)
        Dim arrParam36(3)
        Dim arrParam37(3)
        Dim arrParam38(3)
        Dim arrParam39(3)
        Dim arrParam40(3)
    
        Dim arrParam41(3)
        Dim arrParam42(3)
        Dim arrParam43(3)
        Dim arrParam44(3)
        Dim arrParam45(3)
        Dim arrParam46(3)
        Dim arrParam47(3)
        Dim arrParam48(3)
        Dim arrParam49(3)
        Dim arrParam50(3)
    
        Dim arrParam51(3)
        Dim arrParam52(3)
        Dim arrParam53(3)
        Dim arrParam54(3)
        Dim arrParam55(3)
        Dim arrParam56(3)
        Dim arrParam57(3)
        Dim arrParam58(3)
        Dim arrParam59(3)
        Dim arrParam60(3)
    
        Dim arrParam61(3)
        Dim arrParam62(3)
        Dim arrParam63(3)
        Dim arrParam64(3)
        Dim arrParam65(3)
        Dim arrParam66(3)
        Dim arrParam67(3)
        Dim arrParam68(3)
        Dim arrParam69(3)
        Dim arrParam70(3)
    
        Dim arrParam71(3)
        Dim arrParam72(3)
        Dim arrParam73(3)
        Dim arrParam74(3)
        Dim arrParam75(3)

	


    objFile.ReadLine
    objFile.ReadLine

	Do Until objFile.AtEndOfStream
		splitLine = split(objFile.ReadLine,",")


            if ubound(splitLine)<>75 Then
                Response.Write("File is not in the proper format. &nbsp;")
                Response.Write("<a href='../amex/ImportChase.asp'>Back</a>")
            End If 
       'if ubound(splitLine)>=74 and splitLine(0)<>"REQUESTING_CONTROL_ACCOUNT" and splitLine(0)<>"ACC.Account Number" and splitLine(0)<>" " and splitLine(0)<>"" Then
                        
            arrParam1(0) = "REQUESTING_CONTROL_ACCOUNT"
            arrParam1(1) = 129
            arrParam1(2) = 100
            arrParam1(3) = splitLine(0)
            arrParams(0) = arrParam1

            arrParam2(0) = "BASIC_CONTROL_ACCOUNT"
            arrParam2(1) = 129
            arrParam2(2) = 100
            arrParam2(3) = splitLine(1)
            arrParams(1) = arrParam2

            arrParam3(0) = "CARDMEMBER_ACCOUNT_NUMBER"
            arrParam3(1) = 129
            arrParam3(2) = 100
            arrParam3(3) = splitLine(2)
            arrParams(2) = arrParam3

            arrParam4(0) = "SE_NUMBER"
            arrParam4(1) = 129
            arrParam4(2) = 100
            arrParam4(3) = splitLine(3)
            arrParams(3) = arrParam4
            
            arrParam5(0) = "ROC_ID"
            arrParam5(1) = 129
            arrParam5(2) = 100
            arrParam5(3) = splitLine(4)
            arrParams(4) = arrParam5
            
            arrParam6(0) = "DB_CR_INDICATOR"
            arrParam6(1) = 129
            arrParam6(2) = 100
            arrParam6(3) = splitLine(5)
            arrParams(5) = arrParam6
            
            arrParam7(0) = "TRANSACTION_TYPE_CODE"
            arrParam7(1) = 129
            arrParam7(2) = 100
            arrParam7(3) = splitLine(6)
            arrParams(6) = arrParam7
            
            arrParam8(0) = "FINANCIAL_CATEGORY"
            arrParam8(1) = 129
            arrParam8(2) = 100
            arrParam8(3) = splitLine(7)
            arrParams(7) = arrParam8
            
            arrParam9(0) = "BATCH_NUMBER"
            arrParam9(1) = 129
            arrParam9(2) = 100
            arrParam9(3) = splitLine(8)
            arrParams(8) = arrParam9
            
            arrParam10(0) = "DATE_OF_CHARGE"
            arrParam10(1) = 129
            arrParam10(2) = 100
            arrParam10(3) = splitLine(9)
            arrParams(9) = arrParam10
            
            arrParam11(0) = "LOCAL_CURRENCY_AMOUNT"
            arrParam11(1) = 129
            arrParam11(2) = 100
            arrParam11(3) = splitLine(10)
            arrParams(10) = arrParam11
            
            arrParam12(0) = "CURRENCY_CODE"
            arrParam12(1) = 129
            arrParam12(2) = 100
            arrParam12(3) = splitLine(11)
            arrParams(11) = arrParam12
            
            arrParam13(0) = "CAPTURE_DATE"
            arrParam13(1) = 129
            arrParam13(2) = 100
            arrParam13(3) = splitLine(12)
            arrParams(12) = arrParam13
            
            arrParam14(0) = "PROCESS_DATE"
            arrParam14(1) = 129
            arrParam14(2) = 100
            arrParam14(3) = splitLine(13)
            arrParams(13) = arrParam14
            
            arrParam15(0) = "BILLING_DATE"
            arrParam15(1) = 129
            arrParam15(2) = 100
            arrParam15(3) = splitLine(14)
            arrParams(14) = arrParam15
            
            arrParam16(0) = "BILLING_AMOUNT"
            arrParam16(1) = 129
            arrParam16(2) = 100
            arrParam16(3) = splitLine(15)
            arrParams(15) = arrParam16
            
            arrParam17(0) = "SALES_TAX_AMOUNT"
            arrParam17(1) = 129
            arrParam17(2) = 100
            arrParam17(3) = splitLine(16)
            arrParams(16) = arrParam17
            
            arrParam18(0) = "TIP_AMOUNT"
            arrParam18(1) = 129
            arrParam18(2) = 100
            arrParam18(3) = splitLine(17)
            arrParams(17) = arrParam18
            
            arrParam19(0) = "CARDMEMBER_NAME"
            arrParam19(1) = 129
            arrParam19(2) = 100
            arrParam19(3) = splitLine(18)
            arrParams(18) = arrParam19
            
            arrParam20(0) = "SPECIAL_BILL_IND"
            arrParam20(1) = 129
            arrParam20(2) = 100
            arrParam20(3) = splitLine(19)
            arrParams(19) = arrParam20
            
            arrParam21(0) = "ORIGINATING_BCA"
            arrParam21(1) = 129
            arrParam21(2) = 100
            arrParam21(3) = splitLine(20)
            arrParams(20) = arrParam21
            
            arrParam22(0) = "ORIGINATING_ACCOUNT_NUMBER"
            arrParam22(1) = 129
            arrParam22(2) = 100
            arrParam22(3) = splitLine(21)
            arrParams(21) = arrParam22
            
            arrParam23(0) = "CM_REFERENCE_NUMBER"
            arrParam23(1) = 129
            arrParam23(2) = 100
            arrParam23(3) = splitLine(22)
            arrParams(22) = arrParam23
            
            arrParam24(0) = "SUPPLIER_REFERENCE_NUMBER"
            arrParam24(1) = 129
            arrParam24(2) = 100
            arrParam24(3) = splitLine(23)
            arrParams(23) = arrParam24
            
            arrParam25(0) = "SHIP_TO_ZIP"
            arrParam25(1) = 129
            arrParam25(2) = 100
            arrParam25(3) = splitLine(24)
            arrParams(24) = arrParam25
            
            arrParam26(0) = "SIC_CODE"
            arrParam26(1) = 129
            arrParam26(2) = 100
            arrParam26(3) = splitLine(25)
            arrParams(25) = arrParam25
            
            arrParam27(0) = "COST_CENTER"
            arrParam27(1) = 129
            arrParam27(2) = 100
            arrParam27(3) = splitLine(26)
            arrParams(26) = arrParam27
            
            arrParam28(0) = "EMPLOYEE_ID"
            arrParam28(1) = 129
            arrParam28(2) = 100
            arrParam28(3) = splitLine(27)
            arrParams(27) = arrParam28
            
            arrParam29(0) = "SOCIAL_SECURITY_HASH_CODE"
            arrParam29(1) = 129
            arrParam29(2) = 100
            arrParam29(3) = splitLine(28)
            arrParams(28) = arrParam29
            
            arrParam30(0) = "UNIVERSALHASH_CODE"
            arrParam30(1) = 129
            arrParam30(2) = 100
            arrParam30(3) = splitLine(29)
            arrParams(29) = arrParam30
            
            arrParam31(0) = "STREET"
            arrParam31(1) = 129
            arrParam31(2) = 100
            arrParam31(3) = splitLine(30)
            arrParams(30) = arrParam31
            
            arrParam32(0) = "CITY"
            arrParam32(1) = 129
            arrParam32(2) = 100
            arrParam32(3) = splitLine(31)
            arrParams(31) = arrParam32
            
            arrParam33(0) = "STATE"
            arrParam33(1) = 129
            arrParam33(2) = 100
            arrParam33(3) = splitLine(32)
            arrParams(32) = arrParam33
            
            arrParam34(0) = "ZIP_PLUS__4"
            arrParam34(1) = 129
            arrParam34(2) = 100
            arrParam34(3) = splitLine(33)
            arrParams(33) = arrParam34

            arrParam35(0) = "TRANS_LIMIT"
            arrParam35(1) = 129
            arrParam35(2) = 100
            arrParam35(3) = splitLine(34)
            arrParams(34) = arrParam35

            arrParam36(0) = "MONTHLY_LIMIT"
            arrParam36(1) = 129
            arrParam36(2) = 100
            arrParam36(3) = splitLine(35)
            arrParams(35) = arrParam36

            arrParam37(0) = "EXPOSURE_LIMIT"
            arrParam37(1) = 129
            arrParam37(2) = 100
            arrParam37(3) = splitLine(36)
            arrParams(36) = arrParam37

            arrParam38(0) = "REV_CODE"
            arrParam38(1) = 129
            arrParam38(2) = 100
            arrParam38(3) = splitLine(37)
            arrParams(37) = arrParam38

            arrParam39(0) = "COMPANY_NAME"
            arrParam39(1) = 129
            arrParam39(2) = 100
            arrParam39(3) = splitLine(38)
            arrParams(38) = arrParam39

            arrParam40(0) = "CHARGE_DESCRIPTION_LINE1"
            arrParam40(1) = 129
            arrParam40(2) = 100
            arrParam40(3) = splitLine(39)
            arrParams(39) = arrParam40

            arrParam41(0) = "CHARGE_DESCRIPTION_LINE2"
            arrParam41(1) = 129
            arrParam41(2) = 100
            arrParam41(3) = splitLine(40)
            arrParams(40) = arrParam41

            arrParam42(0) = "CHARGE_DESCRIPTION_LINE3"
            arrParam42(1) = 129
            arrParam42(2) = 100
            arrParam42(3) = splitLine(41)
            arrParams(41) = arrParam42

            arrParam43(0) = "CHARGE_DESCRIPTION_LINE4"
            arrParam43(1) = 129
            arrParam43(2) = 100
            arrParam43(3) = splitLine(42)
            arrParams(42) = arrParam43

            arrParam44(0) = "CAR_RENTAL_CUSTOMER_NAME"
            arrParam44(1) = 129
            arrParam44(2) = 100
            arrParam44(3) = splitLine(43)
            arrParams(43) = arrParam44

            arrParam45(0) = "CAR_RENTAL_CITY"
            arrParam45(1) = 129
            arrParam45(2) = 100
            arrParam45(3) = splitLine(44)
            arrParams(44) = arrParam45

            arrParam46(0) = "CAR_RENTAL_STATE"
            arrParam46(1) = 129
            arrParam46(2) = 100
            arrParam46(3) = splitLine(45)
            arrParams(45) = arrParam46

            arrParam47(0) = "CAR_RENTAL_DATE"
            arrParam47(1) = 129
            arrParam47(2) = 100
            arrParam47(3) = splitLine(46)
            arrParams(46) = arrParam47

            arrParam48(0) = "CAR_RETURN_CITY"
            arrParam48(1) = 129
            arrParam48(2) = 100
            arrParam48(3) = splitLine(47)
            arrParams(47) = arrParam47

            arrParam49(0) = "CAR_RETURN_STATE"
            arrParam49(1) = 129
            arrParam49(2) = 100
            arrParam49(3) = splitLine(48)
            arrParams(48) = arrParam49            

            arrParam50(0) = "CAR_RETURN_DATE"
            arrParam50(1) = 129
            arrParam50(2) = 100
            arrParam50(3) = splitLine(49)
            arrParams(49) = arrParam50

            arrParam51(0) = "CAR_RENTAL_DAYS"
            arrParam51(1) = 129
            arrParam51(2) = 100
            arrParam51(3) = splitLine(50)
            arrParams(50) = arrParam51

            arrParam52(0) = "HOTEL_ARRIVAL_DATE"
            arrParam52(1) = 129
            arrParam52(2) = 100
            arrParam52(3) = splitLine(51)
            arrParams(51) = arrParam52

            arrParam53(0) = "HOTEL_CITY"
            arrParam53(1) = 129
            arrParam53(2) = 100
            arrParam53(3) = splitLine(52)
            arrParams(52) = arrParam53

            arrParam54(0) = "HOTEL_STATE"
            arrParam54(1) = 129
            arrParam54(2) = 100
            arrParam54(3) = splitLine(53)
            arrParams(53) = arrParam54

            arrParam55(0) = "HOTEL_DEPART_DATE"
            arrParam55(1) = 129
            arrParam55(2) = 100
            arrParam55(3) = splitLine(54)
            arrParams(54) = arrParam55

            arrParam56(0) = "HOTEL_STAY_DURATION"
            arrParam56(1) = 129
            arrParam56(2) = 100
            arrParam56(3) = splitLine(55)
            arrParams(55) = arrParam56

            arrParam57(0) = "HOTEL_ROOM_RATE"
            arrParam57(1) = 129
            arrParam57(2) = 100
            arrParam57(3) = splitLine(56)
            arrParams(56) = arrParam57

            arrParam58(0) = "AIR_AGENCY_NUMBER"
            arrParam58(1) = 129
            arrParam58(2) = 100
            arrParam58(3) = splitLine(57)
            arrParams(57) = arrParam58

            arrParam59(0) = "AIR_TICKET_ISSUER"
            arrParam59(1) = 129
            arrParam59(2) = 100
            arrParam59(3) = splitLine(58)
            arrParams(58) = arrParam59

            arrParam60(0) = "AIR_CLASS_OF_SERVICE"
            arrParam60(1) = 129
            arrParam60(2) = 100
            arrParam60(3) = splitLine(59)
            arrParams(59) = arrParam60

            arrParam61(0) = "AIR_CARRIER_CODE"
            arrParam61(1) = 129
            arrParam61(2) = 100
            arrParam61(3) = splitLine(60)
            arrParams(60) = arrParam61

            arrParam62(0) = "AIR_ROUTING"
            arrParam62(1) = 129
            arrParam62(2) = 100
            arrParam62(3) = splitLine(61)
            arrParams(61) = arrParam62

            arrParam63(0) = "AIR_DEPARTURE_DATE"
            arrParam63(1) = 129
            arrParam63(2) = 100
            arrParam63(3) = splitLine(62)
            arrParams(62) = arrParam63

            arrParam64(0) = "AIR_PASSENGER_NAME"
            arrParam64(1) = 129
            arrParam64(2) = 100
            arrParam64(3) = splitLine(63)
            arrParams(63) = arrParam64

            arrParam65(0) = "TELE_DATE_OF_CALL"
            arrParam65(1) = 129
            arrParam65(2) = 100
            arrParam65(3) = splitLine(64)
            arrParams(64) = arrParam65

            arrParam66(0) = "TELE_FROM_CITY"
            arrParam66(1) = 129
            arrParam66(2) = 100
            arrParam66(3) = splitLine(65)
            arrParams(65) = arrParam66

            arrParam67(0) = "TELE_FROM_STATE"
            arrParam67(1) = 129
            arrParam67(2) = 100
            arrParam67(3) = splitLine(66)
            arrParams(66) = arrParam67

            arrParam68(0) = "TELE_CALL_LENGTH"
            arrParam68(1) = 129
            arrParam68(2) = 100
            arrParam68(3) = splitLine(67)
            arrParams(67) = arrParam68

            arrParam69(0) = "TELE_REFERENCE_NUMBER"
            arrParam69(1) = 129
            arrParam69(2) = 100
            arrParam69(3) = splitLine(68)
            arrParams(68) = arrParam69

            arrParam70(0) = "TELE_TIME_OF_CALL"
            arrParam70(1) = 129
            arrParam70(2) = 100
            arrParam70(3) = splitLine(69)
            arrParams(69) = arrParam70

            arrParam71(0) = "TELE_TO_NUMBER"
            arrParam71(1) = 129
            arrParam71(2) = 100
            arrParam71(3) = splitLine(70)
            arrParams(70) = arrParam71

            arrParam72(0) = "INDUSTRY_CODE"
            arrParam72(1) = 129
            arrParam72(2) = 100
            arrParam72(3) = splitLine(71)
            arrParams(71) = arrParam72

            arrParam73(0) = "SEQUENCE_NUMBER"
            arrParam73(1) = 129
            arrParam73(2) = 100
            arrParam73(3) = splitLine(72)
            arrParams(72) = arrParam73

            arrParam74(0) = "MERCATOR_KEY"            
            arrParam74(1) = 129
            arrParam74(2) = 100
            arrParam74(3) = splitLine(73)
            arrParams(73) = arrParam74

            arrParam75(0) = "FEE_ALLOCATOR_IND"
            arrParam75(1) = 129
            arrParam75(2) = 100
            arrParam75(3) = splitLine(74)
            arrParams(74) = arrParam75
            
            set rsMCData = ObjRpt.RunSPReturnRS("SP_ImportMasterCard",arrParams,"")

            'Response.Write(line2(0)&","&line2(1)&","&line2(2)&","&line2(3)&","&line2(4)&","&line2(5)&","&line2(6)&"<br/>")
        'End if
    Loop
    objFile.Close
	Set objFile = Nothing
	Set objConn = Nothing

	End if
	Response.Redirect("../Amex/ImportChase.asp?tabIndex=3&flg=1")

	%>
</body>
</html>