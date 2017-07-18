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
	
	'New Lines For Upload(Start)-------------------------------------
	Set objUpload = Server.CreateObject("PTFMInvoicing.cUpload")
	strUploadPath = server.MapPath("file")
	strFileExtensions = ".csv;"	
	
	fullFilePath = strUploadPath & "/" & objUpload.Form("file").Value
	ImportedFileName = objUpload.Form("file").Value
	
	dtArr=Split(Date(),"/")
	yyyymmdd=dtArr(2) & dtArr(0) & dtArr(1)
	FullFileName = "CCPTO"&yyyymmdd&hour(now)&minute(now)&second(now)&".csv"
	
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

	'-----Parameter for Monthly Transaction Summary-----
	Dim arrParamsMTS(8)'For Number of parameters
	'-----Actual parameters-----------------------------
	Dim arrParamsMTS1(3)'mts_DatePosted
	Dim arrParamsMTS2(3)'mts_Transactionreference
	Dim arrParamsMTS3(3)'mts_AttorneyDocket
	Dim arrParamsMTS4(3)'mts_Status
	Dim arrParamsMTS5(3)'mts_TransactionID
	Dim arrParamsMTS6(3)'mts_Type
	Dim arrParamsMTS7(3)'mts_TotalPayment_Refund
	Dim arrParamsMTS8(3)'mts_CustomerName
	Dim arrParamsMTS9(3)'mts_FileName


	'-----Parameter for Monthly Transaction Detail-----
	Dim arrParamsMTD(13)'For Number of parameters
	'-----Actual parameters-----------------------------
	Dim arrParamsMTD1(3)'mtd_PaymentDatePosted
	Dim arrParamsMTD2(3)'mtd_SaleItemDatePosted
	Dim arrParamsMTD3(3)'mtd_SaleItemReference
	Dim arrParamsMTD4(3)'mtd_AttorneyDocket
	Dim arrParamsMTD5(3)'mtd_Status
	Dim arrParamsMTD6(3)'mtd_TransactionID
	Dim arrParamsMTD7(3)'mtd_SaleID
	Dim arrParamsMTD8(3)'mtd_Feecode
	Dim arrParamsMTD9(3)'mtd_FeeCodeDescription
	Dim arrParamsMTD10(3)'mtd_ItemPrice
	Dim arrParamsMTD11(3)'mtd_Quantity
	Dim arrParamsMTD12(3)'mtd_ItemTotal
	Dim arrParamsMTD13(3)'mtd_CustomerName
	Dim arrParamsMTD14(3)'mtd_FileName
	

	Dim ObjRpt,splitLine
	Set ObjRpt = Server.CreateObject("PTFMInvoicing.DBHelper")


	Do Until objFile.AtEndOfStream
		splitLine = split(objFile.ReadLine,",")
    if ubound(splitLine)>=0 Then
        if splitLine(0)="Monthly Transactions Summary" Then
            mtsFlag = true

        elseif splitLine(0)="Monthly Transactions Details (Payments Only)" Then
            mtdFlag = true
            mtsFlag = false
        end if
        if mtsFlag = true Then
       if ubound(splitLine)>=7 and splitLine(0)<>"Date Posted" and splitLine(0)<>"Monthly Transactions Summary" and splitLine(0)<>" " and splitLine(0)<>"" Then
        'Response.Write(mtsFlag)
            counter=0
            line = split(objFile.ReadLine,"""")

            replacedLine=""
            for each val in line
             
                'Response.Write (val)

            if counter mod 2 =1 Then
        	   	val=Replace(val,",","")
            end if
          
            replacedLine=replacedLine & val
           
            counter=counter + 1
            Next
            'Response.Write "<tr>"

            'Response.Write "<td><b>" &replacedLine&"</b></td>"& chr(13)

            line2 = split(replacedLine,",")

            arrParamsMTS1(0) = "mts_DatePosted"
            arrParamsMTS1(1) = 129
            arrParamsMTS1(2) = 50
            arrParamsMTS1(3) = line2(0)
            arrParamsMTS(0) = arrParamsMTS1

            arrParamsMTS2(0) = "mts_Transactionreference"
            arrParamsMTS2(1) = 129
            arrParamsMTS2(2) = 50
            arrParamsMTS2(3) = line2(1)
            arrParamsMTS(1) = arrParamsMTS2

            arrParamsMTS3(0) = "mts_AttorneyDocket"
            arrParamsMTS3(1) = 129
            arrParamsMTS3(2) = 50
            arrParamsMTS3(3) = line2(2)
            arrParamsMTS(2) = arrParamsMTS3

            arrParamsMTS4(0) = "mts_Status"
            arrParamsMTS4(1) = 129
            arrParamsMTS4(2) = 100
            arrParamsMTS4(3) = line2(3)
            arrParamsMTS(3) = arrParamsMTS4

            arrParamsMTS5(0) = "mts_TransactionID"
            arrParamsMTS5(1) = 129
            arrParamsMTS5(2) = 50
            arrParamsMTS5(3) = line2(4)
            arrParamsMTS(4) = arrParamsMTS5

            arrParamsMTS6(0) = "mts_Type"
            arrParamsMTS6(1) = 129
            arrParamsMTS6(2) = 100
            arrParamsMTS6(3) = line2(5)
            arrParamsMTS(5) = arrParamsMTS6

            arrParamsMTS7(0) = "mts_TotalPayment_Refund"
            arrParamsMTS7(1) = 129
            arrParamsMTS7(2) = 50
            arrParamsMTS7(3) = line2(6)
            arrParamsMTS(6) = arrParamsMTS7

            arrParamsMTS8(0) = "mts_CustomerName"
            arrParamsMTS8(1) = 129
            arrParamsMTS8(2) = 50
            arrParamsMTS8(3) = line2(7)
            arrParamsMTS(7) = arrParamsMTS8

            arrParamsMTS9(0) = "fileName"
            arrParamsMTS9(1) = 129
            arrParamsMTS9(2) = 50
            arrParamsMTS9(3) = ImportedFileName
            arrParamsMTS(8) = arrParamsMTS9

            set rsMTS = ObjRpt.RunSPReturnRS("sp_ImportMTS",arrParamsMTS,"")

            'Response.Write(line2(0)&","&line2(1)&","&line2(2)&","&line2(3)&","&line2(4)&","&line2(5)&","&line2(6)&"<br/>")
        End if

    elseif mtdFlag = true Then
        if ubound(splitLine)>=12 and splitLine(0)<>"Payment Date Posted" and splitLine(0)<>"Monthly Transactions Details (Payments Only)" and splitLine(0)<>" " and splitLine(0)<>"" Then
    'Response.Write(mtdFlag)
        counter=0
            line = split(objFile.ReadLine,"""")

            replacedLine=""
            for each val in line
             
                'Response.Write (val)

            if counter mod 2 =1 Then
                val=Replace(val,",","")
            end if
          
            replacedLine=replacedLine & val
           
            counter=counter + 1
            Next
            'Response.Write "<tr>"

            'Response.Write "<td><b>" &replacedLine&"</b></td>"& chr(13)

            lineMTD = split(replacedLine,",")


            arrParamsMTD1(0) = "mtd_PaymentDatePosted"
            arrParamsMTD1(1) = 129
            arrParamsMTD1(2) = 50
            arrParamsMTD1(3) = lineMTD(0)
            arrParamsMTD(0) = arrParamsMTD1

            arrParamsMTD2(0) = "mtd_SaleItemDatePosted"
            arrParamsMTD2(1) = 129
            arrParamsMTD2(2) = 50
            arrParamsMTD2(3) = lineMTD(1)
            arrParamsMTD(1) = arrParamsMTD2

            arrParamsMTD3(0) = "mtd_SaleItemReference"
            arrParamsMTD3(1) = 129
            arrParamsMTD3(2) = 50
            arrParamsMTD3(3) = lineMTD(2)
            arrParamsMTD(2) = arrParamsMTD3

            arrParamsMTD4(0) = "mtd_AttorneyDocket"
            arrParamsMTD4(1) = 129
            arrParamsMTD4(2) = 50
            arrParamsMTD4(3) = lineMTD(3)
            arrParamsMTD(3) = arrParamsMTD4

            arrParamsMTD5(0) = "mtd_Status"
            arrParamsMTD5(1) = 129
            arrParamsMTD5(2) = 50
            arrParamsMTD5(3) = lineMTD(4)
            arrParamsMTD(4) = arrParamsMTD5

            arrParamsMTD6(0) = "mtd_TransactionID"
            arrParamsMTD6(1) = 129
            arrParamsMTD6(2) = 50
            arrParamsMTD6(3) = lineMTD(5)
            arrParamsMTD(5) = arrParamsMTD6

            arrParamsMTD7(0) = "mtd_SaleID"
            arrParamsMTD7(1) = 129
            arrParamsMTD7(2) = 50
            arrParamsMTD7(3) = lineMTD(6)
            arrParamsMTD(6) = arrParamsMTD7

            arrParamsMTD8(0) = "mtd_Feecode"
            arrParamsMTD8(1) = 129
            arrParamsMTD8(2) = 50
            arrParamsMTD8(3) = lineMTD(7)
            arrParamsMTD(7) = arrParamsMTD8

            arrParamsMTD9(0) = "mtd_FeeCodeDescription"
            arrParamsMTD9(1) = 129
            arrParamsMTD9(2) = 50
            arrParamsMTD9(3) = lineMTD(8)
            arrParamsMTD(8) = arrParamsMTD9

            arrParamsMTD10(0) = "mtd_ItemPrice"
            arrParamsMTD10(1) = 129
            arrParamsMTD10(2) = 50
            arrParamsMTD10(3) = lineMTD(9)
            arrParamsMTD(9) = arrParamsMTD10

            arrParamsMTD11(0) = "mtd_Quantity"
            arrParamsMTD11(1) = 129
            arrParamsMTD11(2) = 50
            arrParamsMTD11(3) = lineMTD(10)
            arrParamsMTD(10) = arrParamsMTD11

            arrParamsMTD12(0) = "mtd_ItemTotal"
            arrParamsMTD12(1) = 129
            arrParamsMTD12(2) = 50
            arrParamsMTD12(3) = lineMTD(11)
            arrParamsMTD(11) = arrParamsMTD12

            arrParamsMTD13(0) = "mtd_CustomerName"
            arrParamsMTD13(1) = 129
            arrParamsMTD13(2) = 50
            arrParamsMTD13(3) = lineMTD(12)
            arrParamsMTD(12) = arrParamsMTD13

            arrParamsMTD14(0) = "fileName"
            arrParamsMTD14(1) = 129
            arrParamsMTD14(2) = 50
            arrParamsMTD14(3) = ImportedFileName
            arrParamsMTD(13) = arrParamsMTD14

            set rsMTD = ObjRpt.RunSPReturnRS("sp_ImportMTD",arrParamsMTD,"")
        End If    
    End if
End if
Loop
objFile.Close

	
	Set objFile = Nothing
	Set objConn = Nothing

	End if
	Response.Redirect("../Amex/ImportPTOTransform.asp?tabIndex=3&flg=1")

	%>
</body>
</html>