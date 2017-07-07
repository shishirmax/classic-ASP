<%@language="vbscript"%>
<table border="1">
<%
'dim csv_to_read, fso, act, imported_text,split_text, total_imported_text,total_split_text,total_num_imported
dim csv_to_read,counter,line,fso,objFile,inRow()
Dim splitLine,mtsFlag,mtdFlag
csv_to_read="C:\file\CC2017.csv"

set fso = createobject("scripting.filesystemobject")
set objFile = fso.opentextfile(csv_to_read)
'objFile.ReadLine
'objFile.ReadLine
'objFile.ReadLine
'objFile.ReadLine

'-----Parameter for Monthly Transaction Summary-----
ReDim arrParamsMTS(8)'For Number of parameters
'-----Actual parameters-----------------------------
ReDim arrParamsMTS1(3)'mts_DatePosted
ReDim arrParamsMTS2(3)'mts_Transactionreference
ReDim arrParamsMTS3(3)'mts_AttorneyDocket
ReDim arrParamsMTS4(3)'mts_Status
ReDim arrParamsMTS5(3)'mts_TransactionID
ReDim arrParamsMTS6(3)'mts_Type
ReDim arrParamsMTS7(3)'mts_TotalPayment_Refund
ReDim arrParamsMTS8(3)'mts_CustomerName
ReDim arrParamsMTS9(3)'mts_FileName


'-----Parameter for Monthly Transaction Detail-----
ReDim arrParamsMTD(13)'For Number of parameters
'-----Actual parameters-----------------------------
ReDim arrParamsMTD1(3)'mtd_PaymentDatePosted
ReDim arrParamsMTD2(3)'mtd_SaleItemDatePosted
ReDim arrParamsMTD3(3)'mtd_SaleItemReference
ReDim arrParamsMTD4(3)'mtd_AttorneyDocket
ReDim arrParamsMTD5(3)'mtd_Status
ReDim arrParamsMTD6(3)'mtd_TransactionID
ReDim arrParamsMTD7(3)'mtd_SaleID
ReDim arrParamsMTD8(3)'mtd_Feecode
ReDim arrParamsMTD9(3)'mtd_FeeCodeDescription
ReDim arrParamsMTD10(3)'mtd_ItemPrice
ReDim arrParamsMTD11(3)'mtd_Quantity
ReDim arrParamsMTD12(3)'mtd_ItemTotal
ReDim arrParamsMTD13(3)'mtd_CustomerName
ReDim arrParamsMTD14(3)'mtd_FileName

Dim ObjRpt
Set ObjRpt = Server.CreateObject("PTFMInvoicing.DBHelper")

Do Until objFile.AtEndOfStream
    splitLine = split(objFile.ReadLine,",")
    'Response.Write(splitLine)
    if ubound(splitLine)>=0 Then
        if splitLine(0)="Monthly Transactions Summary" Then
            mtsFlag = true

        elseif splitLine(0)="Monthly Transactions Details (Payments Only)" Then
            mtdFlag = true
            mtsFlag = false
        end if
    if mtsFlag = true Then
       if ubound(splitLine)>=7 and splitLine(0)<>"Date Posted" and splitLine(0)<>"Monthly Transactions Summary" Then
        'Response.Write(mtsFlag)
            counter=0
            line = split(objFile.ReadLine,"""")

            replacedLine=""
            for each val in line             
            if counter mod 2 =1 Then
        	   	val=Replace(val,",","")
            end if
          
            replacedLine=replacedLine & val
           
            counter=counter + 1
            Next
            Response.Write "<tr>"
            Response.Write "<td><b>" &replacedLine&"</b></td>"& chr(13)
            Response.Write "<tr>"

            line2 = split(replacedLine,",")

            arrParamsMTS1(0) = "mts_DatePosted"
            arrParamsMTS1(1) = 200
            arrParamsMTS1(2) = 150
            arrParamsMTS1(3) = line2(0)
            arrParamsMTS(0) = arrParamsMTS1

            arrParamsMTS2(0) = "mts_Transactionreference"
            arrParamsMTS2(1) = 200
            arrParamsMTS2(2) = 150
            arrParamsMTS2(3) = line2(1)
            arrParamsMTS(1) = arrParamsMTS2

            arrParamsMTS3(0) = "mts_AttorneyDocket"
            arrParamsMTS3(1) = 200
            arrParamsMTS3(2) = 150
            arrParamsMTS3(3) = line2(2)
            arrParamsMTS(2) = arrParamsMTS3

            arrParamsMTS4(0) = "mts_Status"
            arrParamsMTS4(1) = 200
            arrParamsMTS4(2) = 150
            arrParamsMTS4(3) = line2(3)
            arrParamsMTS(3) = arrParamsMTS4

            arrParamsMTS5(0) = "mts_TransactionID"
            arrParamsMTS5(1) = 200
            arrParamsMTS5(2) = 150
            arrParamsMTS5(3) = line2(4)
            arrParamsMTS(4) = arrParamsMTS5

            arrParamsMTS6(0) = "mts_Type"
            arrParamsMTS6(1) = 200
            arrParamsMTS6(2) = 100
            arrParamsMTS6(3) = line2(5)
            arrParamsMTS(5) = arrParamsMTS6

            arrParamsMTS7(0) = "mts_TotalPayment_Refund"
            arrParamsMTS7(1) = 200
            arrParamsMTS7(2) = 150
            arrParamsMTS7(3) = line2(6)
            arrParamsMTS(6) = arrParamsMTS7

            arrParamsMTS8(0) = "mts_CustomerName"
            arrParamsMTS8(1) = 200
            arrParamsMTS8(2) = 150
            arrParamsMTS8(3) = line2(7)
            arrParamsMTS(7) = arrParamsMTS8

            arrParamsMTS9(0) = "fileName"
            arrParamsMTS9(1) = 200
            arrParamsMTS9(2) = 150
            arrParamsMTS9(3) = "CC2017"
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
            Response.Write "<tr>"

            Response.Write "<td><b>" &replacedLine&"</b></td>"& chr(13)

            lineMTD = split(replacedLine,",")


            arrParamsMTD1(0) = "mtd_PaymentDatePosted"
            arrParamsMTD1(1) = 200
            arrParamsMTD1(2) = 50
            arrParamsMTD1(3) = lineMTD(0)
            arrParamsMTD(0) = arrParamsMTD1

            arrParamsMTD2(0) = "mtd_SaleItemDatePosted"
            arrParamsMTD2(1) = 200
            arrParamsMTD2(2) = 50
            arrParamsMTD2(3) = lineMTD(1)
            arrParamsMTD(1) = arrParamsMTD2

            arrParamsMTD3(0) = "mtd_SaleItemReference"
            arrParamsMTD3(1) = 200
            arrParamsMTD3(2) = 50
            arrParamsMTD3(3) = lineMTD(2)
            arrParamsMTD(2) = arrParamsMTD3

            arrParamsMTD4(0) = "mtd_AttorneyDocket"
            arrParamsMTD4(1) = 200
            arrParamsMTD4(2) = 50
            arrParamsMTD4(3) = lineMTD(3)
            arrParamsMTD(3) = arrParamsMTD4

            arrParamsMTD5(0) = "mtd_Status"
            arrParamsMTD5(1) = 200
            arrParamsMTD5(2) = 50
            arrParamsMTD5(3) = lineMTD(4)
            arrParamsMTD(4) = arrParamsMTD5

            arrParamsMTD6(0) = "mtd_TransactionID"
            arrParamsMTD6(1) = 200
            arrParamsMTD6(2) = 50
            arrParamsMTD6(3) = lineMTD(5)
            arrParamsMTD(5) = arrParamsMTD6

            arrParamsMTD7(0) = "mtd_SaleID"
            arrParamsMTD7(1) = 200
            arrParamsMTD7(2) = 50
            arrParamsMTD7(3) = lineMTD(6)
            arrParamsMTD(6) = arrParamsMTD7

            arrParamsMTD8(0) = "mtd_Feecode"
            arrParamsMTD8(1) = 200
            arrParamsMTD8(2) = 50
            arrParamsMTD8(3) = lineMTD(7)
            arrParamsMTD(7) = arrParamsMTD8

            arrParamsMTD9(0) = "mtd_FeeCodeDescription"
            arrParamsMTD9(1) = 200
            arrParamsMTD9(2) = 50
            arrParamsMTD9(3) = lineMTD(8)
            arrParamsMTD(8) = arrParamsMTD9

            arrParamsMTD10(0) = "mtd_ItemPrice"
            arrParamsMTD10(1) = 200
            arrParamsMTD10(2) = 50
            arrParamsMTD10(3) = lineMTD(9)
            arrParamsMTD(9) = arrParamsMTD10

            arrParamsMTD11(0) = "mtd_Quantity"
            arrParamsMTD11(1) = 200
            arrParamsMTD11(2) = 50
            arrParamsMTD11(3) = lineMTD(10)
            arrParamsMTD(10) = arrParamsMTD11

            arrParamsMTD12(0) = "mtd_ItemTotal"
            arrParamsMTD12(1) = 200
            arrParamsMTD12(2) = 50
            arrParamsMTD12(3) = lineMTD(11)
            arrParamsMTD(11) = arrParamsMTD12

            arrParamsMTD13(0) = "mtd_CustomerName"
            arrParamsMTD13(1) = 200
            arrParamsMTD13(2) = 50
            arrParamsMTD13(3) = lineMTD(12)
            arrParamsMTD(12) = arrParamsMTD13

            arrParamsMTD14(0) = "fileName"
            arrParamsMTD14(1) = 200
            arrParamsMTD14(2) = 50
            arrParamsMTD14(3) = "CC2017"
            arrParamsMTD(13) = arrParamsMTD14

            set rsMTD = ObjRpt.RunSPReturnRS("sp_ImportMTD",arrParamsMTD,"")
        End If    
    End if
End if
Loop
objFile.Close
%><caption>Counter: <%=counter%></caption>
</table>