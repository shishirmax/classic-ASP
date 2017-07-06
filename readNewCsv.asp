<%@language="vbscript"%>
<table border="1">
<%
'dim csv_to_read, fso, act, imported_text,split_text, total_imported_text,total_split_text,total_num_imported
dim csv_to_read,counter,line,fso,objFile,inRow()
Dim splitLine,mtsFlag,mtdFlag
csv_to_read="C:\file\CCFile17.csv"

set fso = createobject("scripting.filesystemobject")
set objFile = fso.opentextfile(csv_to_read)
'objFile.ReadLine
'objFile.ReadLine
'objFile.ReadLine
'objFile.ReadLine

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

Dim ObjRpt
Set ObjRpt = Server.CreateObject("PTFMInvoicing.DBHelper")

Do Until objFile.AtEndOfStream
    splitLine = split(objFile.ReadLine,",")
    'Response.Write(splitLine(0))
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
            Response.Write "<tr>"

            Response.Write "<td><b>" &replacedLine&"</b></td>"& chr(13)

            line2 = split(replacedLine,",")

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
        End If    
    End if
End if
Loop
objFile.Close

%><caption>Counter: <%=counter%></caption>
</table>