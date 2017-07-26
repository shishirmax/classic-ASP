<%@language="vbscript"%>
<table border="1">
<%
'dim csv_to_read, fso, act, imported_text,split_text, total_imported_text,total_split_text,total_num_imported
Dim csv_to_read,counter,line,fso,objFile,inRow()
Dim iFlagChkFormat,iFlagMTS,iFlagMTD
iFlagChkFormat = false

csv_to_read="C:\file\Credit_Card_1_May_2017.csv"
set fso = createobject("scripting.filesystemobject")
set objFile = fso.opentextfile(csv_to_read,1)
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
    Bufferline = ""
    Currentline = ""
Do Until objFile.AtEndOfStream    
    counter=0
    ContinueLine = 0

    Bufferline = Replace(Bufferline, Chr(13),"") & objFile.ReadLine
    Currentline = Bufferline

    If (Len(Bufferline) - Len(Replace(Bufferline,"""",""))) mod 2 = 1 Then
        Currentline = ""
    else
        Bufferline = ""
        line = split(Currentline,"""")
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
        Response.Write "</tr>"

        splitline = split(replacedLine,",")
        Response.Write(ubound(splitline)&",")
    'Response.Write(splitline(0)&","&splitline(1)&","&splitline(2)&","&splitline(3)&","&splitline(4)&","&splitline(5)&","&splitline(6)&"<br/>")
        if ubound(splitline)>=0 Then
            if splitline(0) = "Monthly Transactions Summary" Then
                iFlagMTS = true
                iFlagChkFormat = true
                Response.Write("MTS")
            elseif splitline(0) = "Monthly Transactions Details (Payments Only)" Then 
                iFlagMTD = true
                iFlagMTS = false
                Response.Write("MTD")
            end if

            if iFlagMTS = true Then
                if ubound(splitline)>=7 and splitline(0)<>"Date Posted" and splitline(0)<>"Monthly Transactions Summary" and splitline(0)<>" " and splitline(0)<>"" Then
                    arrParamsMTS1(0) = "mts_DatePosted"
                    arrParamsMTS1(1) = 129
                    arrParamsMTS1(2) = 250
                    arrParamsMTS1(3) = splitline(0)
                    arrParamsMTS(0) = arrParamsMTS1

                    arrParamsMTS2(0) = "mts_Transactionreference"
                    arrParamsMTS2(1) = 129
                    arrParamsMTS2(2) = 250
                    arrParamsMTS2(3) = splitline(1)
                    arrParamsMTS(1) = arrParamsMTS2

                    arrParamsMTS3(0) = "mts_AttorneyDocket"
                    arrParamsMTS3(1) = 129
                    arrParamsMTS3(2) = 250
                    arrParamsMTS3(3) = splitline(2)
                    arrParamsMTS(2) = arrParamsMTS3

                    arrParamsMTS4(0) = "mts_Status"
                    arrParamsMTS4(1) = 129
                    arrParamsMTS4(2) = 250
                    arrParamsMTS4(3) = splitline(3)
                    arrParamsMTS(3) = arrParamsMTS4

                    arrParamsMTS5(0) = "mts_TransactionID"
                    arrParamsMTS5(1) = 129
                    arrParamsMTS5(2) = 250
                    arrParamsMTS5(3) = splitline(4)
                    arrParamsMTS(4) = arrParamsMTS5

                    arrParamsMTS6(0) = "mts_Type"
                    arrParamsMTS6(1) = 129
                    arrParamsMTS6(2) = 100
                    arrParamsMTS6(3) = splitline(5)
                    arrParamsMTS(5) = arrParamsMTS6

                    arrParamsMTS7(0) = "mts_TotalPayment_Refund"
                    arrParamsMTS7(1) = 129
                    arrParamsMTS7(2) = 250
                    arrParamsMTS7(3) = splitline(6)
                    arrParamsMTS(6) = arrParamsMTS7

                    arrParamsMTS8(0) = "mts_CustomerName"
                    arrParamsMTS8(1) = 129
                    arrParamsMTS8(2) = 250
                    arrParamsMTS8(3) = splitline(7)
                    arrParamsMTS(7) = arrParamsMTS8

                    arrParamsMTS9(0) = "fileName"
                    arrParamsMTS9(1) = 129
                    arrParamsMTS9(2) = 250
                    arrParamsMTS9(3) = "CC2017"
                    arrParamsMTS(8) = arrParamsMTS9

                    set rsMTS = ObjRpt.RunSPReturnRS("sp_ImportMTS",arrParamsMTS,"")
                    If err.number <> 0 Then
                        SESSION("SMALL")="small"
                        RaiseError "readPTO.ASP","RunSPReturnRS",Err.number,Err.Description
                        Response.Write("error")
                    End If
                End If
            Elseif iFlagMTD = true Then
                If ubound(splitline)>=12 and splitline(0)<>"Payment Date Posted" and splitline(0)<>"Monthly Transactions Details (Payments Only)" and splitline(0)<>" " and splitline(0)<>"" Then
                    arrParamsMTD1(0) = "mtd_PaymentDatePosted"
                    arrParamsMTD1(1) = 129
                    arrParamsMTD1(2) = 250
                    arrParamsMTD1(3) = splitline(0)
                    arrParamsMTD(0) = arrParamsMTD1

                    arrParamsMTD2(0) = "mtd_SaleItemDatePosted"
                    arrParamsMTD2(1) = 129
                    arrParamsMTD2(2) = 250
                    arrParamsMTD2(3) = splitline(1)
                    arrParamsMTD(1) = arrParamsMTD2

                    arrParamsMTD3(0) = "mtd_SaleItemReference"
                    arrParamsMTD3(1) = 129
                    arrParamsMTD3(2) = 250
                    arrParamsMTD3(3) = splitline(2)
                    arrParamsMTD(2) = arrParamsMTD3

                    arrParamsMTD4(0) = "mtd_AttorneyDocket"
                    arrParamsMTD4(1) = 129
                    arrParamsMTD4(2) = 250
                    arrParamsMTD4(3) = splitline(3)
                    arrParamsMTD(3) = arrParamsMTD4

                    arrParamsMTD5(0) = "mtd_Status"
                    arrParamsMTD5(1) = 129
                    arrParamsMTD5(2) = 250
                    arrParamsMTD5(3) = splitline(4)
                    arrParamsMTD(4) = arrParamsMTD5

                    arrParamsMTD6(0) = "mtd_TransactionID"
                    arrParamsMTD6(1) = 129
                    arrParamsMTD6(2) = 250
                    arrParamsMTD6(3) = splitline(5)
                    arrParamsMTD(5) = arrParamsMTD6

                    arrParamsMTD7(0) = "mtd_SaleID"
                    arrParamsMTD7(1) = 129
                    arrParamsMTD7(2) = 250
                    arrParamsMTD7(3) = splitline(6)
                    arrParamsMTD(6) = arrParamsMTD7

                    arrParamsMTD8(0) = "mtd_Feecode"
                    arrParamsMTD8(1) = 129
                    arrParamsMTD8(2) = 250
                    arrParamsMTD8(3) = splitline(7)
                    arrParamsMTD(7) = arrParamsMTD8

                    arrParamsMTD9(0) = "mtd_FeeCodeDescription"
                    arrParamsMTD9(1) = 129
                    arrParamsMTD9(2) = 250
                    arrParamsMTD9(3) = splitline(8)
                    arrParamsMTD(8) = arrParamsMTD9

                    arrParamsMTD10(0) = "mtd_ItemPrice"
                    arrParamsMTD10(1) = 129
                    arrParamsMTD10(2) = 250
                    arrParamsMTD10(3) = splitline(9)
                    arrParamsMTD(9) = arrParamsMTD10

                    arrParamsMTD11(0) = "mtd_Quantity"
                    arrParamsMTD11(1) = 129
                    arrParamsMTD11(2) = 250
                    arrParamsMTD11(3) = splitline(10)
                    arrParamsMTD(10) = arrParamsMTD11

                    arrParamsMTD12(0) = "mtd_ItemTotal"
                    arrParamsMTD12(1) = 129
                    arrParamsMTD12(2) = 250
                    arrParamsMTD12(3) = splitline(11)
                    arrParamsMTD(11) = arrParamsMTD12

                    arrParamsMTD13(0) = "mtd_CustomerName"
                    arrParamsMTD13(1) = 129
                    arrParamsMTD13(2) = 250
                    arrParamsMTD13(3) = splitline(12)
                    arrParamsMTD(12) = arrParamsMTD13

                    arrParamsMTD14(0) = "fileName"
                    arrParamsMTD14(1) = 129
                    arrParamsMTD14(2) = 250
                    arrParamsMTD14(3) = "CC2017"
                    arrParamsMTD(13) = arrParamsMTD14

                    set rsMTD = ObjRpt.RunSPReturnRS("sp_ImportMTD",arrParamsMTD,"")
                    If err.number <> 0 Then
                        SESSION("SMALL")="small"
                        RaiseError "readPTO.ASP","RunSPReturnRS",Err.number,Err.Description
                        Response.Write("error")
                    End If 
                End If   
            End If
        End if 
    End if 
Loop
objFile.Close
if iFlagChkFormat=false Then 
    Response.Write ("File is not in the proper format. &nbsp;")
    'Response.Write ("<a href='../test/viewamexfiles.asp'>Back</a>")
    Response.End 
End if
%><caption>Counter: <%=counter%></caption>
</table>