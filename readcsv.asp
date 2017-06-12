<%@language="vbscript"%>
<table border="1">
<%
'dim csv_to_read, fso, act, imported_text,split_text, total_imported_text,total_split_text,total_num_imported
dim csv_to_read,counter,line,fso,objFile,inRow()
csv_to_read="C:\Inetpub\wwwroot\test\Files\USPTO.csv"

set fso = createobject("scripting.filesystemobject")
set objFile = fso.opentextfile(csv_to_read)
objFile.ReadLine
objFile.ReadLine
objFile.ReadLine
objFile.ReadLine



Do Until objFile.AtEndOfStream

    counter=0
    line = split(objFile.ReadLine,"""")

    replacedLine=""
    for each val in line
     
        'Response.Write (val)

    if counter mod 2 =1 Then
	   	val=Replace(val,",","")
    end if
    'Response.Write("Test 1:"&replacedLine)
    replacedLine=replacedLine & val
    'Response.Write("Test 2:"&replacedLine)
   
    counter=counter + 1
    Next
    Response.Write "<tr>"

    Response.Write "<td><b>" &replacedLine&"</b></td>"& chr(13)

    line2 = split(replacedLine,",")
    Response.Write(line2(0)&","&line2(1)&","&line2(2)&","&line2(3)&","&line2(4)&","&line2(5)&","&line2(6)&"<br/>")

Loop
objFile.Close
%><caption>Counter: <%=counter%></caption>
</table>