<%@language="vbscript"%>
<%
Dim fso

set fso = createobject("scripting.filesystemobject")
csv_to_read="C:\file\MasterCard-June28_2017.txt"
'text = fso.opentextfile(csv_to_read).ReadAll
	'For Each line in split(text,vbcr)
			'response.write(line)
	'Next
set infile = fso.opentextfile(csv_to_read)
set outfile = fso.opentextfile(csv_to_read & ".tmp",8,2)

Do Until infile.AtEndOfStream
	c = infile.Read(1)
	if c = vbCr Then
		outfile.write vbcrlf
	Else
		outfile.write c
	End If
Loop

infile.close
outfile.close

fso.DeleteFile csv_to_read,True
fso.MoveFile csv_to_read & ".tmp",csv_to_read
%>