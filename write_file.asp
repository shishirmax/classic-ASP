    <% Option Explicit %>  
    <html>  
    <head>  
    <title>Write File Sample</title>  
    </head>  
    <body>  
    <%  
    on error resume next

    Dim objFSO, objStream,i  
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    'Response.Write(err.Number & " - "& err.Description & "<br>")
    Set objStream = objFSO.CreateTextFile(Server.MapPath("MyFile.txt"),true)  
    Response.Write(err.Number & " - "& err.Description & "<br>")

    objStream.WriteLine("--------------------")
    objStream.WriteLine("Make me cool ")  
    objStream.WriteLine("I also love Designing")  
    Response.write("Writing")
    Response.Write("<br>")  
    Response.Write(err.Number & " - "& err.Description & "<br>")
    if objFSO.FolderExists("c:\inetpub") = true then
    Response.Write("Folder c:\inetpub exists.....")
    end if
    objStream.Close  
    Set objStream = Nothing  
    Set objFSO = Nothing  
    %>  
    </body>  
    </html>  