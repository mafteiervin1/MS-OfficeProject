Sub ReadHtmlFile() '-- this function read our html file 
    Set docActiv = ActiveDocument
    Dim myFile As String, text As String, textline As String, posLat As Integer, posLong As Integer
    myFile = "F:\Facultate\MS-Office\Proiect\htmlfile.html"
    Open myFile For Input As #1
    Do Until EOF(1)
        Line Input #1, textline
    text = text & vbNewLine & textline
    Loop
    Close #1
    
     'MsgBox text
    
    ' docActiv.Paragraphs(1).Range.text = text  '--this line write in worddocument to check if we read the whole html file
    
End Sub
