Public Sub ReadHtmlFile() '-- this function read our html file
    Dim html As New HTMLDocument
    Set docActiv = ActiveDocument
    Dim myFile As String, text As String, textline As String, posLat As Integer, posLong As Integer, plainText As String
    Dim i As Integer, j As Integer
    myFile = "F:\Facultate\MS-Office\Proiect\htmlfile.html"
    Open myFile For Input As #1
    Do Until EOF(1)
        Line Input #1, textline
    text = text & vbNewLine & textline
    Loop
    Close #1
    With CreateObject("htmlfile") ' aici parsam bodyul
    .Open
    .write text
    .Close
    Dim Elem As Object
    Dim nRow As Integer, nCol As Integer, nTab As Integer
    nRow = 0
    nCol = 0
    nTab = 0
    Set Tables = .body.getElementsByTagName("table")
    For Each Table In Tables
        nTab = nTab + 1
        Set Rows = Table.getElementsByTagName("tr")
        'docActiv.Paragraphs(1).Range.text = Table.innerHTML
        nRow = 0
        For Each Row In Rows
            nRow = nRow + 1
            Set Columns = Row.getElementsByTagName("td")
            'docActiv.Paragraphs(1).Range.text = Columns(0).innerHTML
                For Each Column In Columns
                    nCol = nCol + 1
                    'MsgBox Column.innerText
                Next
        Next
        'MsgBox nTab & " " & nRow & " " & nCol / nRow
        
        Set tbl = docActiv.Tables.Add(docActiv.Paragraphs(nTab).Range, _
                                      NumRows:=nRow, _
                                      NumColumns:=nCol / nRow)
        i = 1
        Set Rows = Table.getElementsByTagName("tr")
        For Each Row In Rows
            j = 1
            Set Columns = Row.getElementsByTagName("td")
            For Each Column In Columns
                tbl.Rows(i).Cells(j).Range.text = Column.innerText
                If InStr(Column.innerHTML, "href") Then
                    docActiv.Hyperlinks.Add Anchor:=tbl.Rows(i).Cells(j).Range, Address:=Column.getElementsByTagName("a")(0).href
                    'MsgBox Column.getElementsByTagName("a")(0).href
                End If
                'MsgBox Column.innerHTML
                j = j + 1
            Next
            i = i + 1
        Next
        
    Next
            
            
            
    'MsgBox "text=" & .body.getElementsByTagName("tr")(0).innerHTML
    'docActiv.Paragraphs(1).Range.text = .body.innerHTML
    End With

    'RemoveHTML (text)
     'MsgBox text
    'docActiv.Paragraphs(1).Range.text = plainText
    'docActiv.Paragraphs(1).Range.text = text  '--this line write in worddocument to check if we read the whole html file
    
End Sub

'https://codingislove.com/parse-html-in-excel-vba/
'https://www.wiseowl.co.uk/blog/s393/scrape-website-html.htm
