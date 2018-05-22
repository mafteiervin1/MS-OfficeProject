Public Sub ReadHtmlFile() '-- this function read our html file
    Dim html As New HTMLDocument
    Dim docActiv As Document
    Set docActiv = ActiveDocument
    Dim myFile As String, Text As String, textline As String, posLat As Integer, posLong As Integer, plainText As String
    Dim i As Integer, j As Integer
    myFile = "F:\Facultate\MS-Office\Proiect\htmlfile.html"
    Open myFile For Input As #1
    Do Until EOF(1)
        Line Input #1, textline
    Text = Text & vbNewLine & textline
    Loop
    Close #1
    docActiv.Paragraphs(1).Range.Text = Text
    With CreateObject("htmlfile") ' aici parsam bodyul
    .Open
    .write Text
    .Close
    
    Dim Titlu As String
    Titlu = .body.getElementsByTagName("div")(0).getElementsByTagName("FONT")(0).innerHTML
    Dim Color As String
    docActiv.Paragraphs(1).Range.Text = Titlu
    Color = .body.getElementsByTagName("div")(0).getElementsByTagName("FONT")(0).getAttribute("color")
    Dim Align As String
    Align = .body.getElementsByTagName("div")(0).getAttribute("style").getAttribute("text-align")
    Dim H As String
    H = .body.getElementsByTagName("div")(0).innerHTML
    Call Font(docActiv, Color, 1)
    Call Headings(docActiv, H, 1)
    Call Aligned(docActiv, Align, 1)
    
    
    
    Dim Elem As Object
    Dim nRow As Integer, nCol As Integer, nTab As Integer, docActivPnr As Integer
    
    nRow = 0
    nCol = 0
    nTab = 0
    
    Set Tables = .body.getElementsByTagName("table") ' iteram tabelele din fisier,numaram cate tabele avem
    For Each Table In Tables
        nTab = nTab + 1
        Set Rows = Table.getElementsByTagName("tr") 'iteram liniile din tabelul ntab,numaram cate linii avem
        'docActiv.Paragraphs(1).Range.text = Table.innerHTML
        nRow = 0
        For Each Row In Rows
            nRow = nRow + 1
            Set Columns = Row.getElementsByTagName("td") 'iteram coloanele din linia nrow,numaram cate coloane avem
            'docActiv.Paragraphs(1).Range.text = Columns(0).innerHTML
                For Each Column In Columns
                    nCol = nCol + 1
                    'MsgBox Column.innerText
                Next
        Next
        'MsgBox nTab & " " & nRow & " " & nCol / nRow
        
        docActivPnr = docActiv.Paragraphs.Count + 1
        docActiv.Paragraphs.Add
        
        'resetam formatarile anterioare
        ActiveDocument.Range(Start:=docActiv.Paragraphs(docActivPnr).Range.Start).Select
            Selection.ClearFormatting
            
        'cream un tabel
        
        Set tbl = docActiv.Tables.Add(docActiv.Paragraphs(docActivPnr).Range, _
                                      NumRows:=nRow, _
                                      NumColumns:=nCol / nRow)
        i = 1
        Set Rows = Table.getElementsByTagName("tr")
        
        For Each Row In Rows
            j = 1
            Set Columns = Row.getElementsByTagName("td")
            For Each Column In Columns
                tbl.Rows(i).Cells(j).Range.Text = Column.innerText
                If InStr(Column.innerHTML, "href") Then
                    docActiv.Hyperlinks.Add Anchor:=tbl.Rows(i).Cells(j).Range, Address:=Column.getElementsByTagName("a")(0).href
                    'MsgBox Column.getElementsByTagName("a")(0).href
                End If
                If InStr(Column.innerHTML, "font") Or InStr(Column.innerHTML, "FONT") Then
                    'MsgBox Column.getElementsByTagName("FONT")(0).getAttribute("color")
                    Color = Column.getElementsByTagName("FONT")(0).getAttribute("color")
                    Call setCellColor(docActiv, Color, tbl, i, j)
                End If
                'MsgBox Column.innerHTML
                j = j + 1
            Next
            i = i + 1
        Next
        
        With tbl.Borders
            .OutsideLineStyle = wdLineStyleDouble
            .InsideLineStyle = wdLineStyleDouble
            End With
        
    Next
    
    Set paragrafeUL = .body.getElementsByTagName("ul")
    jUL = 0
    For Each paragrafUL In paragrafeUL
        Set paragrafeLI = .body.getElementsByTagName("ul")(jUL).getElementsByTagName("li")
        For Each paragrafLI In paragrafeLI
            docActiv.Content.InsertAfter paragrafLI.innerText
            docActiv.Content.InsertAfter vbCr
        Next
       jUL = jUL + 1
       
    Next
    
    Set paragrafeOL = .body.getElementsByTagName("ol")
    jOL = 0
    For Each paragrafOL In paragrafeOL
        Set paragrafeLI = .body.getElementsByTagName("ol")(jOL).getElementsByTagName("li")
        iLIordonata = 1
        For Each paragrafLI In paragrafeLI
            docActiv.Content.InsertAfter iLIordonata & "." & paragrafLI.innerText
            docActiv.Content.InsertAfter vbCr
            iLIordonata = iLIordonata + 1
        Next
       jOL = jOL + 1
       
    Next

    End With

    'RemoveHTML (text)
     'MsgBox text
    'docActiv.Paragraphs(1).Range.text = plainText
    'docActiv.Paragraphs(1).Range.text = text  '--this line write in worddocument to check if we read the whole html file
    
End Sub
Public Function Aligned(ByRef docActiv As Document, Align As String, i As Integer)

    If Align = "center" Then
    docActiv.Paragraphs(i).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    ElseIf Align = "left" Then
    docActiv.Paragraphs(i).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
    ElseIf Align = "right" Then
    docActiv.Paragraphs(i).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
    End If

End Function

Public Function Headings(ByRef docActiv As Document, H As String, i As Integer)

    H = Left(H, 3)
    H = Right(H, 2)
    If H = "H1" Then
    docActiv.Paragraphs(i).Range.Font.Bold = True
    docActiv.Paragraphs(i).Range.Font.Size = 24
    ElseIf H = "H2" Then
    docActiv.Paragraphs(i).Range.Font.Size = 20
    docActiv.Paragraphs(i).Range.Font.Bold = True
    ElseIf H = "H3" Then
    docActiv.Paragraphs(i).Range.Font.Size = 17
    docActiv.Paragraphs(i).Range.Font.Bold = True
    ElseIf H = "H4" Then
    docActiv.Paragraphs(i).Range.Font.Size = 13
    docActiv.Paragraphs(i).Range.Font.Bold = True
    ElseIf H = "H5" Then
    docActiv.Paragraphs(i).Range.Font.Size = 10
    docActiv.Paragraphs(i).Range.Font.Bold = True
    ElseIf H = "H6" Then
    docActiv.Paragraphs(i).Range.Font.Size = 7
    End If
    
End Function



Public Function setColor(ByRef docActiv As Document, Color As String, i As Integer)
   
    Dim strHex1 As String
    Dim strHex2 As String
    Dim strHex3 As String
    Dim lngOut1 As Long
    Dim lngOut2 As Long
    Dim lngOut3 As Long
    
    strHex1 = Right(Left(Color, 3), 2)
    strHex2 = Right(Left(Color, 5), 2)
    strHex3 = Right(Left(Color, 7), 2)
    lngOut1 = CLng("&H" & strHex1)
    lngOut2 = CLng("&H" & strHex2)
    lngOut3 = CLng("&H" & strHex3)
    docActiv.Paragraphs(i).Range.Font.Color = RGB(lngOut1, lngOut2, lngOut3)

End Function
Public Function setCellColor(ByRef docActiv As Document, Color As String, ByRef tbl As Variant, ByRef i As Integer, ByRef j As Integer)
   
    Dim strHex1 As String
    Dim strHex2 As String
    Dim strHex3 As String
    Dim lngOut1 As Long
    Dim lngOut2 As Long
    Dim lngOut3 As Long
    
    strHex1 = Right(Left(Color, 3), 2)
    strHex2 = Right(Left(Color, 5), 2)
    strHex3 = Right(Left(Color, 7), 2)
    lngOut1 = CLng("&H" & strHex1)
    lngOut2 = CLng("&H" & strHex2)
    lngOut3 = CLng("&H" & strHex3)
    tbl.Rows(i).Cells(j).Range.Font.Color = RGB(lngOut1, lngOut2, lngOut3)
    'MsgBox i & " " & j

End Function
Public Sub printP(Text As String, i As Integer, ByRef docActiv As Document)
    With CreateObject("htmlfile") ' aici parsam bodyul
    .Open
    .write Text
    .Close
        paragrafe = .body.getElementsByTagName("p")(i).innerText
        docActiv.Content.InsertAfter paragrafe
        docActiv.Content.InsertAfter vbCr
    End With
End Sub
Public Sub printH1(Text As String, i As Integer, ByRef docActiv As Document)
    With CreateObject("htmlfile") ' aici parsam bodyul
    .Open
    .write Text
    .Close
        paragrafeH1 = .body.getElementsByTagName("h1")(i).innerText
        docActiv.Content.InsertAfter paragrafeH1
        docActiv.Content.InsertAfter vbCr
    End With
End Sub
Public Sub printH2(Text As String, i As Integer, ByRef docActiv As Document)
    With CreateObject("htmlfile") ' aici parsam bodyul
    .Open
    .write Text
    .Close
        paragrafeH2 = .body.getElementsByTagName("h2")(i).innerText
        docActiv.Content.InsertAfter paragrafeH2
        docActiv.Content.InsertAfter vbCr
    End With
End Sub

Public Sub printH3(Text As String, i As Integer, ByRef docActiv As Document)
    With CreateObject("htmlfile") ' aici parsam bodyul
    .Open
    .write Text
    .Close
        paragrafeH3 = .body.getElementsByTagName("h3")(i).innerText
        docActiv.Content.InsertAfter paragrafeH3
        docActiv.Content.InsertAfter vbCr
    End With
End Sub
Public Sub printH4(Text As String, i As Integer, ByRef docActiv As Document)
    With CreateObject("htmlfile") ' aici parsam bodyul
    .Open
    .write Text
    .Close
        paragrafeH4 = .body.getElementsByTagName("h4")(i).innerText
        docActiv.Content.InsertAfter paragrafeH4
        docActiv.Content.InsertAfter vbCr
    End With
End Sub
Public Sub printH5(Text As String, i As Integer, ByRef docActiv As Document)
    With CreateObject("htmlfile") ' aici parsam bodyul
    .Open
    .write Text
    .Close
        paragrafeH5 = .body.getElementsByTagName("h5")(i).innerText
        docActiv.Content.InsertAfter paragrafeH5
        docActiv.Content.InsertAfter vbCr
    End With
End Sub
Public Sub printH6(Text As String, i As Integer, ByRef docActiv As Document)
    With CreateObject("htmlfile") ' aici parsam bodyul
    .Open
    .write Text
    .Close
        paragrafeH6 = .body.getElementsByTagName("h6")(i).innerText
        docActiv.Content.InsertAfter paragrafeH6
        docActiv.Content.InsertAfter vbCr
    End With
End Sub
Public Function setFont(ByRef docActiv As Document, Text As String, i As Integer)
If InStr(Text, "<b><i>") Then
    docActiv.Paragraphs(i).Range.Font.Bold = True
    docActiv.Paragraphs(i).Range.Font.Italic = True
ElseIf InStr(Text, "<b>") Then
    docActiv.Paragraphs(i).Range.Font.Bold = True
ElseIf InStr(Text, "<i>") Then
    docActiv.Paragraphs(i).Range.Font.Italic = True
End If
End Function


