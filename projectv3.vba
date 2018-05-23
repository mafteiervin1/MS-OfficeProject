Dim paragraphNumber As Integer
Public Sub ReadHtmlFile() '-- this function read our html file
    Dim html As New HTMLDocument
    Dim docActiv As Document
    Set docActiv = ActiveDocument
    Dim myFile As String, Text As String, textline As String, posLat As Integer, posLong As Integer, plainText As String
    Dim i As Integer, j As Integer
    myFile = "F:\Facultate\MS-Office\Proiect\htmlfile2.html"
    Open myFile For Input As #1
    Do Until EOF(1)
        Line Input #1, textline
    Text = Text & vbNewLine & textline
    Loop
    Close #1
    'docActiv.Paragraphs(1).Range.Text = Text
    With CreateObject("htmlfile") ' aici parsam bodyul
    .Open
    .write Text
    .Close
    
    'docActiv.Paragraphs(1).Range.Text = Text
    'Call searchTitle(Text, docActiv)
    
    docActiv.Paragraphs.Add
    ActiveDocument.Range(Start:=docActiv.Paragraphs(2).Range.Start).Select
            Selection.ClearFormatting
    Dim nTab As Integer, h1Counter As Integer, h2Counter As Integer, h3Counter As Integer, h4Counter As Integer
    Dim h5Counter As Integer, h6Counter As Integer, pCounter As Integer, tableCounter As Integer, ulCounter As Integer, olCounter As Integer
    Dim line As String, WrdArray() As String, font As String
    
    
    nTab = 0
    paragraphNumber = 1
    h1Counter = 0
    h2Counter = 0
    h3Counter = 0
    h4Counter = 0
    h5Counter = 0
    h6Counter = 0
    pCounter = 0
    tableCounter = 0
    ulCounter = 0
    olCounter = 0
 
    WrdArray() = Split(Text, vbCr)
    For i = 0 To UBound(WrdArray)
        If InStr(WrdArray(i), "<p") Or InStr(WrdArray(i), "<P") Then
            Call printP(Text, pCounter, docActiv)
            Call setFont(docActiv, Text, paragraphNumber)
            If InStr(.body.getElementsByTagName("p")(pCounter).innerHTML, "<font") Or InStr(.body.getElementsByTagName("p")(pCounter).innerHTML, "<FONT") Then
                Call setColor(docActiv, .body.getElementsByTagName("p")(pCounter).getElementsByTagName("FONT")(0).getAttribute("color"), paragraphNumber)
                'MsgBox "pcounter" & " " & pCounter & "paragrn" & " " & paragraphNumber
            End If
            
            paragraphNumber = paragraphNumber + 1
            pCounter = pCounter + 1
        
        
        ElseIf InStr(WrdArray(i), "<h1>") Or InStr(WrdArray(i), "<H1>") Then
            If InStr(.body.getElementsByTagName("h1")(h1Counter).innerHTML, "<b><i>") Or InStr(.body.getElementsByTagName("H1")(h1Counter).innerHTML, "<B><I>") Then
                font = "<b><i>"
                MsgBox .body.getElementsByTagName("h1")(h1Counter).innerHTML
            ElseIf InStr(.body.getElementsByTagName("h1")(h1Counter).innerHTML, "<i><b>") Or InStr(.body.getElementsByTagName("H1")(h1Counter).innerHTML, "<I><B>") Then
                    font = "<b><i>"
                    MsgBox .body.getElementsByTagName("h1")(h1Counter).innerHTML
            ElseIf InStr(.body.getElementsByTagName("h1")(h1Counter).innerHTML, "<i>") Or InStr(.body.getElementsByTagName("H1")(h1Counter).innerHTML, "<I>") Then
                    font = "<i>"
                    MsgBox .body.getElementsByTagName("h1")(h1Counter).innerHTML
            ElseIf InStr(.body.getElementsByTagName("h1")(h1Counter).innerHTML, "<b>") Or InStr(.body.getElementsByTagName("H1")(h1Counter).innerHTML, "<B>") Then
                    font = "<b>"
                    MsgBox .body.getElementsByTagName("h1")(h1Counter).innerHTML
            Else: font = "False"
            End If
            Call printH1(Text, h1Counter, docActiv)
            Call setFont(docActiv, font, paragraphNumber)
            Call Headings(docActiv, "<h1>", paragraphNumber)
            paragraphNumber = paragraphNumber + 1
            h1Counter = h1Counter + 1
        
        ElseIf InStr(WrdArray(i), "<h2") Or InStr(WrdArray(i), "<H2") Then
            Call printH2(Text, h2Counter, docActiv)
            Call Headings(docActiv, "<h2>", paragraphNumber)
            paragraphNumber = paragraphNumber + 1
            h2Counter = h2Counter + 1
        
        ElseIf InStr(WrdArray(i), "<h3") Or InStr(WrdArray(i), "<H3") Then
            Call printH3(Text, h3Counter, docActiv)
            Call Headings(docActiv, "<h3>", paragraphNumber)
            paragraphNumber = paragraphNumber + 1
            h3Counter = h3Counter + 1
        
        ElseIf InStr(WrdArray(i), "<h4") Or InStr(WrdArray(i), "<H4") Then
            Call printH4(Text, h4Counter, docActiv)
            Call Headings(docActiv, "<h4>", paragraphNumber)
            paragraphNumber = paragraphNumber + 1
            h4Counter = h4Counter + 1
        
        ElseIf InStr(WrdArray(i), "<h5") Or InStr(WrdArray(i), "<H5") Then
            Call printH5(Text, h5Counter, docActiv)
            Call Headings(docActiv, "<h5>", paragraphNumber)
            paragraphNumber = paragraphNumber + 1
            h5Counter = h5Counter + 1
        
        ElseIf InStr(WrdArray(i), "<h6") Or InStr(WrdArray(i), "<H6") Then
            Call printH6(Text, h6Counter, docActiv)
            Call Headings(docActiv, "<h6>", paragraphNumber)
            paragraphNumber = paragraphNumber + 1
            h6Counter = h6Counter + 1
        
        ElseIf InStr(WrdArray(i), "<table") Or InStr(WrdArray(i), "<TABLE") Then
            Call printTable(docActiv, tableCounter, Text, paragraphNumber)
            paragraphNumber = paragraphNumber + 1
            tableCounter = tableCounter + 1
        
        ElseIf InStr(WrdArray(i), "<ul") Or InStr(WrdArray(i), "<UL") Then
            Call printUL(docActiv, Text, ulCounter, paragraphNumber)
            ulCounter = ulCounter + 1
            paragraphNumber = paragraphNumber + 1
        
        ElseIf InStr(WrdArray(i), "<ol") Or InStr(WrdArray(i), "<OL") Then
            Call printOL(docActiv, Text, olCounter, paragraphNumber)
            olCounter = olCounter + 1
            paragraphNumber = paragraphNumber + 1
        End If
            
    'MsgBox UBound(WrdArray)
    Next
    
    
    'Set Tables = .body.getElementsByTagName("table") ' iteram tabelele din fisier,numaram cate tabele avem
    'For Each Table In Tables
        'Call printTable(docActiv, nTab, Text)
        'nTab = nTab + 1
    'Next
    End With
End Sub
Public Function searchTitle(Text As String, ByRef docActiv As Document)
With CreateObject("htmlfile") ' aici parsam bodyul
    .Open
    .write Text
    .Close
        paragrafe = .body.getElementsByTagName("p")(i).innerText
        docActiv.Content.InsertAfter paragrafe
        docActiv.Content.InsertAfter vbCr
    
    Dim Titlu As String
    Titlu = .body.getElementsByTagName("div")(0).getElementsByTagName("FONT")(0).innerHTML
    Dim Color As String
    docActiv.Paragraphs(1).Range.Text = Titlu
    Color = .body.getElementsByTagName("div")(0).getElementsByTagName("FONT")(0).getAttribute("color")
    Dim Align As String
    Align = .body.getElementsByTagName("div")(0).getAttribute("style").getAttribute("text-align")
    Dim H As String
    H = .body.getElementsByTagName("div")(0).innerHTML
    Call setColor(docActiv, Color, 1)
    Call Headings(docActiv, H, 1)
    Call Aligned(docActiv, Align, 1)
    End With
End Function
    
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
    If LCase(H) = "h1" Then
    docActiv.Paragraphs(i).Range.font.Bold = True
    docActiv.Paragraphs(i).Range.font.Size = 24
    ElseIf LCase(H) = "h2" Then
    docActiv.Paragraphs(i).Range.font.Size = 20
    docActiv.Paragraphs(i).Range.font.Bold = True
    ElseIf LCase(H) = "h3" Then
    docActiv.Paragraphs(i).Range.font.Size = 17
    docActiv.Paragraphs(i).Range.font.Bold = True
    ElseIf LCase(H) = "h4" Then
    docActiv.Paragraphs(i).Range.font.Size = 13
    docActiv.Paragraphs(i).Range.font.Bold = True
    ElseIf LCase(H) = "h5" Then
    docActiv.Paragraphs(i).Range.font.Size = 10
    docActiv.Paragraphs(i).Range.font.Bold = True
    ElseIf LCase(H) = "h6" Then
    docActiv.Paragraphs(i).Range.font.Size = 7
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
    docActiv.Paragraphs(i).Range.font.Color = RGB(lngOut1, lngOut2, lngOut3)

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
    tbl.Rows(i).Cells(j).Range.font.Color = RGB(lngOut1, lngOut2, lngOut3)
    'MsgBox i & " " & j

End Function
Public Function setCellFont(ByRef docActiv As Document, font As String, ByRef tbl As Variant, ByRef i As Integer, ByRef j As Integer)

If InStr(font, "<b><i>") Then
    tbl.Rows(i).Cells(j).Range.font.Bold = True
    tbl.Rows(i).Cells(j).Range.font.Italic = True
ElseIf InStr(font, "<b>") Then
    tbl.Rows(i).Cells(j).Range.font.Bold = True
ElseIf InStr(font, "<i>") Then
    tbl.Rows(i).Cells(j).Range.font.Italic = True
End If


End Function
Public Function setFont(ByRef docActiv As Document, Text As String, i As Integer)
If InStr(Text, "<b><i>") Then
    docActiv.Paragraphs(i).Range.font.Bold = True
    docActiv.Paragraphs(i).Range.font.Italic = True
ElseIf InStr(Text, "<b>") Then
    docActiv.Paragraphs(i).Range.font.Bold = True
ElseIf InStr(Text, "<i>") Then
    docActiv.Paragraphs(i).Range.font.Italic = True
Else
    docActiv.Paragraphs(i).Range.font.Bold = False
    docActiv.Paragraphs(i).Range.font.Italic = False
End If
End Function
Public Sub printP(Text As String, i As Integer, ByRef docActiv As Document)
    With CreateObject("htmlfile") ' aici parsam bodyul
    .Open
    .write Text
    .Close
        paragrafe = .body.getElementsByTagName("p")(i).innerText
        'MsgBox i & " " & .body.getElementsByTagName("p")(i).innerText
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

Public Function printOL(ByRef docActiv As Document, Text As String, i As Integer, ByRef pnumber As Integer)
    With CreateObject("htmlfile") ' aici parsam bodyul
    Dim iLIordonata As Integer, paragrafeLI As Object
    
        .Open
        .write Text
        .Close
        Set paragrafeLI = .body.getElementsByTagName("ol")(i).getElementsByTagName("li")
            iLIordonata = 1
            For Each paragrafLI In paragrafeLI
                docActiv.Content.InsertAfter iLIordonata & "." & paragrafLI.innerText
                docActiv.Content.InsertAfter vbCr
                pnumber = pnumber + 1
            Next
End With

End Function
Public Function printUL(ByRef docActiv As Document, Text As String, i As Integer, ByRef pnumber As Integer)
    With CreateObject("htmlfile") ' aici parsam bodyul
        .Open
        .write Text
        .Close
        Set paragrafeUL = .body.getElementsByTagName("ul")
            Set paragrafeLI = .body.getElementsByTagName("ul")(i).getElementsByTagName("li")
            For Each paragrafLI In paragrafeLI
                docActiv.Content.InsertAfter paragrafLI.innerText
                docActiv.Content.InsertAfter vbCr
                pnumber = pnumber + 1
                
            Next
     End With
End Function
Public Function printTable(ByRef docActiv As Document, tableNumber As Integer, htmlfile As String, ByRef paragraphNumber As Integer)
    Dim Color As String, i As Integer, j As Integer, docActivPnr As Integer, font As String
    
    With CreateObject("htmlfile") ' aici parsam bodyul
        .Open
        .write htmlfile
        .Close
        nCol = 0
        nTab = 0

        Set Rows = .body.getElementsByTagName("table")(tableNumber).getElementsByTagName("tr") 'iteram liniile din tabelul ntab,numaram cate linii avem
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
        paragraphNumber = paragraphNumber + 1
        'resetam formatarile anterioare
        ActiveDocument.Range(Start:=docActiv.Paragraphs(docActivPnr).Range.Start).Select
            Selection.ClearFormatting
            
        'cream un tabel
        
        Set tbl = docActiv.Tables.Add(docActiv.Paragraphs(docActivPnr).Range, _
                                      NumRows:=nRow, _
                                      NumColumns:=nCol / nRow)
        paragraphNumber = paragraphNumber + 1
        i = 1
        Set Rows = .body.getElementsByTagName("table")(tableNumber).getElementsByTagName("tr")
        
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
                If InStr(Column.innerHTML, "<b><i>") Or InStr(Column.innerHTML, "<B><I>") Then
                    font = "<b><i>"
                    'MsgBox Column.innerHTML
                ElseIf InStr(Column.innerHTML, "<i><b>") Or InStr(Column.innerHTML, "<I><B>") Then
                    font = "<b><i>"
                    'MsgBox Column.innerHTML
                ElseIf InStr(Column.innerHTML, "<i>") Or InStr(Column.innerHTML, "<I>") Then
                    font = "<i>"
                    'MsgBox Column.innerHTML
                ElseIf InStr(Column.innerHTML, "<b>") Or InStr(Column.innerHTML, "<B>") Then
                    font = "<b>"
                    'MsgBox Column.innerHTML
                Else: font = "False"
                End If
                ActiveDocument.Range(Start:=docActiv.Paragraphs(docActiv.Paragraphs.Count).Range.Start).Select
                    Selection.ClearFormatting
                Call setCellFont(docActiv, font, tbl, i, j)
                'MsgBox Column.innerHTML
                j = j + 1
                paragraphNumber = paragraphNumber + 1
                
            Next
            i = i + 1
            paragraphNumber = paragraphNumber + 1
        Next
        
        With tbl.Borders
            .OutsideLineStyle = wdLineStyleDouble
            .InsideLineStyle = wdLineStyleDouble
            End With
        End With
End Function


