Dim paragraphNumber As Integer
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
    'docActiv.Paragraphs(1).Range.Text = Text
    With CreateObject("htmlfile") ' aici parsam bodyul
    .Open
    .write Text
    .Close
    
    'docActiv.Paragraphs(1).Range.Text = Text
    'Call searchTitle(Text, docActiv)
    
    'docActiv.Paragraphs.Add
    'ActiveDocument.Range(Start:=docActiv.Paragraphs(2).Range.Start).Select
            'Selection.ClearFormatting
    Dim nTab As Integer, h1Counter As Integer, h2Counter As Integer, h3Counter As Integer, h4Counter As Integer
    Dim h5Counter As Integer, h6Counter As Integer, pCounter As Integer, tableCounter As Integer, ulCounter As Integer, olCounter As Integer
    Dim line As String, WrdArray() As String, font As String, link As String
    
    
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
            If InStr(.body.getElementsByTagName("p")(pCounter).innerHTML, "<b><i>") Or InStr(.body.getElementsByTagName("P")(pCounter).innerHTML, "<B><I>") Then
                font = "<b><i>"
            ElseIf InStr(.body.getElementsByTagName("p")(pCounter).innerHTML, "<i><b>") Or InStr(.body.getElementsByTagName("P")(pCounter).innerHTML, "<I><B>") Then
                    font = "<b><i>"
            ElseIf InStr(.body.getElementsByTagName("p")(pCounter).innerHTML, "<i>") Or InStr(.body.getElementsByTagName("P")(pCounter).innerHTML, "<I>") Then
                    font = "<i>"
            ElseIf InStr(.body.getElementsByTagName("p")(pCounter).innerHTML, "<b>") Or InStr(.body.getElementsByTagName("P")(pCounter).innerHTML, "<B>") Then
                    font = "<b>"
            Else: font = "False"
            End If
            Call setFont(docActiv, font, docActiv.Paragraphs.Count - 1)
            If InStr(.body.getElementsByTagName("p")(pCounter).innerHTML, "<font") Or InStr(.body.getElementsByTagName("p")(pCounter).innerHTML, "<FONT") Then
                Call setColor(docActiv, .body.getElementsByTagName("p")(pCounter).getElementsByTagName("FONT")(0).getAttribute("color"), docActiv.Paragraphs.Count - 1)
            End If
            If InStr(LCase(.body.getElementsByTagName("p")(pCounter).innerHTML), "<a") Then
                link = .body.getElementsByTagName("p")(pCounter).getElementsByTagName("a")(0).href
                docActiv.Hyperlinks.Add Anchor:=docActiv.Paragraphs(docActiv.Paragraphs.Count - 1).Range, Address:=link
            End If
            paragraphNumber = paragraphNumber + 1
            docActiv.Paragraphs.Add
            ActiveDocument.Range(Start:=docActiv.Paragraphs(docActiv.Paragraphs.Count).Range.Start).Select
                Selection.ClearFormatting
            pCounter = pCounter + 1
        
        
        ElseIf InStr(WrdArray(i), "<h1>") Or InStr(WrdArray(i), "<H1>") Then
            If InStr(.body.getElementsByTagName("h1")(h1Counter).innerHTML, "<b><i>") Or InStr(.body.getElementsByTagName("H1")(h1Counter).innerHTML, "<B><I>") Then
                font = "<b><i>"
            ElseIf InStr(.body.getElementsByTagName("h1")(h1Counter).innerHTML, "<i><b>") Or InStr(.body.getElementsByTagName("H1")(h1Counter).innerHTML, "<I><B>") Then
                    font = "<b><i>"
            ElseIf InStr(.body.getElementsByTagName("h1")(h1Counter).innerHTML, "<i>") Or InStr(.body.getElementsByTagName("H1")(h1Counter).innerHTML, "<I>") Then
                    font = "<i>"
            ElseIf InStr(.body.getElementsByTagName("h1")(h1Counter).innerHTML, "<b>") Or InStr(.body.getElementsByTagName("H1")(h1Counter).innerHTML, "<B>") Then
                    font = "<b>"
            Else: font = "False"
            End If
            Call printH1(Text, h1Counter, docActiv)
            Call setFont(docActiv, font, docActiv.Paragraphs.Count - 1)
            Call Headings(docActiv, "<h1>", docActiv.Paragraphs.Count - 1)
            If InStr(.body.getElementsByTagName("h1")(h1Counter).innerHTML, "<font") Or InStr(.body.getElementsByTagName("h1")(h1Counter).innerHTML, "<FONT") Then
                Call setColor(docActiv, .body.getElementsByTagName("h1")(h1Counter).getElementsByTagName("FONT")(0).getAttribute("color"), docActiv.Paragraphs.Count - 1)
            End If
            If InStr(LCase(.body.getElementsByTagName("h1")(h1Counter).innerHTML), "<a") Then
                link = .body.getElementsByTagName("h1")(h1Counter).getElementsByTagName("a")(0).href
                docActiv.Hyperlinks.Add Anchor:=docActiv.Paragraphs(docActiv.Paragraphs.Count - 1).Range, Address:=link
            End If
            paragraphNumber = paragraphNumber + 1
            docActiv.Paragraphs.Add
            ActiveDocument.Range(Start:=docActiv.Paragraphs(docActiv.Paragraphs.Count).Range.Start).Select
                Selection.ClearFormatting
            h1Counter = h1Counter + 1
        
        ElseIf InStr(WrdArray(i), "<h2") Or InStr(WrdArray(i), "<H2") Then
        If InStr(.body.getElementsByTagName("h2")(h2Counter).innerHTML, "<b><i>") Or InStr(.body.getElementsByTagName("H2")(h2Counter).innerHTML, "<B><I>") Then
                font = "<b><i>"
            ElseIf InStr(.body.getElementsByTagName("h2")(h2Counter).innerHTML, "<i><b>") Or InStr(.body.getElementsByTagName("H2")(h2Counter).innerHTML, "<I><B>") Then
                    font = "<b><i>"
            ElseIf InStr(.body.getElementsByTagName("h2")(h2Counter).innerHTML, "<i>") Or InStr(.body.getElementsByTagName("H2")(h2Counter).innerHTML, "<I>") Then
                    font = "<i>"
            ElseIf InStr(.body.getElementsByTagName("h2")(h2Counter).innerHTML, "<b>") Or InStr(.body.getElementsByTagName("H2")(h2Counter).innerHTML, "<B>") Then
                    font = "<b>"
            Else: font = "False"
            End If
            Call printH2(Text, h2Counter, docActiv)
            Call setFont(docActiv, font, docActiv.Paragraphs.Count - 1)
            Call Headings(docActiv, "<h2>", docActiv.Paragraphs.Count - 1)
            If InStr(.body.getElementsByTagName("h2")(h2Counter).innerHTML, "<font") Or InStr(.body.getElementsByTagName("h2")(h2Counter).innerHTML, "<FONT") Then
                Call setColor(docActiv, .body.getElementsByTagName("h2")(h2Counter).getElementsByTagName("FONT")(0).getAttribute("color"), docActiv.Paragraphs.Count - 1)
            End If
            If InStr(LCase(.body.getElementsByTagName("h2")(h2Counter).innerHTML), "<a") Then
                link = .body.getElementsByTagName("h2")(h2Counter).getElementsByTagName("a")(0).href
                docActiv.Hyperlinks.Add Anchor:=docActiv.Paragraphs(docActiv.Paragraphs.Count - 1).Range, Address:=link
            End If
            paragraphNumber = paragraphNumber + 1
            docActiv.Paragraphs.Add
            ActiveDocument.Range(Start:=docActiv.Paragraphs(docActiv.Paragraphs.Count).Range.Start).Select
                Selection.ClearFormatting
            h2Counter = h2Counter + 1
        
        ElseIf InStr(WrdArray(i), "<h3") Or InStr(WrdArray(i), "<H3") Then
            Call printH3(Text, h3Counter, docActiv)
            If InStr(.body.getElementsByTagName("h3")(h3Counter).innerHTML, "<b><i>") Or InStr(.body.getElementsByTagName("H3")(h3Counter).innerHTML, "<B><I>") Then
                font = "<b><i>"
            ElseIf InStr(.body.getElementsByTagName("h3")(h3Counter).innerHTML, "<i><b>") Or InStr(.body.getElementsByTagName("H3")(h3Counter).innerHTML, "<I><B>") Then
                    font = "<b><i>"
            ElseIf InStr(.body.getElementsByTagName("h3")(h3Counter).innerHTML, "<i>") Or InStr(.body.getElementsByTagName("H3")(h3Counter).innerHTML, "<I>") Then
                    font = "<i>"
            ElseIf InStr(.body.getElementsByTagName("h3")(h3Counter).innerHTML, "<b>") Or InStr(.body.getElementsByTagName("H3")(h3Counter).innerHTML, "<B>") Then
                    font = "<b>"

            Else: font = "False"
            End If
            Call setFont(docActiv, font, docActiv.Paragraphs.Count - 1)
            Call Headings(docActiv, "<h3>", docActiv.Paragraphs.Count - 1)
            If InStr(.body.getElementsByTagName("h3")(h3Counter).innerHTML, "<font") Or InStr(.body.getElementsByTagName("h3")(h3Counter).innerHTML, "<FONT") Then
                Call setColor(docActiv, .body.getElementsByTagName("h3")(h3Counter).getElementsByTagName("FONT")(0).getAttribute("color"), docActiv.Paragraphs.Count - 1)
            End If
            If InStr(LCase(.body.getElementsByTagName("h3")(h3Counter).innerHTML), "<a") Then
                link = .body.getElementsByTagName("h3")(h3Counter).getElementsByTagName("a")(0).href
                docActiv.Hyperlinks.Add Anchor:=docActiv.Paragraphs(docActiv.Paragraphs.Count - 1).Range, Address:=link
            End If
            paragraphNumber = paragraphNumber + 1
            docActiv.Paragraphs.Add
            ActiveDocument.Range(Start:=docActiv.Paragraphs(docActiv.Paragraphs.Count).Range.Start).Select
                Selection.ClearFormatting
            h3Counter = h3Counter + 1
        
        ElseIf InStr(WrdArray(i), "<h4") Or InStr(WrdArray(i), "<H4") Then
            Call printH4(Text, h4Counter, docActiv)
            If InStr(.body.getElementsByTagName("h4")(h4Counter).innerHTML, "<b><i>") Or InStr(.body.getElementsByTagName("H4")(h4Counter).innerHTML, "<B><I>") Then
                font = "<b><i>"

            ElseIf InStr(.body.getElementsByTagName("h4")(h4Counter).innerHTML, "<i><b>") Or InStr(.body.getElementsByTagName("H4")(h4Counter).innerHTML, "<I><B>") Then
                    font = "<b><i>"
        
            ElseIf InStr(.body.getElementsByTagName("h4")(h4Counter).innerHTML, "<i>") Or InStr(.body.getElementsByTagName("H4")(h4Counter).innerHTML, "<I>") Then
                    font = "<i>"
      
            ElseIf InStr(.body.getElementsByTagName("h4")(h4Counter).innerHTML, "<b>") Or InStr(.body.getElementsByTagName("H4")(h4Counter).innerHTML, "<B>") Then
                    font = "<b>"
  
            Else: font = "False"
            End If
            Call setFont(docActiv, font, docActiv.Paragraphs.Count - 1)
            Call Headings(docActiv, "<h4>", docActiv.Paragraphs.Count - 1)
            If InStr(.body.getElementsByTagName("h4")(h4Counter).innerHTML, "<font") Or InStr(.body.getElementsByTagName("h4")(h4Counter).innerHTML, "<FONT") Then
                Call setColor(docActiv, .body.getElementsByTagName("h4")(h4Counter).getElementsByTagName("FONT")(0).getAttribute("color"), docActiv.Paragraphs.Count - 1)
            End If
            If InStr(LCase(.body.getElementsByTagName("h4")(h4Counter).innerHTML), "<a") Then
                link = .body.getElementsByTagName("h4")(h4Counter).getElementsByTagName("a")(0).href
                docActiv.Hyperlinks.Add Anchor:=docActiv.Paragraphs(docActiv.Paragraphs.Count - 1).Range, Address:=link
            End If
            paragraphNumber = paragraphNumber + 1
            docActiv.Paragraphs.Add
            ActiveDocument.Range(Start:=docActiv.Paragraphs(docActiv.Paragraphs.Count).Range.Start).Select
                Selection.ClearFormatting
            h4Counter = h4Counter + 1
        
        ElseIf InStr(WrdArray(i), "<h5") Or InStr(WrdArray(i), "<H5") Then
            Call printH5(Text, h5Counter, docActiv)
            If InStr(.body.getElementsByTagName("h5")(h5Counter).innerHTML, "<b><i>") Or InStr(.body.getElementsByTagName("H5")(h5Counter).innerHTML, "<B><I>") Then
                font = "<b><i>"
            ElseIf InStr(.body.getElementsByTagName("h5")(h5Counter).innerHTML, "<i><b>") Or InStr(.body.getElementsByTagName("H5")(h5Counter).innerHTML, "<I><B>") Then
                    font = "<b><i>"
            ElseIf InStr(.body.getElementsByTagName("h5")(h5Counter).innerHTML, "<i>") Or InStr(.body.getElementsByTagName("H5")(h5Counter).innerHTML, "<I>") Then
                    font = "<i>"
            ElseIf InStr(.body.getElementsByTagName("h5")(h5Counter).innerHTML, "<b>") Or InStr(.body.getElementsByTagName("H5")(h5Counter).innerHTML, "<B>") Then
                    font = "<b>"
            Else: font = "False"
            End If
            Call setFont(docActiv, font, docActiv.Paragraphs.Count - 1)
            Call Headings(docActiv, "<h5>", docActiv.Paragraphs.Count - 1)
            If InStr(.body.getElementsByTagName("h5")(h5Counter).innerHTML, "<font") Or InStr(.body.getElementsByTagName("h5")(h5Counter).innerHTML, "<FONT") Then
                Call setColor(docActiv, .body.getElementsByTagName("h5")(h5Counter).getElementsByTagName("FONT")(0).getAttribute("color"), docActiv.Paragraphs.Count - 1)
            End If
            If InStr(LCase(.body.getElementsByTagName("h5")(h5Counter).innerHTML), "<a") Then
                link = .body.getElementsByTagName("h5")(h5Counter).getElementsByTagName("a")(0).href
                docActiv.Hyperlinks.Add Anchor:=docActiv.Paragraphs(docActiv.Paragraphs.Count - 1).Range, Address:=link
            End If
            paragraphNumber = paragraphNumber + 1
            docActiv.Paragraphs.Add
            ActiveDocument.Range(Start:=docActiv.Paragraphs(docActiv.Paragraphs.Count).Range.Start).Select
                Selection.ClearFormatting
            h5Counter = h5Counter + 1
        
        ElseIf InStr(WrdArray(i), "<h6") Or InStr(WrdArray(i), "<H6") Then
            Call printH6(Text, h6Counter, docActiv)
            If InStr(.body.getElementsByTagName("h6")(h6Counter).innerHTML, "<b><i>") Or InStr(.body.getElementsByTagName("H6")(h6Counter).innerHTML, "<B><I>") Then
                font = "<b><i>"
            ElseIf InStr(.body.getElementsByTagName("h6")(h6Counter).innerHTML, "<i><b>") Or InStr(.body.getElementsByTagName("H6")(h6Counter).innerHTML, "<I><B>") Then
                    font = "<b><i>"
            ElseIf InStr(.body.getElementsByTagName("h6")(h6Counter).innerHTML, "<i>") Or InStr(.body.getElementsByTagName("H6")(h6Counter).innerHTML, "<I>") Then
                    font = "<i>"
            ElseIf InStr(.body.getElementsByTagName("h6")(h6Counter).innerHTML, "<b>") Or InStr(.body.getElementsByTagName("H6")(h6Counter).innerHTML, "<B>") Then
                    font = "<b>"
            Else: font = "False"
            End If
            Call setFont(docActiv, font, docActiv.Paragraphs.Count - 1)
            Call Headings(docActiv, "<h6>", docActiv.Paragraphs.Count - 1)
            If InStr(.body.getElementsByTagName("h6")(h6Counter).innerHTML, "<font") Or InStr(.body.getElementsByTagName("h6")(h6Counter).innerHTML, "<FONT") Then
                Call setColor(docActiv, .body.getElementsByTagName("h6")(h6Counter).getElementsByTagName("FONT")(0).getAttribute("color"), docActiv.Paragraphs.Count - 1)
            End If
            If InStr(LCase(.body.getElementsByTagName("h6")(h6Counter).innerHTML), "<a") Then
                link = .body.getElementsByTagName("h6")(h6Counter).getElementsByTagName("a")(0).href
                docActiv.Hyperlinks.Add Anchor:=docActiv.Paragraphs(docActiv.Paragraphs.Count - 1).Range, Address:=link
            End If
            paragraphNumber = paragraphNumber + 1
            docActiv.Paragraphs.Add
            ActiveDocument.Range(Start:=docActiv.Paragraphs(docActiv.Paragraphs.Count).Range.Start).Select
                Selection.ClearFormatting
            h6Counter = h6Counter + 1
        
        ElseIf InStr(WrdArray(i), "<table") Or InStr(WrdArray(i), "<TABLE") Then
            Call printTable(docActiv, tableCounter, Text, docActiv.Paragraphs.Count)
            paragraphNumber = paragraphNumber + 1
            tableCounter = tableCounter + 1
        
        ElseIf InStr(WrdArray(i), "<ul") Or InStr(WrdArray(i), "<UL") Then
            Call printUL(docActiv, Text, ulCounter, docActiv.Paragraphs.Count)
            ulCounter = ulCounter + 1
            docActiv.Paragraphs.Add
            ActiveDocument.Range(Start:=docActiv.Paragraphs(docActiv.Paragraphs.Count).Range.Start).Select
                Selection.ClearFormatting
        
        ElseIf InStr(WrdArray(i), "<ol") Or InStr(WrdArray(i), "<OL") Then
            Call printOL(docActiv, Text, olCounter, docActiv.Paragraphs.Count)
            olCounter = olCounter + 1
            docActiv.Paragraphs.Add
            ActiveDocument.Range(Start:=docActiv.Paragraphs(docActiv.Paragraphs.Count).Range.Start).Select
                Selection.ClearFormatting
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
    Dim iLIordonata As Integer, paragrafeLI As Object, link As String
    Dim font As String
        .Open
        .write Text
        .Close
        Set paragrafeLI = .body.getElementsByTagName("ol")(i).getElementsByTagName("li")
            iLIordonata = 1
            j = 0
            For Each paragrafLI In paragrafeLI
                j = j + 1
                If InStr(paragrafLI.innerHTML, "<b><i>") Or InStr(paragrafLI.innerHTML, "<B><I>") Then
                font = "<b><i>"
                ElseIf InStr(paragrafLI.innerHTML, "<i><b>") Or InStr(paragrafLI.innerHTML, "<I><B>") Then
                    font = "<b><i>"
                ElseIf InStr(paragrafLI.innerHTML, "<i>") Or InStr(paragrafLI.innerHTML, "<I>") Then
                    font = "<i>"
                ElseIf InStr(paragrafLI.innerHTML, "<b>") Or InStr(paragrafLI.innerHTML, "<B>") Then
                    font = "<b>"
                Else: font = "False"
                End If
                
                docActiv.Content.InsertAfter j & "." & paragrafLI.innerText
                If InStr(LCase(paragrafLI.innerHTML), "<a") Then
                    link = paragrafLI.getElementsByTagName("a")(0).href
                    docActiv.Hyperlinks.Add Anchor:=docActiv.Paragraphs(docActiv.Paragraphs.Count).Range, Address:=link
                End If
                Call setFont(docActiv, font, docActiv.Paragraphs.Count - 1)
                docActiv.Content.InsertAfter vbCr
                pnumber = pnumber + 1
            Next
End With

End Function
Public Function printUL(ByRef docActiv As Document, Text As String, i As Integer, ByRef pnumber As Integer)
    Dim font As String, j As Integer
    With CreateObject("htmlfile") ' aici parsam bodyul
        .Open
        .write Text
        .Close
        j = 0
        Set paragrafeUL = .body.getElementsByTagName("ul")
            Set paragrafeLI = .body.getElementsByTagName("ul")(i).getElementsByTagName("li")
            For Each paragrafLI In paragrafeLI
            j = j + 1
            If InStr(paragrafLI.innerHTML, "<b><i>") Or InStr(paragrafLI.innerHTML, "<B><I>") Then
                font = "<b><i>"
            ElseIf InStr(paragrafLI.innerHTML, "<i><b>") Or InStr(paragrafLI.innerHTML, "<I><B>") Then
                    font = "<b><i>"
            ElseIf InStr(paragrafLI.innerHTML, "<i>") Or InStr(paragrafLI.innerHTML, "<I>") Then
                    font = "<i>"
            ElseIf InStr(paragrafLI.innerHTML, "<b>") Or InStr(paragrafLI.innerHTML, "<B>") Then
                    font = "<b>"
            Else: font = "False"
            End If
            
                docActiv.Content.InsertAfter Chr(183) & "  " & paragrafLI.innerText
                If InStr(LCase(paragrafLI.innerHTML), "<a") Then
                    link = paragrafLI.getElementsByTagName("a")(0).href
                    docActiv.Hyperlinks.Add Anchor:=docActiv.Paragraphs(docActiv.Paragraphs.Count).Range, Address:=link
                End If
                Call setFont(docActiv, font, docActiv.Paragraphs.Count)
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
        
        docActivPnr = docActiv.Paragraphs.Count
        'docActiv.Paragraphs.Add
        paragraphNumber = paragraphNumber + 1
        'resetam formatarile anterioare
        ActiveDocument.Range(Start:=docActiv.Paragraphs(docActivPnr).Range.Start).Select
            Selection.ClearFormatting
            
        'cream un tabel
        
        Set tbl = docActiv.Tables.Add(docActiv.Paragraphs(docActivPnr).Range, _
                                      NumRows:=nRow, _
                                      NumColumns:=nCol / nRow)
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
        docActiv.Paragraphs.Add
        
End Function


