Attribute VB_Name = "PortfelDywidendy"
Option Explicit

' ---------------------------------------------------------------------------
' ARKUSZ "DYWIDENDY" (+ dodatkowo szacunek podatku Belki)
' èrÛd≥o: wpisy typu Dywidenda w arkuszu Transakcje (lista w Dane!AP2#:
' data, spÛ≥ka, kwota, podatek, netto). Kwota z "Nowe transakcje" to netto.
' ---------------------------------------------------------------------------
Public Sub GoToDividends()
    ThisWorkbook.Worksheets(DIVIDEND_SHEET).Activate
End Sub

Public Sub BuildDividendsSheet(Optional ByVal fromCode As Boolean = False)
    Dim ws As Worksheet, wsD As Worksheet
    Dim keepStarts As Object
    Dim co As ChartObject
    Dim s As Series
    Dim keepYear As Variant, keepRates As Variant
    Dim n As Long, i As Long
    Dim compRow As Long, histRow As Long
    Dim labels As Variant

    Set wsD = ThisWorkbook.Worksheets(DATA_SHEET)
    Application.Calculate
    n = 1
    If IsNumeric(wsD.Range("AB15").Value) Then n = CLng(wsD.Range("AB15").Value)
    If n < 1 Then n = 1

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DIVIDEND_SHEET)
    On Error GoTo 0
    keepYear = Year(Date)
    keepRates = Array(0.19, 0.15, 0, 0)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ANALYSIS_SHEET))
        ws.Name = DIVIDEND_SHEET
    Else
        ' Zachowaj rok i stawki wpisane przez uøytkownika.
        If IsNumeric(ws.Range("M6").Value) And Not IsEmpty(ws.Range("M6").Value) Then keepYear = ws.Range("M6").Value
        For i = 0 To 3
            If IsNumeric(ws.Cells(15 + i, "M").Value) And Not IsEmpty(ws.Cells(15 + i, "M").Value) Then keepRates(i) = ws.Cells(15 + i, "M").Value
        Next i
        Set keepStarts = ReadStartPanel(ws)
        Do While ws.ChartObjects.Count > 0
            ws.ChartObjects(1).Delete
        Loop
        Do While ws.Shapes.Count > 0
            ws.Shapes(1).Delete
        Loop
        On Error Resume Next
        ws.Cells.Validation.Delete
        ws.Cells.FormatConditions.Delete
        On Error GoTo 0
        ws.Cells.UnMerge
        ws.Cells.Clear
    End If
    If keepStarts Is Nothing Then Set keepStarts = CreateObject("Scripting.Dictionary")

    ThemeSheet ws
    ShowSheetLikeApp ws, True
    ws.Cells.RowHeight = 18
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 28
    ws.Columns("C:J").ColumnWidth = 13
    ws.Columns("K").ColumnWidth = 3
    ws.Columns("L").ColumnWidth = 34
    ws.Columns("M").ColumnWidth = 14
    ws.Columns("N").ColumnWidth = 3

    With ws
        .Range("B1").Value = "Dywidendy"
        .Range("B1").Font.Size = 20
        .Rows(1).RowHeight = 36
        .Range("B1").VerticalAlignment = xlCenter
        .Range("B2").Value = "Kwoty netto w PLN. Liczone automatycznie raz dziennie z danych Yahoo (dywidenda na akcjÍ " & ChrW(215) & _
                             " liczba akcji w dniu ustalenia prawa, po podatku u ürÛd≥a) - tylko do statystyk, gotÛwka siÍ nie zmienia."
        .Range("B2").Font.Color = C_MUTED

        ' Kafelki
        labels = Array("Prognoza na 12 mies. (netto)", "Stopa dywidendy portfela", "Stopa od kosztu", _
                       "Otrzymane w 12 mies.", "Otrzymane w tym roku", "Otrzymane ≥πcznie")
        For i = 0 To 5
            .Cells(4, 2 + i).Value = labels(i)
        Next i
        .Range("B5").Formula2 = "=" & DATA_SHEET & "!$AB$28"
        .Range("C5").Formula2 = "=" & DATA_SHEET & "!$AB$29"
        .Range("D5").Formula2 = "=" & DATA_SHEET & "!$AB$24"
        .Range("E5").Formula2 = "=" & DATA_SHEET & "!$AB$20"
        .Range("F5").Formula2 = "=" & DATA_SHEET & "!$AB$22"
        .Range("G5").Formula2 = "=" & DATA_SHEET & "!$AB$21"
        With .Range("B4:G4")
            .Font.Size = 9
            .Font.Color = C_MUTED
            .Interior.Color = C_PANEL
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        .Rows(4).RowHeight = 28
        With .Range("B5:G5")
            .Font.Size = 15
            .Interior.Color = C_PANEL
            .HorizontalAlignment = xlLeft
        End With
        .Range("B5:G5").Font.Color = C_FG
        .Range("B5").Font.Color = C_OK
        .Rows(5).RowHeight = 30
        .Range("B5,E5:G5").NumberFormat = "#,##0.00 ""z≥"""
        .Range("C5:D5").NumberFormat = "0.00%"
    End With

    ' Wykres miesiÍczny (24 miesiπce).
    Set co = ws.ChartObjects.Add(ws.Range("B7").Left, ws.Range("B7").Top, ws.Range("B7:J7").Width, ws.Range("B7:B20").Height)
    co.Name = "WykresDywidend"
    co.Placement = xlFreeFloating
    With co.Chart
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        Set s = .SeriesCollection.NewSeries
        s.Name = "Dywidendy netto (PLN)"
        s.XValues = wsD.Range("AY2:AY25")
        s.Values = wsD.Range("AX2:AX25")
        s.Format.Fill.ForeColor.RGB = C_OK
        s.Format.Line.Visible = msoFalse
        .ChartType = xlColumnClustered
        .HasLegend = False
        .ChartGroups(1).GapWidth = 40
        StyleChartDark co.Chart, "Dywidendy otrzymane w ostatnich 24 miesiπcach (netto, PLN)"
        On Error Resume Next
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        On Error GoTo 0
    End With

    ' Tabela spÛ≥ek.
    compRow = 23
    ws.Cells(compRow - 1, "B").Value = "SpÛ≥ki"
    ws.Cells(compRow - 1, "B").Font.Size = 12
    ws.Cells(compRow - 1, "B").Font.Bold = True
    ws.Range(ws.Cells(compRow, "B"), ws.Cells(compRow, "J")).Value = Array("SpÛ≥ka / ETF", "Dywidenda", "Prognoza 12 mies. (PLN)", _
        "Stopa dywidendy", "Stopa od kosztu", "Otrzymane w 12 mies.", "Otrzymane ≥πcznie", "Ostatnia dywidenda (ex)", "Liczba wyp≥at")
    StyleTableHeader ws.Range(ws.Cells(compRow, "B"), ws.Cells(compRow, "J"))
    ws.Cells(compRow + 1, "B").Formula2 = "=LET(nazwy,Pozycje!$B$5:$B$500,statusy,Pozycje!$C$5:$C$500,koszty,Pozycje!$F$5:$F$500," & _
        "fNazwy," & DIV_FORECAST_TABLE & "[SpÛ≥ka],fPrognozy," & DIV_FORECAST_TABLE & "[Prognoza 12 mies. netto (PLN)]," & _
        "fStopy," & DIV_FORECAST_TABLE & "[Stopa dywidendy],fOstatnie," & DIV_FORECAST_TABLE & "[Ostatnia dywidenda]," & _
        "prognoza,IFERROR(XLOOKUP(nazwy,fNazwy,fPrognozy,0),0),stopa,IFERROR(XLOOKUP(nazwy,fNazwy,fStopy,0),0)," & _
        "ostEx,IFERROR(XLOOKUP(nazwy,fNazwy,fOstatnie,0),0),odKosztu,IFERROR(prognoza/koszty,0)," & _
        "dDaty,INDEX(" & DATA_SHEET & "!$AP$2#,0,1),dNazwy,INDEX(" & DATA_SHEET & "!$AP$2#,0,2),dNetto,INDEX(" & DATA_SHEET & "!$AP$2#,0,5)," & _
        "lacznie,IFERROR(SUMIFS(dNetto,dNazwy,nazwy),0)," & _
        "ost12,IFERROR(SUMIFS(dNetto,dNazwy,nazwy,dDaty,"">""&EDATE(TODAY(),-12)),0)," & _
        "ilosc,IFERROR(COUNTIFS(dNazwy,nazwy),0),opis,IF(statusy=""Tak"",""Wyp≥aca"",statusy)," & _
        "maska,(nazwy<>"""")*(((prognoza>0)+(lacznie>0)+(statusy=""Tak""))>0)," & _
        "IF(SUM(maska)=0,""Brak spÛ≥ek wyp≥acajπcych dywidendy""," & _
        "SORTBY(FILTER(HSTACK(nazwy,opis,prognoza,stopa,odKosztu,ost12,lacznie,ostEx,ilosc),maska),FILTER(prognoza,maska),-1,FILTER(lacznie,maska),-1)))"
    StyleTableBody ws.Range(ws.Cells(compRow + 1, "B"), ws.Cells(compRow + n + 1, "J"))
    ws.Range(ws.Cells(compRow + 1, "B"), ws.Cells(compRow + n + 1, "C")).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(compRow + 1, "C"), ws.Cells(compRow + n + 1, "C")).Font.Color = C_MUTED
    ws.Range(ws.Cells(compRow + 1, "D"), ws.Cells(compRow + n + 1, "D")).NumberFormat = "#,##0.00;-#,##0.00;""-"""
    ws.Range(ws.Cells(compRow + 1, "E"), ws.Cells(compRow + n + 1, "F")).NumberFormat = "0.00%;;""-"""
    ws.Range(ws.Cells(compRow + 1, "G"), ws.Cells(compRow + n + 1, "H")).NumberFormat = "#,##0.00;-#,##0.00;""-"""
    ws.Range(ws.Cells(compRow + 1, "I"), ws.Cells(compRow + n + 1, "I")).NumberFormat = "dd.mm.yyyy;;""-"""
    ws.Range(ws.Cells(compRow + 1, "J"), ws.Cells(compRow + n + 1, "J")).NumberFormat = "0;;""-"""
    ws.Range(ws.Cells(compRow + 1, "D"), ws.Cells(compRow + n + 1, "D")).Font.Color = C_OK

    ' Historia wyp≥at (najnowsze na gÛrze).
    histRow = compRow + n + 4
    ws.Cells(histRow - 1, "B").Value = "Historia wyp≥at"
    ws.Cells(histRow - 1, "B").Font.Size = 12
    ws.Cells(histRow - 1, "B").Font.Bold = True
    ws.Range(ws.Cells(histRow, "B"), ws.Cells(histRow, "G")).Value = Array("Data", "SpÛ≥ka / ETF", "Brutto", "Podatek u ürÛd≥a", "Netto", "èrÛd≥o")
    StyleTableHeader ws.Range(ws.Cells(histRow, "B"), ws.Cells(histRow, "G"))
    ws.Cells(histRow + 1, "B").Formula2 = "=IFERROR(IF(INDEX(" & DATA_SHEET & "!$AP$2#,1,1)="""",""Brak dywidend - pojawiπ siÍ po pierwszym pobraniu z Yahoo (przycisk POBIERZ DYWIDENDY)."", " & _
        DATA_SHEET & "!$AP$2#),""Brak zapisanych dywidend."")"
    StyleTableBody ws.Range(ws.Cells(histRow + 1, "B"), ws.Cells(histRow + 200, "G"))
    ws.Range(ws.Cells(histRow + 1, "G"), ws.Cells(histRow + 200, "G")).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(histRow + 1, "G"), ws.Cells(histRow + 200, "G")).Font.Color = C_MUTED
    ws.Range(ws.Cells(histRow + 1, "B"), ws.Cells(histRow + 200, "C")).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(histRow + 1, "B"), ws.Cells(histRow + 200, "B")).NumberFormat = "dd.mm.yyyy"
    ws.Range(ws.Cells(histRow + 1, "D"), ws.Cells(histRow + 200, "F")).NumberFormat = "#,##0.00"

    BuildBelkaPanel ws, keepYear, keepRates
    BuildStartPanel ws, keepStarts
End Sub

Public Function ReadStartPanel(ByVal ws As Worksheet) As Object
    Dim d As Object
    Dim r As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    For r = START_PANEL_ROW + 1 To START_PANEL_ROW + 60
        If Len(Trim$(CStr(ws.Cells(r, "L").Value))) = 0 Then Exit For
        If IsDate(ws.Cells(r, "M").Value) Then d(Trim$(CStr(ws.Cells(r, "L").Value))) = CDate(ws.Cells(r, "M").Value)
    Next r
    Set ReadStartPanel = d
End Function

' Panel "od kiedy masz akcje": pozycje z salda poczπtkowego, dla ktÛrych nie ma
' jeszcze øadnej zapisanej dywidendy (np. spoza Trading 212). Na tej podstawie
' automat liczy dywidendy wstecz.
Public Sub BuildStartPanel(ByVal ws As Worksheet, ByVal keepStarts As Object)
    Dim initNames As Variant
    Dim loH As ListObject
    Dim lr As ListRow
    Dim haveDiv As Object
    Dim i As Long, r As Long
    Dim nm As String

    Set haveDiv = CreateObject("Scripting.Dictionary")
    haveDiv.CompareMode = vbTextCompare
    Set loH = FindTable(DIV_HISTORY_TABLE)
    If Not loH Is Nothing Then
        For Each lr In loH.ListRows
            If CStr(lr.Range.Cells(1, 4).Value) <> DIV_AUTO_SOURCE Then haveDiv(Trim$(CStr(lr.Range.Cells(1, 2).Value))) = True
        Next lr
    End If

    With ws
        .Range("L" & START_PANEL_ROW - 1).Value = "Od kiedy masz akcje (opcjonalnie)"
        .Range("L" & START_PANEL_ROW - 1).Font.Bold = True
        .Range("L" & START_PANEL_ROW).Value = "SpÛ≥ka / data zakupu (pierwszego)"
        .Range("L" & START_PANEL_ROW).Font.Color = C_MUTED
    End With
    initNames = ThisWorkbook.Worksheets("Transakcje").Range("B7:B29").Value
    r = START_PANEL_ROW
    For i = 1 To UBound(initNames, 1)
        nm = Trim$(CStr(initNames(i, 1)))
        If Len(nm) > 0 And Not haveDiv.Exists(nm) Then
            r = r + 1
            ws.Cells(r, "L").Value = nm
            If keepStarts.Exists(nm) Then ws.Cells(r, "M").Value = keepStarts(nm)
        End If
    Next i
    If r > START_PANEL_ROW Then
        With ws.Range(ws.Cells(START_PANEL_ROW + 1, "M"), ws.Cells(r, "M"))
            .NumberFormat = "dd.mm.yyyy"
            .Interior.Color = C_INPUT
            .Borders.Color = C_BORDER
            .HorizontalAlignment = xlRight
        End With
    End If
    ws.Cells(r + 2, "L").Value = "SpÛ≥ki bez zaimportowanych dywidend. Data pierwszego zakupu pozwala automatowi policzyÊ dywidendy wstecz; " & _
                                 "po imporcie z brokerÛw zwykle nie jest potrzebna."
    ws.Cells(r + 2, "L").Font.Size = 9
    ws.Cells(r + 2, "L").Font.Color = C_MUTED
    ws.Range(ws.Cells(START_PANEL_ROW - 1, "L"), ws.Cells(r + 2, "M")).Interior.Color = C_PANEL
    ws.Range(ws.Cells(START_PANEL_ROW + 1, "M"), ws.Cells(Application.Max(r, START_PANEL_ROW + 1), "M")).Interior.Color = C_INPUT
End Sub

' Dodatek: szacunek podatku Belki (19%) za wybrany rok.
Public Sub BuildBelkaPanel(ByVal ws As Worksheet, ByVal taxYear As Variant, ByVal rates As Variant)
    Dim i As Long
    With ws
        .Range("L4").Value = "Podatek Belki - szacunek (dodatek)"
        .Range("L4").Font.Bold = True
        .Range("L4").Font.Size = 11
        .Range("L6").Value = "Rok podatkowy"
        .Range("M6").Value = taxYear
        .Range("M6").Interior.Color = C_INPUT
        .Range("M6").Borders.Color = C_BORDER
        .Range("M6").NumberFormat = "0"
        .Range("M6").HorizontalAlignment = xlRight

        .Range("L8").Value = "Zrealizowany wynik ze sprzedaøy"
        .Range("M8").Formula2 = "=IFERROR(LET(rokP,$M$6,daty,Transakcje!$A$30:$A$500,typy,Transakcje!$C$30:$C$500,wyniki,Transakcje!$I$30:$I$500," & _
            "SUM(IF((typy=""Sprzedaø"")*ISNUMBER(daty)*ISNUMBER(wyniki),IF(YEAR(daty)=rokP,wyniki,0),0))),0)" & _
            "+IFERROR(SUMPRODUCT((YEAR(" & REALIZED_TABLE & "[Data])=$M$6)*" & REALIZED_TABLE & "[Wynik (PLN)]),0)"
        .Range("L9").Value = "Podatek od sprzedaøy (19%)"
        .Range("M9").Formula2 = "=MAX(0,M8)*0.19"
        .Range("L10").Value = "Dop≥ata od dywidend zagranicznych"
        .Range("M10").Formula2 = "=IFERROR(LET(rokP,$M$6,dDaty,INDEX(" & DATA_SHEET & "!$AP$2#,0,1),dNazwy,INDEX(" & DATA_SHEET & "!$AP$2#,0,2)," & _
            "dKwoty,INDEX(" & DATA_SHEET & "!$AP$2#,0,3),dPodatki,INDEX(" & DATA_SHEET & "!$AP$2#,0,4),dNetto,INDEX(" & DATA_SHEET & "!$AP$2#,0,5)," & _
            "dZrodla,INDEX(" & DATA_SHEET & "!$AP$2#,0,6)," & _
            "waluty,XLOOKUP(dNazwy,Notowania!$A$7:$A$502,Notowania!$F$7:$F$502,""PLN"")," & _
            "stawki,XLOOKUP(waluty,$L$15:$L$18,$M$15:$M$18,0.19)," & _
            "znane,(dPodatki>0)+(dZrodla=""" & DIV_T212_SOURCE & """)+(dZrodla=""" & DIV_AUTO_SOURCE & """)," & _
            "brutto,IF(znane,dKwoty,dNetto/(1-stawki)),pobrany,IF(znane,dPodatki,brutto*stawki)," & _
            "zwolnione,ISNUMBER(SEARCH(""IKE"",dZrodla))," & _
            "doplata,IF((YEAR(dDaty)=rokP)*NOT(zwolnione),IF(0.19*brutto-pobrany>0,0.19*brutto-pobrany,0),0),SUM(doplata)),0)"
        .Range("L11").Value = "Razem do zap≥aty (szacunek)"
        .Range("M11").Formula2 = "=M9+M10"
        .Range("L11:M11").Font.Bold = True
        .Range("M8:M11").NumberFormat = "#,##0.00 ""z≥"";-#,##0.00 ""z≥"""
        .Range("L8:M11").Borders(xlInsideHorizontal).Color = C_ROWLINE
        .Range("L11:M11").Borders(xlEdgeTop).Color = C_BORDER

        .Range("L13").Value = "Podatek pobrany u ürÛd≥a (za≥oøenia, moøesz zmieniÊ)"
        .Range("L13").Font.Color = C_MUTED
        .Range("L14:M14").Value = Array("Waluta notowania", "Stawka")
        .Range("L14:M14").Font.Color = C_MUTED
        .Range("L15:L18").Value = Application.WorksheetFunction.Transpose(Array("PLN", "USD", "GBP", "EUR"))
        For i = 0 To 3
            .Cells(15 + i, "M").Value = rates(i)
        Next i
        .Range("M15:M18").NumberFormat = "0%"
        .Range("M15:M18").Interior.Color = C_INPUT
        .Range("M15:M18").Borders.Color = C_BORDER

        .Range("L20").Value = "To tylko szacunek do planowania. Urzπd liczy koszt metodπ FIFO i po kursach NBP z dnia " & _
            "poprzedzajπcego transakcjÍ, a tu uøywana jest úrednia cena i kwoty w PLN z Twoich wpisÛw. " & _
            "Dywidendy z GPW sπ juø opodatkowane u ürÛd≥a, a IKE i IKZE sπ zwolnione z podatku Belki. Do PIT-38 uøyj PIT-8C od brokera."
        .Range("L20:M27").Merge
        .Range("L20").WrapText = True
        .Range("L20").VerticalAlignment = xlTop
        .Range("L20").Font.Size = 9
        .Range("L20").Font.Color = C_MUTED
        .Range("L4:M27").Interior.Color = C_PANEL
    End With
End Sub

Public Sub StyleTableHeader(ByVal rng As Range)
    With rng
        .Font.Color = C_MUTED
        .Font.Size = 9
        .Interior.Color = C_PANEL
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlRight
        .Borders(xlEdgeBottom).Color = C_BORDER
        .RowHeight = 30
    End With
    rng.Cells(1, 1).HorizontalAlignment = xlLeft
End Sub

Public Sub StyleTableBody(ByVal rng As Range)
    With rng
        .Interior.Color = C_PANEL
        .Borders(xlInsideHorizontal).Color = C_ROWLINE
        .Borders(xlEdgeBottom).Color = C_ROWLINE
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .RowHeight = 20
    End With
End Sub

' ---------------------------------------------------------------------------
' DYWIDENDY AUTOMATYCZNIE
' Raz dziennie (przy aktualizacji kursÛw) makro pobiera z Yahoo dywidendy na
' akcjÍ kaødego instrumentu, liczy ile jednostek by≥o w posiadaniu w dniu
' ustalenia prawa (ex-date), przelicza na PLN i odejmuje podatek u ürÛd≥a.
' Wyniki trafiajπ do tabeli DywidendyHistoria (ürÛd≥o "automatycznie") -
' tylko do statystyk, gotÛwka siÍ nie zmienia. Wyp≥aty wpisane rÍcznie
' (tabela i rejestr transakcji) majπ pierwszeÒstwo.
' ---------------------------------------------------------------------------
Public Sub SyncDividendsNow()
    If SyncDividends(True) Then
        MsgBox "Dywidendy zaktualizowane.", vbInformation, "Dywidendy"
    Else
        MsgBox "Nie uda≥o siÍ pobraÊ dywidend: " & YahooQuoteLastError, vbExclamation, "Dywidendy"
    End If
End Sub

Public Function SyncDividends(ByVal force As Boolean) As Boolean
    Dim wsQ As Worksheet, wsD As Worksheet, wsT As Worksheet, wsP As Worksheet
    Dim loH As ListObject, loFc As ListObject
    Dim ledger As Variant, initRows As Variant
    Dim r As Long, lastQ As Long, k As Long, i As Long, posRow As Long
    Dim sym As String, nm As String, cur As String
    Dim amounts() As Double, exDates() As Date, cnt As Long
    Dim fx As Double, w As Double, units As Double
    Dim threshold As Date, trackStart As Date
    Dim dps12 As Double, lastEx As Date, priceCur As Double, unitsNow As Double, costNow As Double
    Dim autoRows As Collection, fRows As Collection
    Dim item As Variant, okAny As Boolean
    Dim lastSync As Variant

    On Error GoTo Failed
    Set wsD = ThisWorkbook.Worksheets(DATA_SHEET)
    Set wsQ = ThisWorkbook.Worksheets("Notowania")
    Set wsT = ThisWorkbook.Worksheets("Transakcje")
    Set wsP = ThisWorkbook.Worksheets("Pozycje")
    Set loH = FindTable(DIV_HISTORY_TABLE)
    Set loFc = FindTable(DIV_FORECAST_TABLE)
    If loH Is Nothing Or loFc Is Nothing Then Exit Function

    lastSync = wsD.Range("AB30").Value
    If Not force And IsDate(lastSync) Then
        If Int(CDbl(CDate(lastSync))) = CLng(Date) Then
            SyncDividends = True
            Exit Function
        End If
    End If
    If IsDate(wsD.Range("AB31").Value) Then
        trackStart = CDate(wsD.Range("AB31").Value)
    Else
        trackStart = Date
        wsD.Range("AB31").Value = Date
    End If

    ledger = wsT.Range("A30:L500").Value
    initRows = wsT.Range("B7:D29").Value
    Set autoRows = New Collection
    Set fRows = New Collection

    lastQ = wsQ.Cells(wsQ.Rows.Count, "S").End(xlUp).Row
    For r = 7 To lastQ
        sym = Trim$(CStr(wsQ.Cells(r, "S").Value))
        nm = Trim$(CStr(wsQ.Cells(r, "A").Value))
        If Len(sym) > 0 And Len(nm) > 0 Then
            Application.StatusBar = "Pobieranie dywidend: " & sym
            If GetYahooDividends(sym, amounts, exDates, cnt, cur) Then
                okAny = True
                fx = NumVal(wsQ.Cells(r, "G").Value)
                If fx <= 0 Then fx = 1
                w = WithholdingRate(CStr(wsQ.Cells(r, "F").Value))

                ' Od kiedy liczyÊ: pÛüniejsza z dat - poczπtek posiadania / ostatnia wpisana wyp≥ata.
                threshold = StartDateFor(nm, trackStart)
                If LastManualDividend(loH, ledger, nm) > threshold Then threshold = LastManualDividend(loH, ledger, nm)

                dps12 = 0
                lastEx = 0
                For k = 1 To cnt
                    If exDates(k) > Date - 365 And exDates(k) <= Date Then dps12 = dps12 + amounts(k)
                    If exDates(k) <= Date And exDates(k) > lastEx Then lastEx = exDates(k)
                    If exDates(k) > threshold And exDates(k) <= Date Then
                        units = UnitsHeldAt(nm, exDates(k), initRows, ledger, StartDateFor(nm, trackStart))
                        If units > 0 Then
                            autoRows.Add Array(exDates(k), nm, Round(amounts(k) * units * fx * (1 - w), 2), DIV_AUTO_SOURCE, _
                                               Round(amounts(k) * units * fx * w, 2), "")
                        End If
                    End If
                Next k

                posRow = FindPositionRow(wsP, nm)
                unitsNow = 0
                costNow = 0
                If posRow > 0 Then
                    unitsNow = NumVal(wsP.Cells(posRow, "D").Value)
                    costNow = NumVal(wsP.Cells(posRow, "F").Value)
                    If dps12 > 0 Then
                        If Len(Trim$(CStr(wsP.Cells(posRow, "C").Value))) = 0 Or CStr(wsP.Cells(posRow, "C").Value) = "Nieustalone" Then
                            wsP.Cells(posRow, "C").Value = "Tak"
                        End If
                    End If
                End If
                priceCur = 0
                If IsNumeric(wsQ.Cells(r, "E").Value) And Not IsEmpty(wsQ.Cells(r, "E").Value) Then
                    priceCur = CDbl(wsQ.Cells(r, "E").Value) / fx
                End If
                fRows.Add Array(nm, dps12, cur, unitsNow, Round(dps12 * unitsNow * fx * (1 - w), 2), _
                                IIf(priceCur > 0, dps12 / priceCur, 0), costNow, IIf(lastEx > 0, lastEx, Empty))
            End If
            Sleep 150
        End If
    Next r
    Application.StatusBar = False
    If Not okAny Then Exit Function

    ' Zastπp wczeúniejsze wyliczenia automatyczne.
    For i = loH.ListRows.Count To 1 Step -1
        If CStr(loH.ListRows(i).Range.Cells(1, 4).Value) = DIV_AUTO_SOURCE Or _
           Len(Trim$(CStr(loH.ListRows(i).Range.Cells(1, 2).Value))) = 0 Then loH.ListRows(i).Delete
    Next i
    For Each item In autoRows
        loH.ListRows.Add.Range.Value = item
    Next item
    If Not loH.DataBodyRange Is Nothing Then
        loH.ListColumns(1).DataBodyRange.NumberFormat = "dd.mm.yyyy"
        loH.ListColumns(3).DataBodyRange.NumberFormat = "#,##0.00"
        With loH.Sort
            .SortFields.Clear
            .SortFields.Add Key:=loH.ListColumns(1).DataBodyRange, Order:=xlDescending
            .Header = xlYes
            .Apply
        End With
    End If

    If Not loFc.DataBodyRange Is Nothing Then loFc.DataBodyRange.Delete
    For Each item In fRows
        loFc.ListRows.Add.Range.Value = item
    Next item
    If Not loFc.DataBodyRange Is Nothing Then
        loFc.ListColumns(2).DataBodyRange.NumberFormat = "0.0000"
        loFc.ListColumns(5).DataBodyRange.NumberFormat = "#,##0.00"
        loFc.ListColumns(6).DataBodyRange.NumberFormat = "0.00%"
        loFc.ListColumns(7).DataBodyRange.NumberFormat = "#,##0.00"
        loFc.ListColumns(8).DataBodyRange.NumberFormat = "dd.mm.yyyy"
    End If

    wsD.Range("AB30").Value = Now
    wsD.Range("AB30").NumberFormat = "dd.mm.yyyy hh:mm"
    SyncDividends = True
    Exit Function

Failed:
    Application.StatusBar = False
    If Len(YahooQuoteLastError) = 0 Then YahooQuoteLastError = Err.Description
    SyncDividends = False
End Function

' Dywidendy na akcjÍ z Yahoo (ostatnie 5 lat). Kwoty w walucie notowania,
' pensy przeliczone na funty.
Public Function GetYahooDividends(ByVal symbol As String, ByRef amounts() As Double, ByRef exDates() As Date, _
                                  ByRef cnt As Long, ByRef quoteCurrency As String) As Boolean
    Dim http As Object, re As Object, ms As Object
    Dim body As String
    Dim k As Long
    Dim divisor As Double

    cnt = 0
    ReDim amounts(1 To 1)
    ReDim exDates(1 To 1)
    On Error GoTo Fail
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", YAHOO_QUOTE_ROOT & Replace(symbol, "=", "%3D") & "?range=5y&interval=1mo&events=div", False
    http.SetRequestHeader "User-Agent", "Mozilla/5.0"
    http.SetRequestHeader "Accept", "application/json"
    http.SetTimeouts 5000, 5000, 10000, 10000
    http.Send
    If http.Status <> 200 Then
        YahooQuoteLastError = "HTTP " & CStr(http.Status) & " dla " & symbol
        Exit Function
    End If
    body = CStr(http.ResponseText)
    quoteCurrency = YahooQuoteJsonString(body, "currency")
    divisor = 1
    If quoteCurrency = "GBp" Or quoteCurrency = "GBX" Or quoteCurrency = "GBPX" Then
        divisor = 100
        quoteCurrency = "GBP"
    End If

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = """amount""\s*:\s*([0-9.eE+-]+)\s*,\s*""date""\s*:\s*([0-9]+)"
    Set ms = re.Execute(body)
    If ms.Count > 0 Then
        ReDim amounts(1 To ms.Count)
        ReDim exDates(1 To ms.Count)
        For k = 0 To ms.Count - 1
            cnt = cnt + 1
            amounts(cnt) = Val(ms(k).SubMatches(0)) / divisor
            exDates(cnt) = Int(DateAdd("s", CDbl(ms(k).SubMatches(1)), DateSerial(1970, 1, 1)))
        Next k
    End If
    GetYahooDividends = True
    Exit Function
Fail:
    YahooQuoteLastError = "B≥πd pobierania dywidend " & symbol & ": " & Err.Description
    GetYahooDividends = False
End Function

Public Function FindTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = tableName Then
                Set FindTable = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Public Function FindPositionRow(ByVal wsP As Worksheet, ByVal nm As String) As Long
    Dim r As Long, lastRow As Long
    lastRow = wsP.Cells(wsP.Rows.Count, "B").End(xlUp).Row
    For r = 5 To lastRow
        If StrComp(Trim$(CStr(wsP.Cells(r, "B").Value)), nm, vbTextCompare) = 0 Then
            FindPositionRow = r
            Exit Function
        End If
    Next r
End Function

' Stawka podatku pobieranego u ürÛd≥a - z tabeli za≥oøeÒ w arkuszu Dywidendy.
Public Function WithholdingRate(ByVal currencyCode As String) As Double
    Dim ws As Worksheet
    Dim i As Long
    currencyCode = UCase$(Trim$(currencyCode))
    Select Case currencyCode
        Case "PLN": WithholdingRate = 0.19
        Case "USD": WithholdingRate = 0.15
        Case Else: WithholdingRate = 0
    End Select
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DIVIDEND_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For i = 15 To 18
        If UCase$(Trim$(CStr(ws.Cells(i, "L").Value))) = currencyCode And IsNumeric(ws.Cells(i, "M").Value) _
           And Not IsEmpty(ws.Cells(i, "M").Value) Then
            WithholdingRate = CDbl(ws.Cells(i, "M").Value)
        End If
    Next i
End Function

' Data ostatniej dywidendy wpisanej rÍcznie (tabela lub rejestr transakcji).
Public Function LastManualDividend(ByVal loH As ListObject, ByVal ledger As Variant, ByVal nm As String) As Date
    Dim lr As ListRow
    Dim i As Long
    For Each lr In loH.ListRows
        If CStr(lr.Range.Cells(1, 4).Value) <> DIV_AUTO_SOURCE And _
           StrComp(Trim$(CStr(lr.Range.Cells(1, 2).Value)), nm, vbTextCompare) = 0 And IsDate(lr.Range.Cells(1, 1).Value) Then
            If CDate(lr.Range.Cells(1, 1).Value) > LastManualDividend Then LastManualDividend = CDate(lr.Range.Cells(1, 1).Value)
        End If
    Next lr
    For i = 1 To UBound(ledger, 1)
        If CStr(ledger(i, 3)) = OP_DIVIDEND And StrComp(Trim$(CStr(ledger(i, 2))), nm, vbTextCompare) = 0 And IsDate(ledger(i, 1)) Then
            If CDate(ledger(i, 1)) > LastManualDividend Then LastManualDividend = CDate(ledger(i, 1))
        End If
    Next i
End Function

' Liczba jednostek w posiadaniu w dniu ex-date (przed tym dniem).
Public Function UnitsHeldAt(ByVal nm As String, ByVal exDate As Date, ByVal initRows As Variant, _
                             ByVal ledger As Variant, ByVal startDate As Date) As Double
    Dim i As Long
    Dim u As Double
    If exDate > startDate Then
        For i = 1 To UBound(initRows, 1)
            If StrComp(Trim$(CStr(initRows(i, 1))), nm, vbTextCompare) = 0 And IsNumeric(initRows(i, 3)) Then
                u = u + CDbl(initRows(i, 3))
            End If
        Next i
    End If
    For i = 1 To UBound(ledger, 1)
        If StrComp(Trim$(CStr(ledger(i, 2))), nm, vbTextCompare) = 0 And IsDate(ledger(i, 1)) And CStr(ledger(i, 12)) = "OK" Then
            If CDate(ledger(i, 1)) < exDate And IsNumeric(ledger(i, 4)) And Not IsEmpty(ledger(i, 4)) Then
                If CStr(ledger(i, 3)) = OP_BUY Then u = u + CDbl(ledger(i, 4))
                If CStr(ledger(i, 3)) = OP_SELL Then u = u - CDbl(ledger(i, 4))
            End If
        End If
    Next i
    UnitsHeldAt = u
End Function

' Tabela w arkuszu Dane (tworzona, gdy jej nie ma).
Public Function EnsureTable(ByVal wsD As Worksheet, ByVal tableName As String, ByVal anchor As String, ByVal headers As Variant) As ListObject
    Dim lo As ListObject
    Dim n As Long
    Set lo = FindTable(tableName)
    If lo Is Nothing Then
        n = UBound(headers) - LBound(headers) + 1
        wsD.Range(anchor).Resize(1, n).Value = headers
        Set lo = wsD.ListObjects.Add(xlSrcRange, wsD.Range(anchor).Resize(2, n), , xlYes)
        lo.Name = tableName
        lo.TableStyle = "TableStyleLight1"
        If lo.ListRows.Count = 1 Then
            If Len(Trim$(CStr(lo.ListRows(1).Range.Cells(1, 1).Value))) = 0 Then lo.ListRows(1).Delete
        End If
    End If
    Set EnsureTable = lo
End Function

' Tabele dywidend i wynikÛw sprzedaøy (arkusz Dane).
Public Sub SetupDividendTables(ByVal wsD As Worksheet)
    Dim loH As ListObject

    Set loH = EnsureTable(wsD, DIV_HISTORY_TABLE, "BA1", Array("Data", "SpÛ≥ka", "Kwota netto (PLN)", "èrÛd≥o", "Podatek u ürÛd≥a (PLN)", "Klucz"))
    EnsureColumns loH, Array("Data", "SpÛ≥ka", "Kwota netto (PLN)", "èrÛd≥o", "Podatek u ürÛd≥a (PLN)", "Klucz")
    EnsureTable wsD, DIV_FORECAST_TABLE, "BI1", Array("SpÛ≥ka", "Dywidenda na akcjÍ 12 mies.", "Waluta", "Liczba jednostek", _
                "Prognoza 12 mies. netto (PLN)", "Stopa dywidendy", "Koszt pozycji (PLN)", "Ostatnia dywidenda")
    EnsureTable wsD, REALIZED_TABLE, "BS1", Array("Data", "SpÛ≥ka", "Wynik (PLN)", "Klucz")

    If Not IsDate(wsD.Range("AB31").Value) Then wsD.Range("AB31").Value = Date
    wsD.Range("AB31").NumberFormat = "dd.mm.yyyy"
    wsD.Range("AA30").Value = "Ostatnie pobranie dywidend"
    wsD.Range("AA31").Value = "Poczπtek automatycznego liczenia dywidend"
    wsD.Range("AA32").Value = "DzieÒ salda poczπtkowego (import Trading 212)"
    ThisWorkbook.Worksheets("Transakcje").Range("S6").Value = "ID importu"
End Sub

Public Sub EnsureColumns(ByVal lo As ListObject, ByVal headers As Variant)
    Dim h As Variant
    For Each h In headers
        If Not ListColumnExists(lo, CStr(h)) Then AddTableColumn lo, CStr(h)
    Next h
End Sub

' Data zakupu pozycji z salda poczπtkowego - z panelu w arkuszu Dywidendy
' (kolumny L:M od wiersza START_PANEL_ROW). Brak daty = od poczπtku automatu.
Public Function StartDateFor(ByVal nm As String, ByVal fallback As Date) As Date
    Dim ws As Worksheet
    Dim r As Long
    StartDateFor = fallback
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DIVIDEND_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For r = START_PANEL_ROW + 1 To START_PANEL_ROW + 60
        If Len(Trim$(CStr(ws.Cells(r, "L").Value))) = 0 Then Exit Function
        If StrComp(Trim$(CStr(ws.Cells(r, "L").Value)), nm, vbTextCompare) = 0 Then
            If IsDate(ws.Cells(r, "M").Value) Then StartDateFor = CDate(ws.Cells(r, "M").Value)
            Exit Function
        End If
    Next r
End Function
