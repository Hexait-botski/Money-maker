Attribute VB_Name = "PortfelWidoki"
Option Explicit

' ---------------------------------------------------------------------------
' Jednorazowa instalacja / naprawa skoroszytu. Moøna uruchamiaÊ ponownie -
' nie kasuje transakcji ani wpisanych, niezapisanych wierszy.
' ---------------------------------------------------------------------------
Public Sub InstallPortfolioTools()
    On Error GoTo Failed
    Application.ScreenUpdating = False

    SetupEntrySheet
    SetupQuotesSheet
    SetupDataSheet
    SetupAnalysisSheet
    Application.Calculate
    BuildDashboard
    BuildDividendsSheet
    Application.StatusBar = "Pobieranie dywidend z Yahoo..."
    SyncDividends True
    Application.StatusBar = False
    SetupButtons

    Application.Calculate
    ThisWorkbook.Worksheets(DASH_SHEET).Activate
    Application.ScreenUpdating = True
    MsgBox "Gotowe." & vbCrLf & vbCrLf & _
           "- nowy arkusz ""Nowe transakcje"" z przyciskiem ZAPISZ TRANSAKCJE," & vbCrLf & _
           "- nowy wyglπd Dashboardu (ciemny motyw, kafelki, wykresy, tabela pozycji z filtrem)," & vbCrLf & _
           "- poprawka cen z Londynu podawanych w pensach," & vbCrLf & _
           "- historia portfela zapisywana przy kaødej aktualizacji kursÛw," & vbCrLf & _
           "- arkusz ""Analiza"": stopa zwrotu skorygowana o wp≥aty i porÛwnanie z benchmarkiem," & vbCrLf & _
           "- arkusz ""Dywidendy"" z szacunkiem podatku Belki," & vbCrLf & _
           "- import historii z Trading 212 i XTB (przyciski IMPORT TRADING 212 / IMPORT XTB)." & vbCrLf & vbCrLf & _
           "Zapisz plik (Ctrl+S).", vbInformation, "Portfel inwestycyjny"
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Instalacja przerwana: " & Err.Description, vbExclamation, "Portfel inwestycyjny"
End Sub

Public Sub SetupButtons(Optional ByVal fromCode As Boolean = False)
    Dim ledger As Worksheet, entry As Worksheet

    Set ledger = ThisWorkbook.Worksheets("Transakcje")
    Set entry = ThisWorkbook.Worksheets(ENTRY_SHEET)

    ' Przyciski Dashboardu tworzy BuildDashboard.
    AddButton ThisWorkbook.Worksheets(ANALYSIS_SHEET), "btnRefreshAnalysis", "AKTUALIZUJ KURSY", _
              ThisWorkbook.Worksheets(ANALYSIS_SHEET).Range("E4").Left, ThisWorkbook.Worksheets(ANALYSIS_SHEET).Range("B1").Top + 4, 150, "RefreshPortfolio"
    AddButton ThisWorkbook.Worksheets(ANALYSIS_SHEET), "btnAnalysisToDashboard", "DASHBOARD", _
              ThisWorkbook.Worksheets(ANALYSIS_SHEET).Range("E4").Left + 160, ThisWorkbook.Worksheets(ANALYSIS_SHEET).Range("B1").Top + 4, 110, "GoToDashboard", False
    AddButton ThisWorkbook.Worksheets(DIVIDEND_SHEET), "btnSyncDividends", "POBIERZ DYWIDENDY", _
              ThisWorkbook.Worksheets(DIVIDEND_SHEET).Range("E1").Left, ThisWorkbook.Worksheets(DIVIDEND_SHEET).Range("B1").Top + 4, 150, "SyncDividendsNow", True
    AddButton ThisWorkbook.Worksheets(DIVIDEND_SHEET), "btnDividendsToDashboard", "DASHBOARD", _
              ThisWorkbook.Worksheets(DIVIDEND_SHEET).Range("E1").Left + 160, ThisWorkbook.Worksheets(DIVIDEND_SHEET).Range("B1").Top + 4, 110, "GoToDashboard", False
    AddButton ThisWorkbook.Worksheets(DIVIDEND_SHEET), "btnImportT212Div", "IMPORT TRADING 212", _
              ThisWorkbook.Worksheets(DIVIDEND_SHEET).Range("E1").Left + 280, ThisWorkbook.Worksheets(DIVIDEND_SHEET).Range("B1").Top + 4, 140, "ImportTrading212", False
    AddButton ThisWorkbook.Worksheets(DIVIDEND_SHEET), "btnImportXTBDiv", "IMPORT XTB", _
              ThisWorkbook.Worksheets(DIVIDEND_SHEET).Range("E1").Left + 430, ThisWorkbook.Worksheets(DIVIDEND_SHEET).Range("B1").Top + 4, 100, "ImportXTB", False
    AddButton ledger, "btnAddTransaction", "DODAJ TRANSAKCJE", ledger.Range("O2").Left, ledger.Range("O2").Top, 175, "OpenTransactionEntry"
    AddButton entry, "btnSaveTransactions", "ZAPISZ TRANSAKCJE", entry.Range("C1").Left, entry.Range("A1").Top + 2, 175, "SaveNewTransactions"
    AddButton entry, "btnClearTransactions", "WYCZYå∆", entry.Range("C1").Left + 185, entry.Range("A1").Top + 2, 90, "ClearNewTransactions"
End Sub

' Przycisk: g≥Ûwny (niebieski) albo pomocniczy (ciemny z obramowaniem).
Public Sub AddButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal caption As String, _
                      ByVal leftPos As Double, ByVal topPos As Double, ByVal widthPts As Double, _
                      ByVal macroName As String, Optional ByVal primary As Boolean = True)
    Dim shp As Shape
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, widthPts, 28)
    With shp
        .Name = shapeName
        .Placement = xlFreeFloating
        .Shadow.Visible = msoFalse
        .TextFrame.Characters.Text = caption
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .OnAction = macroName
        .Fill.Solid
        If primary Then
            .Fill.ForeColor.RGB = C_ACCENT
            .Line.Visible = msoFalse
            .TextFrame.Characters.Font.Color = C_BG
        Else
            .Fill.ForeColor.RGB = C_PANEL
            .Line.Visible = msoTrue
            .Line.ForeColor.RGB = C_BORDER
            .TextFrame.Characters.Font.Color = C_FG
        End If
        .TextFrame.Characters.Font.Name = THEME_FONT
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.Bold = True
    End With
End Sub

Public Sub GoToDashboard()
    ThisWorkbook.Worksheets(DASH_SHEET).Activate
End Sub

' ---------------------------------------------------------------------------
' ARKUSZ "ANALIZA"
' Stopa zwrotu skorygowana o wp≥aty (metoda TWR: zwroty miÍdzy kolejnymi
' wpisami historii, z wy≥πczeniem wp≥at/wyp≥at, mnoøone ≥aÒcuchowo)
' i porÛwnanie z benchmarkiem (domyúlnie VWCE, ceny z STOCKHISTORY).
' Dane: tabela HistoriaPortfela w ukrytym arkuszu Dane.
' ---------------------------------------------------------------------------
Public Sub GoToAnalysis()
    ThisWorkbook.Worksheets(ANALYSIS_SHEET).Activate
    ThisWorkbook.Worksheets(ANALYSIS_SHEET).Range("A1").Select
End Sub

Public Sub SetupAnalysisSheet(Optional ByVal fromCode As Boolean = False)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim labels As Variant
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ANALYSIS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Dashboard"))
        ws.Name = ANALYSIS_SHEET
    Else
        Do While ws.ChartObjects.Count > 0
            ws.ChartObjects(1).Delete
        Loop
        Do While ws.Shapes.Count > 0
            ws.Shapes(1).Delete
        Loop
        On Error Resume Next
        ws.Cells.Validation.Delete
        On Error GoTo 0
        ws.Cells.Clear
        ws.Columns.Hidden = False
    End If
    ThemeSheet ws
    ShowSheetLikeApp ws, True

    With ws
        .Columns("A").ColumnWidth = 2
        .Columns("B").ColumnWidth = 22
        .Columns("C:H").ColumnWidth = 19

        .Range("B1").Value = "Analiza portfela"
        .Range("B1").Font.Size = 20
        .Range("B1").Font.Bold = True
        .Rows(1).RowHeight = 34
        .Range("B1").VerticalAlignment = xlCenter
        .Range("B2").Value = "Stopa zwrotu skorygowana o wp≥aty (TWR) i porÛwnanie z benchmarkiem. Nowy wpis w historii powstaje przy kaødej aktualizacji kursÛw, raz dziennie."
        .Range("B2").Font.Color = C_MUTED

        ' Ustawienia
        .Range("B4").Value = "Okres w tabeli"
        .Range("C4").Value = "Miesiπc"
        .Range("B5").Value = "Benchmark (ticker)"
        .Range("C5").Value = "XETR:VWCE"
        .Range("B6").Value = "Waluta benchmarku"
        .Range("C6").Value = "EUR"
        .Range("D5").Value = "np. XETR:VWCE, XNAS:QQQ, XLON:VUSA - zapis jak w STOCKHISTORY"
        .Range("D5:D6").Font.Color = C_MUTED
        .Range("D6").Value = "waluta notowaÒ benchmarku (PLN, EUR, USD, GBP)"
        .Range("B7").Value = "Historia od / do"
        .Range("C7").Formula2 = "=IFERROR(MIN(HistoriaPortfela[Data]),"""")"
        .Range("D7").Formula2 = "=IFERROR(MAX(HistoriaPortfela[Data]),"""")"
        .Range("C7:D7").NumberFormat = "dd.mm.yyyy"
        .Range("C7:D7").HorizontalAlignment = xlLeft
        .Range("B4:B7").Font.Bold = True
        .Range("C4:C6").Interior.Color = C_INPUT
        .Range("C4:C6").Borders.Color = C_BORDER
        .Range("Y5:Y7").Value = Application.WorksheetFunction.Transpose(Array("DzieÒ", "Miesiπc", "Rok"))
        .Range("Z5:Z8").Value = Application.WorksheetFunction.Transpose(Array("PLN", "EUR", "USD", "GBP"))
        AddListValidation .Range("C4"), "=$Y$5:$Y$7", True
        AddListValidation .Range("C6"), "=$Z$5:$Z$8", True

        ' Ceny benchmarku i kursy walut (ukryte kolumny AA:AE)
        .Range("AA4").Value = "Benchmark: data, cena"
        .Range("AA5").Formula2 = "=IFERROR(STOCKHISTORY($C$5,INT(MIN(HistoriaPortfela[Data]))-10,TODAY(),0,0,0,1),""brak"")"
        .Range("AD4").Value = "Kurs waluty: data, kurs"
        .Range("AD5").Formula2 = "=IF($C$6=""PLN"",""PLN"",IFERROR(STOCKHISTORY($C$6&""/PLN"",INT(MIN(HistoriaPortfela[Data]))-10,TODAY(),0,0,0,1),""brak""))"
    End With

    ' Kolumny obliczeniowe w tabeli historii (wymagajπ arkusza Analiza).
    Set lo = SetupHistoryColumns()

    With ws
        ' Wskaüniki
        labels = Array("WartoúÊ portfela", "Stopa zwrotu portfela (TWR)", "Stopa zwrotu benchmarku", _
                       "Portfel - benchmark (pkt proc.)", "Gdyby wp≥aty sz≥y w benchmark", "Portfel - benchmark (PLN)")
        For i = 0 To 5
            .Cells(9, 2 + i).Value = labels(i)
        Next i
        With .Range("B9:G9")
            .Font.Size = 9
            .Font.Color = C_MUTED
            .Interior.Color = C_PANEL
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        .Rows(9).RowHeight = 28
        .Range("B10").Formula2 = "=IFERROR(INDEX(HistoriaPortfela[WartoúÊ portfela],ROWS(HistoriaPortfela[WartoúÊ portfela])),"""")"
        .Range("C10").Formula2 = "=IFERROR(INDEX(HistoriaPortfela[Stopa portfela],ROWS(HistoriaPortfela[Stopa portfela])),"""")"
        .Range("D10").Formula2 = "=IFERROR(INDEX(HistoriaPortfela[Stopa benchmarku],ROWS(HistoriaPortfela[Stopa benchmarku])),"""")"
        .Range("E10").Formula2 = "=IFERROR(C10-D10,"""")"
        .Range("F10").Formula2 = "=IFERROR(INDEX(HistoriaPortfela[WartoúÊ w benchmarku],ROWS(HistoriaPortfela[WartoúÊ w benchmarku])),"""")"
        .Range("G10").Formula2 = "=IFERROR(B10-F10,"""")"
        With .Range("B10:G10")
            .Font.Size = 15
            .Font.Bold = True
            .Font.Color = C_FG
            .Interior.Color = C_PANEL
            .HorizontalAlignment = xlLeft
        End With
        .Rows(10).RowHeight = 30
        .Range("B10,F10,G10").NumberFormat = "#,##0.00 ""z≥"";[Red]-#,##0.00 ""z≥"""
        .Range("C10:E10").NumberFormat = "0.00%;[Red]-0.00%"

        ' Tabela okresÛw
        .Range("B33").Value = "Wyniki w okresach (najnowsze na gÛrze) - okres zmienisz w komÛrce C4"
        .Range("B33").Font.Bold = True
        .Range("B34:H34").Value = Array("Okres", "WartoúÊ na koniec", "Wp≥aty netto w okresie", _
                                        "Zysk w okresie (PLN)", "Stopa zwrotu portfela", "Stopa zwrotu benchmarku", "RÛønica (pkt proc.)")
        With .Range("B34:H34")
            .Font.Bold = True
            .Font.Color = C_MUTED
            .Interior.Color = C_PANEL
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        .Rows(34).RowHeight = 30
        .Range("B35").Formula2 = PeriodTableFormula()
        .Range("B35:H2000").Interior.Color = C_PANEL
        .Range("B35:H2000").Borders(xlInsideHorizontal).Color = C_ROWLINE
        With .Range("F35:H2000")
            .FormatConditions.Delete
            .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0").Font.Color = C_OK
            .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0").Font.Color = C_BAD
        End With
        .Range("C35:E2000").NumberFormat = "#,##0.00;[Red]-#,##0.00"
        .Range("F35:H2000").NumberFormat = "0.00%;[Red]-0.00%"
        .Range("B35:B2000").HorizontalAlignment = xlLeft

        .Columns("Y:AE").Hidden = True
    End With

    BuildAnalysisCharts ws, lo
End Sub

Public Function PeriodTableFormula() As String
    PeriodTableFormula = "=IFERROR(LET(" & _
        "dni,INT(HistoriaPortfela[Data]),okres,$C$4," & _
        "keyArr,IF(okres=""Rok"",YEAR(dni),IF(okres=""Miesiπc"",YEAR(dni)*100+MONTH(dni),dni))," & _
        "uniq,UNIQUE(keyArr),seq,SEQUENCE(ROWS(uniq))," & _
        "endIx,XMATCH(uniq,keyArr,0,-1)," & _
        "prvIx,IF(seq=1,1,INDEX(endIx,IF(seq=1,1,seq-1)))," & _
        "valA,HistoriaPortfela[WartoúÊ portfela],depA,HistoriaPortfela[Wp≥aty netto]," & _
        "idxA,HistoriaPortfela[Indeks portfela],bixA,HistoriaPortfela[Indeks benchmarku]," & _
        "vEnd,INDEX(valA,endIx),vSt,INDEX(valA,prvIx),dEnd,INDEX(depA,endIx),dSt,INDEX(depA,prvIx)," & _
        "rPort,INDEX(idxA,endIx)/INDEX(idxA,prvIx)-1," & _
        "rBen,IFERROR(INDEX(bixA,endIx)/INDEX(bixA,prvIx)-1,"""")," & _
        "lbl,IF(okres=""Rok"",""""&uniq,IF(okres=""Miesiπc"",INT(uniq/100)&""-""&RIGHT(""0""&MOD(uniq,100),2)," & _
            "YEAR(uniq)&""-""&RIGHT(""0""&MONTH(uniq),2)&""-""&RIGHT(""0""&DAY(uniq),2)))," & _
        "SORTBY(HSTACK(lbl,vEnd,dEnd-dSt,vEnd-vSt-(dEnd-dSt),rPort,rBen,IFERROR(rPort-rBen,"""")),uniq,-1))," & _
        """Brak danych w historii portfela"")"
End Function

' Dodaje do tabeli HistoriaPortfela kolumny obliczeniowe (wype≥niajπ siÍ same
' dla kaødego nowego wpisu) i usuwa wpisy bez wartoúci portfela.
Public Function SetupHistoryColumns() As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim specs As Variant
    Dim firstRow As String

    Set lo = GetHistoryTable()
    Set ws = lo.Parent

    For i = lo.ListRows.Count To 1 Step -1
        If Len(Trim$(CStr(lo.ListRows(i).Range.Cells(1, 2).Value))) = 0 Then lo.ListRows(i).Delete
    Next i

    firstRow = "ROW()=ROW(HistoriaPortfela[#Headers])+1"
    specs = Array( _
        Array("Przep≥yw", "=IF(" & firstRow & ",0,[@[Wp≥aty netto]]-OFFSET([@[Wp≥aty netto]],-1,0))", "#,##0.00"), _
        Array("Zwrot okresu", "=IF(" & firstRow & ",0,IFERROR([@[WartoúÊ portfela]]/(OFFSET([@[WartoúÊ portfela]],-1,0)+[@Przep≥yw])-1,0))", "0.00%"), _
        Array("Indeks portfela", "=IF(" & firstRow & ",1,OFFSET([@[Indeks portfela]],-1,0)*(1+[@[Zwrot okresu]]))", "0.0000"), _
        Array("Stopa portfela", "=[@[Indeks portfela]]-1", "0.00%"), _
        Array("Cena benchmarku (PLN)", "=IFERROR(XLOOKUP(INT([@Data]),INDEX(Analiza!$AA$5#,0,1),INDEX(Analiza!$AA$5#,0,2),,-1)*" & _
              "IF(Analiza!$C$6=""PLN"",1,XLOOKUP(INT([@Data]),INDEX(Analiza!$AD$5#,0,1),INDEX(Analiza!$AD$5#,0,2),,-1))," & _
              "IF(" & firstRow & ","""",OFFSET([@[Cena benchmarku (PLN)]],-1,0)))", "#,##0.00"), _
        Array("Indeks benchmarku", "=IFERROR([@[Cena benchmarku (PLN)]]/INDEX([Cena benchmarku (PLN)],1),"""")", "0.0000"), _
        Array("Stopa benchmarku", "=IFERROR([@[Indeks benchmarku]]-1,NA())", "0.00%"), _
        Array("Jednostki benchmarku", "=IFERROR(IF(" & firstRow & ",[@[WartoúÊ portfela]]/[@[Cena benchmarku (PLN)]]," & _
              "OFFSET([@[Jednostki benchmarku]],-1,0)+[@Przep≥yw]/[@[Cena benchmarku (PLN)]]),"""")", "0.0000"), _
        Array("WartoúÊ w benchmarku", "=IFERROR([@[Jednostki benchmarku]]*[@[Cena benchmarku (PLN)]],NA())", "#,##0.00"))

    For i = LBound(specs) To UBound(specs)
        If Not ListColumnExists(lo, CStr(specs(i)(0))) Then
            lo.ListColumns.Add.Name = CStr(specs(i)(0))
        End If
    Next i
    If lo.ListRows.Count > 0 Then
        For i = LBound(specs) To UBound(specs)
            With lo.ListColumns(CStr(specs(i)(0))).DataBodyRange
                .Formula2 = CStr(specs(i)(1))
                .NumberFormat = CStr(specs(i)(2))
            End With
        Next i
    End If
    Set SetupHistoryColumns = lo
End Function

Public Function ListColumnExists(ByVal lo As ListObject, ByVal colName As String) As Boolean
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then
            ListColumnExists = True
            Exit Function
        End If
    Next lc
End Function

Public Sub BuildAnalysisCharts(ByVal ws As Worksheet, ByVal lo As ListObject)
    Dim co As ChartObject
    Dim topPos As Double, leftPos As Double, w As Double, h As Double

    If lo.ListRows.Count = 0 Then Exit Sub
    topPos = ws.Range("B12").Top
    leftPos = ws.Range("B12").Left
    w = ws.Range("B12:E12").Width - 6
    h = ws.Range("B12:B31").Height

    Set co = ws.ChartObjects.Add(leftPos, topPos, w, h)
    co.Name = "WykresWartosciPortfela"
    PrepareLineChart co.Chart, "WartoúÊ portfela, wp≥aty i benchmark (PLN)"
    AddLineSeries co.Chart, "WartoúÊ portfela", lo, "WartoúÊ portfela", C_ACCENT
    AddLineSeries co.Chart, "Wp≥aty netto", lo, "Wp≥aty netto", C_MUTED
    AddLineSeries co.Chart, "Gdyby wp≥aty sz≥y w benchmark", lo, "WartoúÊ w benchmarku", C_ACCENT2
    FormatChartAxes co.Chart, "#,##0"

    Set co = ws.ChartObjects.Add(leftPos + w + 12, topPos, ws.Range("F12:H12").Width + 40, h)
    co.Name = "WykresStopyZwrotu"
    PrepareLineChart co.Chart, "Skumulowana stopa zwrotu (TWR) vs benchmark"
    AddLineSeries co.Chart, "Portfel", lo, "Stopa portfela", C_ACCENT
    AddLineSeries co.Chart, "Benchmark", lo, "Stopa benchmarku", C_ACCENT2
    FormatChartAxes co.Chart, "0.0%"
End Sub

Public Sub PrepareLineChart(ByVal ch As Chart, ByVal title As String)
    ch.ChartType = xlLine
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
    ch.PlotVisibleOnly = False          ' dane historii sπ w ukrytym arkuszu
    ch.DisplayBlanksAs = xlInterpolated
    ch.HasLegend = True
    ch.HasTitle = True
    ch.ChartTitle.Text = title
End Sub

Public Sub AddLineSeries(ByVal ch As Chart, ByVal seriesName As String, ByVal lo As ListObject, _
                          ByVal valueColumn As String, ByVal colorRGB As Long)
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.Name = seriesName
    s.XValues = lo.ListColumns("Data").DataBodyRange
    s.Values = lo.ListColumns(valueColumn).DataBodyRange
    s.Format.Line.ForeColor.RGB = colorRGB
    s.Format.Line.Weight = 2.25
    s.MarkerStyle = xlMarkerStyleCircle
    s.MarkerSize = 3
    s.MarkerBackgroundColor = colorRGB
    s.MarkerForegroundColor = colorRGB
End Sub

Public Sub FormatChartAxes(ByVal ch As Chart, ByVal valueFormat As String)
    On Error Resume Next
    With ch.Axes(xlCategory)
        .CategoryType = xlTimeScale
        .BaseUnit = xlDays
        .TickLabels.NumberFormat = "dd.mm"
    End With
    ch.Axes(xlValue).TickLabels.NumberFormat = valueFormat
    StyleChartDark ch, ch.ChartTitle.Text
End Sub

' ---------------------------------------------------------------------------
' DASHBOARD (ciemny motyw)
' Liczby sπ liczone w ukrytym arkuszu "Dane"; kafelki to kszta≥ty z tekstem
' po≥πczonym z komÛrkami, wiÍc odúwieøajπ siÍ same. Kolory zysku/straty
' i s≥upkÛw ustawia RefreshDashboardVisuals (po kaødej aktualizacji kursÛw).
' ---------------------------------------------------------------------------
Public Function GetHistoryTable() As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = HISTORY_TABLE Then
                Set GetHistoryTable = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

' Ukryty arkusz "Dane": tabela historii (przeniesiona z Dashboardu) i wszystkie
' liczby, z ktÛrych korzysta Dashboard.
Public Sub SetupDataSheet(Optional ByVal fromCode As Boolean = False)
    Dim wsD As Worksheet
    Dim lo As ListObject
    Dim vals As Variant
    Dim nRows As Long
    Dim specs As Collection
    Dim spec As Variant
    Dim i As Long

    On Error Resume Next
    Set wsD = ThisWorkbook.Worksheets(DATA_SHEET)
    On Error GoTo 0
    If wsD Is Nothing Then
        Set wsD = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsD.Name = DATA_SHEET
    End If

    Set lo = GetHistoryTable()
    If lo Is Nothing Then
        wsD.Range("A1:C1").Value = Array("Data", "WartoúÊ portfela", "Wp≥aty netto")
        Set lo = wsD.ListObjects.Add(xlSrcRange, wsD.Range("A1:C2"), , xlYes)
        lo.Name = HISTORY_TABLE
    ElseIf lo.Parent.Name <> DATA_SHEET Then
        vals = lo.Range.Resize(, 3).Value
        nRows = UBound(vals, 1)
        lo.Delete
        wsD.Range("A1").Resize(nRows, 3).Value = vals
        If nRows < 2 Then nRows = 2
        Set lo = wsD.ListObjects.Add(xlSrcRange, wsD.Range("A1").Resize(nRows, 3), , xlYes)
        lo.Name = HISTORY_TABLE
    End If
    lo.TableStyle = "TableStyleMedium2"
    ' DzieÒ salda poczπtkowego (dla importu z brokerÛw): pierwszy wpis historii portfela.
    If Not IsDate(wsD.Range("AB32").Value) And Not lo.DataBodyRange Is Nothing Then
        If IsDate(lo.ListRows(1).Range.Cells(1, 1).Value) Then
            wsD.Range("AB32").Value = Int(CDbl(CDate(lo.ListRows(1).Range.Cells(1, 1).Value)))
            wsD.Range("AB32").NumberFormat = "dd.mm.yyyy"
        End If
    End If
    If Not lo.DataBodyRange Is Nothing Then
        lo.ListColumns(1).DataBodyRange.NumberFormat = "dd.mm.yyyy hh:mm"
        lo.ListColumns(2).DataBodyRange.NumberFormat = "#,##0.00"
        lo.ListColumns(3).DataBodyRange.NumberFormat = "#,##0.00"
    End If

    ' Kluczowe liczby: etykieta w AA, formu≥a w AB.
    Set specs = New Collection
    specs.Add Array(2, "WartoúÊ portfela", "=SUM(Pozycje!$G$5:$G$500)+AB6", "#,##0.00 ""z≥"";-#,##0.00 ""z≥""")
    specs.Add Array(3, "Wp≥aty netto", "=Transakcje!$R$7", "#,##0.00 ""z≥"";-#,##0.00 ""z≥""")
    specs.Add Array(4, "Zysk / strata", "=AB2-AB3", "+#,##0.00 ""z≥"";-#,##0.00 ""z≥"";0.00 ""z≥""")
    specs.Add Array(5, "Stopa zwrotu", "=IF(AB3=0,"""",AB4/AB3)", "+0.0%;-0.0%;0.0%")
    specs.Add Array(6, "GotÛwka", "=Transakcje!$R$6", "#,##0.00 ""z≥"";-#,##0.00 ""z≥""")
    specs.Add Array(7, "Udzia≥ gotÛwki", "=IF(AB2=0,"""",AB6/AB2)", "0.0%")
    specs.Add Array(8, "Poprzedni odczyt - wartoúÊ", "=IFERROR(LET(daty," & HISTORY_TABLE & "[Data],wart," & HISTORY_TABLE & "[WartoúÊ portfela],ileW,ROWS(wart)," & _
              "pozW,IF(INT(INDEX(daty,ileW))=TODAY(),ileW-1,ileW),IF(pozW<1,"""",INDEX(wart,pozW))),"""")", "#,##0.00 ""z≥""")
    specs.Add Array(9, "Poprzedni odczyt - data", "=IFERROR(LET(daty," & HISTORY_TABLE & "[Data],ileW,ROWS(daty)," & _
              "pozW,IF(INT(INDEX(daty,ileW))=TODAY(),ileW-1,ileW),IF(pozW<1,"""",INDEX(daty,pozW))),"""")", "dd.mm.yyyy hh:mm")
    specs.Add Array(10, "Zmiana od odczytu", "=IFERROR(IF(AB8="""","""",AB2/AB8-1),"""")", ChrW(9650) & " 0.0%;" & ChrW(9660) & " 0.0%;0.0%")
    specs.Add Array(11, "Kursy z", "=IFERROR(IF(COUNT(" & HISTORY_TABLE & "[Data])=0,"""",MAX(" & HISTORY_TABLE & "[Data])),"""")", "dd.mm.yyyy hh:mm")
    specs.Add Array(12, "Najlepsza pozycja", "=IFERROR(INDEX(AH2#,1,1),"""")", "General")
    specs.Add Array(13, "Najs≥absza pozycja", "=IFERROR(INDEX(AH2#,ROWS(AH2#),1),"""")", "General")
    specs.Add Array(14, "Wynik najs≥abszej", "=IFERROR(INDEX(AH2#,ROWS(AH2#),2),0)", "#,##0.00")
    specs.Add Array(15, "Liczba pozycji", "=SUMPRODUCT((Pozycje!$B$5:$B$500<>"""")*ISNUMBER(Pozycje!$H$5:$H$500))", "0")
    specs.Add Array(16, "Zdanie podsumowania", "=""Portfel jest wart ""&FIXED(AB2,2)&"" z≥, czyli ""&IF(AB4>=0,FIXED(AB4,2)&"" z≥ ponad wp≥aty."",FIXED(-AB4,2)&"" z≥ poniøej wp≥at."")" & _
              "&IF(AB12="""","""","" NajwiÍcej zarobi≥a pozycja ""&AB12&IF(AB14<0,"", najwiÍcej traci ""&AB13&""."","".""))", "General")
    specs.Add Array(17, "Tekst od odczytu", "=IF(ISNUMBER(AB9),""od odczytu ""&RIGHT(""0""&DAY(AB9),2)&"".""&RIGHT(""0""&MONTH(AB9),2)&"" ""&RIGHT(""0""&HOUR(AB9),2)&"":""&RIGHT(""0""&MINUTE(AB9),2),""brak wczeúniejszego odczytu"")", "General")
    specs.Add Array(18, "Tekst kursy z", "=IF(ISNUMBER(AB11),""Kursy z ""&RIGHT(""0""&DAY(AB11),2)&"".""&RIGHT(""0""&MONTH(AB11),2)&"".""&YEAR(AB11)&"" ""&RIGHT(""0""&HOUR(AB11),2)&"":""&RIGHT(""0""&MINUTE(AB11),2),""Kursy: brak aktualizacji"")", "General")
    specs.Add Array(19, "Tekst gotÛwki", "=IFERROR(FIXED(AB7*100,1)&""% portfela w gotÛwce"","""")", "General")
    specs.Add Array(20, "Dywidendy netto - 12 mies.", "=IFERROR(SUMPRODUCT((INDEX(AP2#,0,1)>EDATE(TODAY(),-12))*INDEX(AP2#,0,5)),0)", "#,##0.00 ""z≥""")
    specs.Add Array(21, "Dywidendy netto - ≥πcznie", "=IFERROR(SUM(INDEX(AP2#,0,5)),0)", "#,##0.00 ""z≥""")
    specs.Add Array(22, "Dywidendy netto - bieøπcy rok", "=IFERROR(SUMPRODUCT((YEAR(INDEX(AP2#,0,1))=YEAR(TODAY()))*INDEX(AP2#,0,5)),0)", "#,##0.00 ""z≥""")
    specs.Add Array(23, "Koszt pozycji wyp≥acajπcych", "=SUMIF(Pozycje!$C$5:$C$500,""Tak"",Pozycje!$F$5:$F$500)", "#,##0.00 ""z≥""")
    specs.Add Array(24, "Stopa od kosztu (prognoza)", "=IFERROR(AB28/SUMIFS(" & DIV_FORECAST_TABLE & "[Koszt pozycji (PLN)]," & DIV_FORECAST_TABLE & "[Prognoza 12 mies. netto (PLN)],"">0""),0)", "0.00%")
    specs.Add Array(25, "SpÛ≥ki wyp≥acajπce", "=IFERROR(COUNTIF(" & DIV_FORECAST_TABLE & "[Prognoza 12 mies. netto (PLN)],"">0""),0)", "0")
    specs.Add Array(26, "Tekst dywidend", "=""otrzymane w 12 mies.: ""&FIXED(AB20,2)&"" z≥""", "General")
    specs.Add Array(27, "Dywidendy - úrednio miesiÍcznie", "=AB20/12", "#,##0.00 ""z≥""")
    specs.Add Array(28, "Prognoza dywidend 12 mies. (netto)", "=IFERROR(SUM(" & DIV_FORECAST_TABLE & "[Prognoza 12 mies. netto (PLN)]),0)", "#,##0.00 ""z≥""")
    specs.Add Array(29, "Stopa dywidendy portfela", "=IFERROR(AB28/(AB2-AB6),0)", "0.00%")
    wsD.Range("AA1:AB1").Value = Array("Kluczowe liczby (Dashboard)", "WartoúÊ")
    For Each spec In specs
        wsD.Cells(spec(0), "AA").Value = spec(1)
        wsD.Cells(spec(0), "AB").NumberFormat = spec(3)
    Next spec

    ' Sk≥ad portfela: klasy (wiersze 2-4) i waluty (wiersze 7-10).
    wsD.Range("AD1:AF1").Value = Array("Klasa / waluta", "WartoúÊ (PLN)", "Udzia≥")
    wsD.Range("AD2:AD4").Value = Application.WorksheetFunction.Transpose(Array("Akcje", "ETF-y", "GotÛwka"))
    wsD.Range("AE2").Formula2 = "=SUMIF(Pozycje!$A$5:$A$500,AD2,Pozycje!$G$5:$G$500)"
    wsD.Range("AE3").Formula2 = "=SUMIF(Pozycje!$A$5:$A$500,AD3,Pozycje!$G$5:$G$500)"
    wsD.Range("AE4").Formula2 = "=AB6"
    wsD.Range("AD6").Value = "Waluta notowania (gotÛwka jako PLN)"
    wsD.Range("AD7:AD10").Value = Application.WorksheetFunction.Transpose(Array("PLN", "USD", "EUR", "GBP"))
    For i = 7 To 10
        wsD.Cells(i, "AE").Formula2 = "=SUMIF(Notowania!$F$7:$F$502,AD" & i & ",Notowania!$H$7:$H$502)" & IIf(i = 7, "+AB6", "")
    Next i
    For i = 2 To 10
        If i <> 5 And i <> 6 Then wsD.Cells(i, "AF").Formula2 = "=IFERROR(AE" & i & "/$AB$2,0)"
    Next i
    wsD.Range("AE2:AE10").NumberFormat = "#,##0.00 ""z≥"""
    wsD.Range("AF2:AF10").NumberFormat = "0.0%"

    ' Zysk/strata kaødej pozycji, posortowane (dla wykresu s≥upkowego).
    wsD.Range("AH1:AK1").Value = Array("Pozycja", "Zysk / strata (PLN)", "Zwrot", "Etykieta wykresu")
    wsD.Range("AH2").Formula2 = "=LET(nazwy,Pozycje!$B$5:$B$500,zyski,Pozycje!$H$5:$H$500,zwroty,Pozycje!$I$5:$I$500," & _
        "maska,(nazwy<>"""")*ISNUMBER(zyski),IF(SUM(maska)=0,"""",SORTBY(HSTACK(FILTER(nazwy,maska),FILTER(zyski,maska),FILTER(zwroty,maska)),FILTER(zyski,maska),-1)))"
    wsD.Range("AK2").Formula2 = "=IFERROR(INDEX(AH2#,0,1)&""  (""&IF(INDEX(AH2#,0,3)>0,""+"","""")&FIXED(INDEX(AH2#,0,3)*100,1)&""%)"",INDEX(AH2#,0,1))"
    wsD.Range("AI2:AI500").NumberFormat = "#,##0.00"
    wsD.Range("AJ2:AJ500").NumberFormat = "0.0%"

    ' Dywidendy: wyp≥aty z rejestru transakcji + tabela DywidendyHistoria (otrzymane i automatyczne).
    ' Kolumny: data, spÛ≥ka, kwota, podatek, netto, ürÛd≥o - najnowsze na gÛrze.
    SetupDividendTables wsD
    wsD.Range("AP1:AU1").Value = Array("Data", "SpÛ≥ka", "Kwota", "Podatek", "Netto", "èrÛd≥o")
    wsD.Range("AP2").Formula2 = "=LET(daty,Transakcje!$A$30:$A$500,nazwy,Transakcje!$B$30:$B$500,typy,Transakcje!$C$30:$C$500," & _
        "kwoty,Transakcje!$E$30:$E$500,podatki,Transakcje!$F$30:$F$500,kontrola,Transakcje!$L$30:$L$500," & _
        "maskaR,(typy=""Dywidenda"")*(kontrola=""OK"")*ISNUMBER(daty),podatkiN,IF(ISNUMBER(podatki),podatki,0)," & _
        "tDaty,IFERROR(" & DIV_HISTORY_TABLE & "[Data],""""),tNazwy,IFERROR(" & DIV_HISTORY_TABLE & "[SpÛ≥ka],"""")," & _
        "tKwoty,IFERROR(" & DIV_HISTORY_TABLE & "[Kwota netto (PLN)],0),tZrodla,IFERROR(" & DIV_HISTORY_TABLE & "[èrÛd≥o],"""")," & _
        "tPodatki,IFERROR(" & DIV_HISTORY_TABLE & "[Podatek u ürÛd≥a (PLN)],0),tPodN,IF(ISNUMBER(tPodatki),tPodatki,0)," & _
        "maskaT,ISNUMBER(tDaty)*(tNazwy<>"""")," & _
        "wszystkie,VSTACK(HSTACK(daty,nazwy,kwoty,podatkiN,kwoty-podatkiN,IF(ISNUMBER(daty)+1,""rejestr transakcji""))," & _
        "HSTACK(tDaty,tNazwy,tKwoty+tPodN,tPodN,tKwoty,tZrodla))," & _
        "maska,VSTACK(maskaR,maskaT)," & _
        "IF(SUM(maska)=0,"""",SORTBY(FILTER(wszystkie,maska),FILTER(INDEX(wszystkie,0,1),maska),-1)))"
    wsD.Range("AP2:AP2000").NumberFormat = "dd.mm.yyyy"
    wsD.Range("AR2:AT2000").NumberFormat = "#,##0.00"
    ' Dywidendy w ostatnich 24 miesiπcach (dla wykresu).
    wsD.Range("AW1:AY1").Value = Array("Miesiπc", "Dywidendy netto", "Etykieta")
    For i = 2 To 25
        wsD.Cells(i, "AW").Formula2 = "=EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1)," & (i - 25) & ")"
        wsD.Cells(i, "AX").Formula2 = "=IFERROR(SUMPRODUCT((INDEX($AP$2#,0,1)>=AW" & i & ")*(INDEX($AP$2#,0,1)<EDATE(AW" & i & ",1))*INDEX($AP$2#,0,5)),0)"
        wsD.Cells(i, "AY").Formula2 = "=RIGHT(""0""&MONTH(AW" & i & "),2)&"".""&RIGHT(YEAR(AW" & i & "),2)"
    Next i
    wsD.Range("AW2:AW25").NumberFormat = "mm.yyyy"

    ' Listy wyboru.
    wsD.Range("AN1").Value = "Filtr grupy"
    wsD.Range("AN2:AN4").Value = Application.WorksheetFunction.Transpose(Array("Wszystkie", "Akcje", "ETF-y"))

    ' Formu≥y AB wpisywane na koÒcu (korzystajπ z AH2#).
    For Each spec In specs
        wsD.Cells(spec(0), "AB").Formula2 = spec(2)
    Next spec

    wsD.Columns("AA").ColumnWidth = 30
    wsD.Columns("AB").ColumnWidth = 22
    wsD.Columns("AD").ColumnWidth = 30
    wsD.Columns("AH").ColumnWidth = 30
    wsD.Columns("AK").ColumnWidth = 40
    wsD.Visible = xlSheetHidden
End Sub

Public Sub ThemeSheet(ByVal ws As Worksheet)
    With ws.Cells
        .Interior.Color = C_BG
        .Font.Name = THEME_FONT
        .Font.Size = 10
        .Font.Color = C_FG
    End With
End Sub

Public Sub ShowSheetLikeApp(ByVal ws As Worksheet, ByVal hideHeadings As Boolean)
    Dim prev As Object
    Set prev = ActiveSheet
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = Not hideHeadings
    If Not prev Is Nothing Then
        If prev.Visible = xlSheetVisible Then prev.Activate
    End If
End Sub

Public Function AddPanel(ByVal ws As Worksheet, ByVal shapeName As String, ByVal l As Double, ByVal t As Double, _
                          ByVal w As Double, ByVal h As Double, ByVal fillColor As Long, ByVal lineColor As Long) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, l, t, w, h)
    With shp
        .Name = shapeName
        .Placement = xlFreeFloating
        .Adjustments.Item(1) = 10 / IIf(w < h, w, h)
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = lineColor
        .Line.Weight = 0.75
        .Shadow.Visible = msoFalse
    End With
    Set AddPanel = shp
End Function

Public Function AddText(ByVal ws As Worksheet, ByVal shapeName As String, ByVal txt As String, _
                         ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal h As Double, _
                         ByVal fontSize As Single, ByVal isBold As Boolean, ByVal fontColor As Long, _
                         Optional ByVal linkTo As String = "", Optional ByVal align As Long = 1) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, l, t, w, h)
    shp.Name = shapeName
    shp.Placement = xlFreeFloating
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    If Len(linkTo) > 0 Then
        shp.DrawingObject.Formula = linkTo
    Else
        shp.TextFrame2.TextRange.Text = txt
    End If
    With shp.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoTrue
        .AutoSize = msoAutoSizeNone
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = align    ' 1 = do lewej, 3 = do prawej
        With .TextRange.Font
            .Name = THEME_FONT
            .Size = fontSize
            .Bold = IIf(isBold, msoTrue, msoFalse)
            .Fill.ForeColor.RGB = fontColor
        End With
    End With
    Set AddText = shp
End Function

Public Sub StyleChartDark(ByVal ch As Chart, ByVal title As String)
    On Error Resume Next
    With ch
        .ChartArea.Format.Fill.Visible = msoTrue
        .ChartArea.Format.Fill.Solid
        .ChartArea.Format.Fill.ForeColor.RGB = C_PANEL
        .ChartArea.Format.Line.Visible = msoTrue
        .ChartArea.Format.Line.ForeColor.RGB = C_BORDER
        .ChartArea.Format.Line.Weight = 0.75
        .ChartArea.RoundedCorners = True
        .ChartArea.Font.Name = THEME_FONT
        .ChartArea.Font.Size = 9
        .ChartArea.Font.Color = C_MUTED
        .PlotArea.Format.Fill.Visible = msoFalse
        .HasTitle = True
        .ChartTitle.Text = title
        .ChartTitle.Font.Size = 11
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = C_FG
        .ChartTitle.Position = xlChartElementPositionAutomatic
        .ChartTitle.Left = 12
        If .HasLegend Then
            .Legend.Position = xlLegendPositionTop
            .Legend.Font.Color = C_MUTED
        End If
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = C_BORDER
        .Axes(xlValue).Format.Line.Visible = msoFalse
        .Axes(xlValue).TickLabels.Font.Color = C_MUTED
        .Axes(xlCategory).Format.Line.ForeColor.RGB = C_BORDER
        .Axes(xlCategory).TickLabels.Font.Color = C_MUTED
        .Axes(xlCategory).MajorTickMark = xlNone
    End With
End Sub

' Pasek 100% (jedna kategoria, jedna seria na kaødπ czÍúÊ).
Public Sub AddStackBar(ByVal ws As Worksheet, ByVal chartName As String, ByVal l As Double, ByVal t As Double, _
                        ByVal w As Double, ByVal valueCells As Range, ByVal partNames As Variant, ByVal colors As Variant)
    Dim co As ChartObject
    Dim s As Series
    Dim i As Long

    Set co = ws.ChartObjects.Add(l, t, w, 14)
    co.Name = chartName
    co.Placement = xlFreeFloating
    With co.Chart
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        For i = 1 To valueCells.Cells.Count
            Set s = .SeriesCollection.NewSeries
            s.Name = partNames(i - 1)
            s.Values = valueCells.Cells(i)
            s.Format.Fill.ForeColor.RGB = colors(i - 1)
            s.Format.Line.Visible = msoFalse
        Next i
        .ChartType = xlBarStacked100
        .HasTitle = False
        .HasLegend = False
        .ChartGroups(1).GapWidth = 0
        .ChartGroups(1).Overlap = 100
        On Error Resume Next
        .Axes(xlValue).HasMajorGridlines = False
        .HasAxis(xlCategory, xlPrimary) = False
        .HasAxis(xlValue, xlPrimary) = False
        .ChartArea.Format.Fill.Visible = msoTrue
        .ChartArea.Format.Fill.ForeColor.RGB = C_TRACK
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.Visible = msoFalse
        .PlotArea.InsideLeft = 0
        .PlotArea.InsideTop = 0
        .PlotArea.InsideWidth = w
        .PlotArea.InsideHeight = 14
        On Error GoTo 0
    End With
End Sub

' Wiersz listy: kolorowy kwadrat, nazwa, wartoúÊ, udzia≥.
Public Sub AddLegendRow(ByVal ws As Worksheet, ByVal key As String, ByVal l As Double, ByVal t As Double, ByVal w As Double, _
                         ByVal label As String, ByVal colorRGB As Long, ByVal valueRef As String, ByVal shareRef As String)
    Dim sw As Shape
    Set sw = ws.Shapes.AddShape(msoShapeRoundedRectangle, l, t + 4, 9, 9)
    sw.Name = "sw" & key
    sw.Fill.ForeColor.RGB = colorRGB
    sw.Line.Visible = msoFalse
    sw.Placement = xlFreeFloating
    AddText ws, "lbl" & key, label, l + 16, t, w * 0.42, 17, 10, False, C_FG
    AddText ws, "val" & key, "", l + 16 + w * 0.42, t, w * 0.36, 17, 10, False, C_FG, valueRef, 3
    AddText ws, "pct" & key, "", l + w - 52, t, 52, 17, 10, False, C_MUTED, shareRef, 3
End Sub

Public Sub BuildDashboard(Optional ByVal fromCode As Boolean = False)
    Dim ws As Worksheet, wsD As Worksheet
    Dim lo As ListObject
    Dim co As ChartObject
    Dim s As Series
    Dim x0 As Double, totalW As Double, y As Double, gap As Double
    Dim cardW As Double, cardL As Double, lineW As Double, compL As Double, compW As Double
    Dim plTop As Double, plH As Double, yT As Double
    Dim i As Long, n As Long, r As Long
    Dim tblRow As Long, ctrlRow As Long, hdrRow As Long, dataRow As Long, lastRow As Long
    Dim kpiLabels As Variant, kpiRefs As Variant, kpiSubs As Variant
    Dim spill As String

    Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Set wsD = ThisWorkbook.Worksheets(DATA_SHEET)
    Set lo = GetHistoryTable()
    Application.Calculate
    n = 1
    If IsNumeric(wsD.Range("AB15").Value) Then n = CLng(wsD.Range("AB15").Value)
    If n < 1 Then n = 1

    ' 1) Czyszczenie arkusza.
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
    ws.Cells.Clear
    ws.Columns.Hidden = False
    ws.Rows.Hidden = False
    ws.Cells.RowHeight = 18
    ThemeSheet ws
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C:L").ColumnWidth = 12.5
    ws.Columns("M").ColumnWidth = 3
    ShowSheetLikeApp ws, True

    x0 = ws.Range("B1").Left
    totalW = ws.Range("B1:L1").Width
    gap = 12

    ' 2) Nag≥Ûwek.
    AddText ws, "txtEyebrow", "M”J PORTFEL", x0, 12, 200, 14, 8.5, True, C_ACCENT
    AddText ws, "txtTitle", "Portfel inwestycyjny", x0, 26, 420, 32, 22, False, C_FG
    AddText ws, "txtLede", "", x0, 62, totalW - 230, 34, 10.5, False, C_MUTED, "=" & DATA_SHEET & "!$AB$16"
    AddPanel ws, "pnlStamp", x0 + totalW - 205, 14, 205, 24, C_PANEL, C_BORDER
    With ws.Shapes.AddShape(msoShapeOval, x0 + totalW - 193, 22.5, 7, 7)
        .Name = "dotStamp"
        .Fill.ForeColor.RGB = C_OK
        .Line.Visible = msoFalse
        .Placement = xlFreeFloating
    End With
    AddText ws, "txtStamp", "", x0 + totalW - 180, 14, 172, 24, 9, False, C_MUTED, "=" & DATA_SHEET & "!$AB$18"

    AddButton ws, "btnRefreshPortfolio", "AKTUALIZUJ KURSY", x0, 102, 150, "RefreshPortfolio", True
    AddButton ws, "btnOpenEntryDashboard", "DODAJ TRANSAKCJE", x0 + 160, 102, 150, "OpenTransactionEntry", False
    AddButton ws, "btnOpenAnalysis", "ANALIZA", x0 + 320, 102, 110, "GoToAnalysis", False
    AddButton ws, "btnOpenDividends", "DYWIDENDY", x0 + 440, 102, 110, "GoToDividends", False
    AddButton ws, "btnImportT212", "IMPORT TRADING 212", x0 + 560, 102, 140, "ImportTrading212", False
    AddButton ws, "btnImportXTB", "IMPORT XTB", x0 + 710, 102, 100, "ImportXTB", False

    ' 3) Kafelki z kluczowymi liczbami.
    y = 146
    cardW = (totalW - 5 * gap) / 6
    kpiLabels = Array("WartoúÊ portfela", "Zysk / strata", "Stopa zwrotu", "Wp≥aty netto", "Wolne úrodki", "Dywidendy - prognoza 12 mies.")
    kpiRefs = Array("$AB$2", "$AB$4", "$AB$5", "$AB$3", "$AB$6", "$AB$28")
    kpiSubs = Array("", "wartoúÊ minus wp≥aty netto", "zysk " & ChrW(247) & " wp≥aty netto", "wp≥aty minus wyp≥aty", "", "")
    For i = 0 To 5
        cardL = x0 + i * (cardW + gap)
        If i = 0 Then
            AddPanel ws, "kpiCard" & i, cardL, y, cardW, 90, C_HERO, C_HEROLINE
        Else
            AddPanel ws, "kpiCard" & i, cardL, y, cardW, 90, C_PANEL, C_BORDER
        End If
        AddText ws, "kpiLbl" & i, CStr(kpiLabels(i)), cardL + 14, y + 12, cardW - 28, 16, 9, False, C_MUTED
        AddText ws, "kpiVal" & i, "", cardL + 14, y + 30, cardW - 20, 30, 15, False, IIf(i = 5, C_OK, C_FG), "=" & DATA_SHEET & "!" & kpiRefs(i)
        Select Case i
            Case 0
                AddText ws, "kpiChip", "", cardL + 14, y + 64, 60, 16, 9, True, C_MUTED, "=" & DATA_SHEET & "!$AB$10"
                AddText ws, "kpiSub0", "", cardL + 76, y + 64, cardW - 90, 16, 9, False, C_MUTED, "=" & DATA_SHEET & "!$AB$17"
            Case 4
                AddText ws, "kpiSub4", "", cardL + 14, y + 64, cardW - 28, 16, 9, False, C_MUTED, "=" & DATA_SHEET & "!$AB$19"
            Case 5
                AddText ws, "kpiSub5", "", cardL + 14, y + 64, cardW - 28, 16, 9, False, C_MUTED, "=" & DATA_SHEET & "!$AB$26"
            Case Else
                AddText ws, "kpiSub" & i, CStr(kpiSubs(i)), cardL + 14, y + 64, cardW - 28, 16, 9, False, C_MUTED
        End Select
    Next i

    ' 4) Wykres wartoúci i sk≥ad portfela.
    y = y + 90 + gap + 4
    lineW = (totalW - gap) * 0.62
    compL = x0 + lineW + gap
    compW = totalW - lineW - gap

    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            Set co = ws.ChartObjects.Add(x0, y, lineW, 300)
            co.Name = "WykresWartosci"
            co.Placement = xlFreeFloating
            With co.Chart
                Do While .SeriesCollection.Count > 0
                    .SeriesCollection(1).Delete
                Loop
                Set s = .SeriesCollection.NewSeries
                s.Name = "WartoúÊ"
                s.XValues = lo.ListColumns("Data").DataBodyRange
                s.Values = lo.ListColumns("WartoúÊ portfela").DataBodyRange
                s.Format.Line.ForeColor.RGB = C_ACCENT
                s.Format.Line.Weight = 2.25
                s.MarkerStyle = xlMarkerStyleNone
                Set s = .SeriesCollection.NewSeries
                s.Name = "Wp≥aty netto"
                s.XValues = lo.ListColumns("Data").DataBodyRange
                s.Values = lo.ListColumns("Wp≥aty netto").DataBodyRange
                s.Format.Line.ForeColor.RGB = C_MUTED
                s.Format.Line.Weight = 1.5
                s.Format.Line.DashStyle = msoLineDash
                s.MarkerStyle = xlMarkerStyleNone
                .ChartType = xlLine
                .PlotVisibleOnly = False
                .DisplayBlanksAs = xlInterpolated
                .HasLegend = True
                StyleChartDark co.Chart, "WartoúÊ portfela a wp≥aty"
                On Error Resume Next
                .Axes(xlCategory).CategoryType = xlTimeScale
                .Axes(xlCategory).BaseUnit = xlDays
                .Axes(xlCategory).TickLabels.NumberFormat = "dd.mm"
                .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
                On Error GoTo 0
            End With
        End If
    End If

    AddPanel ws, "pnlSklad", compL, y, compW, 300, C_PANEL, C_BORDER
    AddText ws, "txtSklad", "Sk≥ad portfela", compL + 16, y + 12, compW - 32, 20, 11, True, C_FG
    AddText ws, "txtSkladKlasa", "WED£UG KLASY", compL + 16, y + 42, compW - 32, 14, 8, False, C_MUTED
    AddStackBar ws, "barKlasa", compL + 16, y + 60, compW - 32, wsD.Range("AE2:AE4"), _
                Array("Akcje", "ETF-y", "GotÛwka"), Array(C_ACCENT, C_ACCENT2, C_MUTED)
    AddLegendRow ws, "Akcje", compL + 16, y + 82, compW - 32, "Akcje", C_ACCENT, "=" & DATA_SHEET & "!$AE$2", "=" & DATA_SHEET & "!$AF$2"
    AddLegendRow ws, "Etf", compL + 16, y + 101, compW - 32, "ETF-y", C_ACCENT2, "=" & DATA_SHEET & "!$AE$3", "=" & DATA_SHEET & "!$AF$3"
    AddLegendRow ws, "Gotowka", compL + 16, y + 120, compW - 32, "GotÛwka", C_MUTED, "=" & DATA_SHEET & "!$AE$4", "=" & DATA_SHEET & "!$AF$4"

    AddText ws, "txtSkladWaluta", "WED£UG WALUTY NOTOWANIA", compL + 16, y + 154, compW - 32, 14, 8, False, C_MUTED
    AddStackBar ws, "barWaluta", compL + 16, y + 172, compW - 32, wsD.Range("AE7:AE10"), _
                Array("PLN", "USD", "EUR", "GBP"), Array(C_WARN, C_ACCENT, C_ACCENT2, C_OK)
    AddLegendRow ws, "PLN", compL + 16, y + 194, compW - 32, "PLN (z gotÛwkπ)", C_WARN, "=" & DATA_SHEET & "!$AE$7", "=" & DATA_SHEET & "!$AF$7"
    AddLegendRow ws, "USD", compL + 16, y + 213, compW - 32, "USD", C_ACCENT, "=" & DATA_SHEET & "!$AE$8", "=" & DATA_SHEET & "!$AF$8"
    AddLegendRow ws, "EUR", compL + 16, y + 232, compW - 32, "EUR", C_ACCENT2, "=" & DATA_SHEET & "!$AE$9", "=" & DATA_SHEET & "!$AF$9"
    AddLegendRow ws, "GBP", compL + 16, y + 251, compW - 32, "GBP", C_OK, "=" & DATA_SHEET & "!$AE$10", "=" & DATA_SHEET & "!$AF$10"

    ' 5) Zysk i strata kaødej pozycji (posortowane, kolory ustawia RefreshDashboardVisuals).
    plTop = y + 300 + gap + 4
    plH = 60 + n * 17
    Set co = ws.ChartObjects.Add(x0, plTop, totalW, plH)
    co.Name = "WykresZyskuPozycji"
    co.Placement = xlFreeFloating
    With co.Chart
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        Set s = .SeriesCollection.NewSeries
        s.Name = "Zysk / strata (PLN)"
        s.XValues = wsD.Range("AK2").Resize(n, 1)
        s.Values = wsD.Range("AI2").Resize(n, 1)
        s.Format.Fill.ForeColor.RGB = C_OK
        s.Format.Line.Visible = msoFalse
        .ChartType = xlBarClustered
        .HasLegend = False
        .ChartGroups(1).GapWidth = 45
        StyleChartDark co.Chart, "Zysk i strata na kaødej pozycji (PLN, w nawiasie % od ceny nabycia)"
        On Error Resume Next
        .Axes(xlCategory).ReversePlotOrder = True
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        .Axes(xlCategory).TickLabels.Font.Color = C_FG
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        s.HasDataLabels = True
        s.DataLabels.ShowValue = True
        s.DataLabels.NumberFormat = "+#,##0 ""z≥"";-#,##0 ""z≥"";0 ""z≥"""
        s.DataLabels.Position = xlLabelPositionOutsideEnd
        s.DataLabels.Font.Color = C_FG
        s.DataLabels.Font.Size = 8.5
        On Error GoTo 0
    End With

    ' 6) Tabela pozycji (komÛrki pod wykresami).
    yT = plTop + plH + 22
    tblRow = 1
    Do While ws.Rows(tblRow).Top < yT
        tblRow = tblRow + 1
    Loop
    ctrlRow = tblRow + 1
    hdrRow = tblRow + 3
    dataRow = tblRow + 4
    lastRow = dataRow + n + 5

    With ws.Cells(tblRow, "B")
        .Value = "Pozycje"
        .Font.Size = 12
        .Font.Bold = True
    End With
    ws.Rows(tblRow).RowHeight = 24
    ws.Cells(tblRow, "C").Formula2 = "=IFERROR(IF(ISNUMBER(H" & dataRow & "),ROWS(B" & dataRow & "#)&"" z ""&" & DATA_SHEET & "!$AB$15&"" pozycji"",""0 pozycji""),"""")"
    ws.Cells(tblRow, "C").Font.Color = C_MUTED

    ws.Cells(ctrlRow, "B").Value = "Wszystkie"
    ws.Cells(ctrlRow, "C").Value = "Szukaj:"
    ws.Cells(ctrlRow, "C").HorizontalAlignment = xlRight
    ws.Cells(ctrlRow, "C").Font.Color = C_MUTED
    ws.Range(ws.Cells(ctrlRow, "D"), ws.Cells(ctrlRow, "E")).Merge
    ws.Cells(ctrlRow, "F").Value = "fragment nazwy lub tickera; pole puste = wszystkie"
    ws.Cells(ctrlRow, "F").Font.Color = C_MUTED
    ws.Cells(ctrlRow, "F").Font.Size = 9
    With ws.Range(ws.Cells(ctrlRow, "B"), ws.Cells(ctrlRow, "B"))
        .Interior.Color = C_INPUT
        .Borders.Color = C_BORDER
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & DATA_SHEET & "!$AN$2:$AN$4"
    End With
    With ws.Range(ws.Cells(ctrlRow, "D"), ws.Cells(ctrlRow, "E"))
        .Interior.Color = C_INPUT
        .Borders.Color = C_BORDER
    End With
    ws.Rows(ctrlRow).RowHeight = 22

    ws.Range(ws.Cells(hdrRow, "B"), ws.Cells(hdrRow, "L")).Value = Array("SpÛ≥ka / ETF", "Ticker", "Grupa", "Liczba", "Cena (PLN)", _
        "Koszt (PLN)", "WartoúÊ (PLN)", "Zysk / strata", "Zwrot", "Udzia≥", "Dywidenda")
    With ws.Range(ws.Cells(hdrRow, "B"), ws.Cells(hdrRow, "L"))
        .Font.Color = C_MUTED
        .Font.Size = 9
        .Interior.Color = C_PANEL
        .HorizontalAlignment = xlRight
        .Borders(xlEdgeBottom).Color = C_BORDER
    End With
    ws.Range(ws.Cells(hdrRow, "B"), ws.Cells(hdrRow, "D")).HorizontalAlignment = xlLeft
    ws.Cells(hdrRow, "L").HorizontalAlignment = xlLeft
    ws.Rows(hdrRow).RowHeight = 24

    spill = "=LET(nazwy,Pozycje!$B$5:$B$500,grupy,Pozycje!$A$5:$A$500,dywid,Pozycje!$C$5:$C$500,liczby,Pozycje!$D$5:$D$500," & _
            "koszty,Pozycje!$F$5:$F$500,wartosci,Pozycje!$G$5:$G$500,zyski,Pozycje!$H$5:$H$500,zwroty,Pozycje!$I$5:$I$500," & _
            "tickery,Notowania!$B$7:$B$502,ceny,Notowania!$E$7:$E$502," & _
            "krotki,IFERROR(MID(tickery,FIND("":"",tickery)+1,20),tickery)," & _
            "wybor,$B$" & ctrlRow & ",szukaj,TRIM($D$" & ctrlRow & ")," & _
            "maska,(nazwy<>"""")*IF(OR(wybor="""",wybor=""Wszystkie""),1,grupy=wybor)*IF(szukaj="""",1,ISNUMBER(SEARCH(szukaj,nazwy&"" ""&krotki)))," & _
            "udzialy,IFERROR(wartosci/" & DATA_SHEET & "!$AB$2,0),dywTxt,IF(dywid=""Tak"",""Wyp≥aca"",dywid)," & _
            "IF(SUM(maska)=0,""Brak pozycji dla wybranego filtra""," & _
            "SORTBY(FILTER(HSTACK(nazwy,krotki,grupy,liczby,ceny,koszty,wartosci,zyski,zwroty,udzialy,dywTxt),maska)," & _
            "FILTER(IF(ISNUMBER(wartosci),wartosci,0),maska),-1)))"
    ws.Cells(dataRow, "B").Formula2 = spill

    With ws.Range(ws.Cells(dataRow, "B"), ws.Cells(lastRow, "L"))
        .Interior.Color = C_PANEL
        .Borders(xlInsideHorizontal).Color = C_ROWLINE
        .Borders(xlEdgeBottom).Color = C_ROWLINE
        .HorizontalAlignment = xlRight
        .RowHeight = 22
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(dataRow, "B"), ws.Cells(lastRow, "D")).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(dataRow, "L"), ws.Cells(lastRow, "L")).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(dataRow, "C"), ws.Cells(lastRow, "D")).Font.Color = C_MUTED
    ws.Range(ws.Cells(dataRow, "L"), ws.Cells(lastRow, "L")).Font.Color = C_MUTED
    ws.Range(ws.Cells(dataRow, "E"), ws.Cells(lastRow, "E")).NumberFormat = "0.0000"
    ws.Range(ws.Cells(dataRow, "F"), ws.Cells(lastRow, "H")).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(dataRow, "I"), ws.Cells(lastRow, "I")).NumberFormat = "+#,##0.00;-#,##0.00;0.00"
    ws.Range(ws.Cells(dataRow, "J"), ws.Cells(lastRow, "J")).NumberFormat = "+0.0%;-0.0%;0.0%"
    ws.Range(ws.Cells(dataRow, "K"), ws.Cells(lastRow, "K")).NumberFormat = "0.0%"

    With ws.Range(ws.Cells(dataRow, "I"), ws.Cells(lastRow, "J"))
        .FormatConditions.Delete
        .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0").Font.Color = C_OK
        .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0").Font.Color = C_BAD
    End With
    With ws.Range(ws.Cells(dataRow, "K"), ws.Cells(lastRow, "K")).FormatConditions.AddDatabar
        .BarColor.Color = C_ACCENT
        .BarFillType = xlDataBarFillSolid
        .ShowValue = True
    End With

    With ws.Cells(lastRow + 2, "B")
        .Value = "èrÛd≥o: arkusze Pozycje, Notowania i Transakcje  ∑  kursy Yahoo mogπ byÊ opÛünione  ∑  wszystkie wartoúci w PLN"
        .Font.Size = 9
        .Font.Color = C_MUTED
    End With

    RefreshDashboardVisuals
End Sub

' Kolory zaleøne od znaku (zysk / strata) - wywo≥ywane po kaødej aktualizacji.
Public Sub RefreshDashboardVisuals(Optional ByVal fromCode As Boolean = False)
    Dim ws As Worksheet, wsD As Worksheet
    Dim co As ChartObject
    Dim s As Series
    Dim vals As Variant
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    Set wsD = ThisWorkbook.Worksheets(DATA_SHEET)
    If ws Is Nothing Or wsD Is Nothing Then Exit Sub

    SetSignColor ws, "kpiVal1", wsD.Range("AB4").Value, C_FG
    SetSignColor ws, "kpiVal2", wsD.Range("AB5").Value, C_FG
    SetSignColor ws, "kpiChip", wsD.Range("AB10").Value, C_MUTED

    Set co = ws.ChartObjects("WykresZyskuPozycji")
    If Not co Is Nothing Then
        Set s = co.Chart.SeriesCollection(1)
        vals = s.Values
        For i = LBound(vals) To UBound(vals)
            If IsNumeric(vals(i)) Then
                If vals(i) < 0 Then
                    s.Points(i - LBound(vals) + 1).Format.Fill.ForeColor.RGB = C_BAD
                Else
                    s.Points(i - LBound(vals) + 1).Format.Fill.ForeColor.RGB = C_OK
                End If
            End If
        Next i
    End If
End Sub

Public Sub SetSignColor(ByVal ws As Worksheet, ByVal shapeName As String, ByVal v As Variant, ByVal neutralColor As Long)
    Dim c As Long
    c = neutralColor
    If Not IsError(v) Then
        If IsNumeric(v) And Not IsEmpty(v) And VarType(v) <> vbString Then
            If v > 0 Then c = C_OK
            If v < 0 Then c = C_BAD
        End If
    End If
    ws.Shapes(shapeName).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = c
End Sub
