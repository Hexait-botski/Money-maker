Attribute VB_Name = "PortfelImport"
Option Explicit

' ---------------------------------------------------------------------------
' IMPORT Z TRADING 212 (pliki CSV z historii konta)
' - dywidendy (ca³a historia) -> tabela DywidendyHistoria, z podatkiem u Ÿród³a,
' - wynik sprzeda¿y sprzed salda pocz¹tkowego -> tabela WynikiSprzedazy (Belka),
' - kupno, sprzeda¿, wp³aty i wyp³aty (tak¿e wydatki kart¹) PO dniu salda pocz¹tkowego
'   -> rejestr w arkuszu Transakcje (wczeœniejsze s¹ ju¿ w saldzie pocz¹tkowym).
' Ka¿dy wiersz ma klucz, wiêc ponowny import tych samych plików niczego nie dubluje.
' ---------------------------------------------------------------------------
Public Sub ImportTrading212()
    Dim files As Variant
    Dim f As Variant
    Dim wsT As Worksheet, wsP As Worksheet, wsQ As Worksheet, wsD As Worksheet
    Dim loH As ListObject, loR As ListObject
    Dim keys As Object
    Dim lines As Variant, header As Variant, fields As Variant
    Dim col As Object
    Dim i As Long, j As Long
    Dim action As String, rowKey As String, ticker As String, t212Name As String, isin As String
    Dim priceCur As String, totalCur As String
    Dim tradeDate As Date, cutoff As Date
    Dim total As Double, shares As Double, wht As Double, divFxRate As Double
    Dim nDiv As Long, nTx As Long, nReal As Long, nSkip As Long, nNew As Long, nErr As Long
    Dim displayName As String, quoteRow As Long
    Dim tx As PendingTx
    Dim errList As String

    files = Application.GetOpenFilename("Eksport Trading 212 (*.csv),*.csv", , "Wybierz pliki CSV z Trading 212 (mo¿na kilka)", , True)
    If Not IsArray(files) Then Exit Sub

    On Error GoTo Failed
    Set wsT = ThisWorkbook.Worksheets("Transakcje")
    Set wsP = ThisWorkbook.Worksheets("Pozycje")
    Set wsQ = ThisWorkbook.Worksheets("Notowania")
    Set wsD = ThisWorkbook.Worksheets(DATA_SHEET)
    Set loH = FindTable(DIV_HISTORY_TABLE)
    Set loR = FindTable(REALIZED_TABLE)
    If loH Is Nothing Or loR Is Nothing Then Err.Raise vbObjectError + 2300, , "uruchom najpierw InstallPortfolioTools."
    cutoff = ImportCutoff()

    Set keys = CreateObject("Scripting.Dictionary")
    CollectKeys keys, loH, 6
    CollectKeys keys, loR, 4
    For i = FIRST_LEDGER_ROW To LAST_SUPPORTED_ROW
        If Len(CStr(wsT.Cells(i, "S").Value)) > 0 Then keys(CStr(wsT.Cells(i, "S").Value)) = True
    Next i

    Application.ScreenUpdating = False
    For Each f In files
        lines = Split(Replace(ReadUtf8File(CStr(f)), vbCrLf, vbLf), vbLf)
        If UBound(lines) < 1 Then GoTo NextFile
        header = SplitCsvLine(CStr(lines(0)))
        Set col = CreateObject("Scripting.Dictionary")
        For j = LBound(header) To UBound(header)
            col(CStr(header(j))) = j
        Next j
        If Not col.Exists("Action") Or Not col.Exists("Time (UTC)") Or Not col.Exists("Total") Then
            Err.Raise vbObjectError + 2301, , "plik " & CStr(f) & " nie wygl¹da na eksport z Trading 212."
        End If

        For i = 1 To UBound(lines)
            If Len(Trim$(CStr(lines(i)))) = 0 Then GoTo NextLine
            fields = SplitCsvLine(CStr(lines(i)))
            action = CsvField(fields, col, "Action")
            tradeDate = ParseT212Date(CsvField(fields, col, "Time (UTC)"))
            ticker = CsvField(fields, col, "Ticker")
            t212Name = CsvField(fields, col, "Name")
            isin = CsvField(fields, col, "ISIN")
            priceCur = CsvField(fields, col, "Currency (Price / share)")
            totalCur = CsvField(fields, col, "Currency (Total)")
            total = Val(CsvField(fields, col, "Total"))
            shares = Val(CsvField(fields, col, "No. of shares"))
            If totalCur = "EUR" Then total = total * FxToPln(wsQ, "EUR")
            If totalCur = "USD" Then total = total * FxToPln(wsQ, "USD")
            If totalCur = "GBP" Then total = total * FxToPln(wsQ, "GBP")
            rowKey = "T212|" & action & "|" & CsvField(fields, col, "Time (UTC)") & "|" & ticker & "|" & CsvField(fields, col, "ID") & "|" & CsvField(fields, col, "Total")
            If keys.Exists(rowKey) Then
                nSkip = nSkip + 1
                GoTo NextLine
            End If

            displayName = vbNullString
            quoteRow = 0
            If Len(ticker) > 0 Then
                quoteRow = FindQuoteByYahooSymbol(wsQ, T212YahooSymbol(ticker, priceCur, isin))
                If quoteRow > 0 Then displayName = CStr(wsQ.Cells(quoteRow, "A").Value) Else displayName = t212Name
            End If

            If Left$(action, 8) = "Dividend" Then
                wht = Val(CsvField(fields, col, "Withholding tax"))
                divFxRate = Val(CsvField(fields, col, "Exchange rate"))
                If wht < 0 Then wht = 0
                loH.ListRows.Add.Range.Value = Array(tradeDate, displayName, Round(total, 2), DIV_T212_SOURCE, Round(wht * divFxRate, 2), rowKey)
                keys(rowKey) = True
                nDiv = nDiv + 1

            ElseIf tradeDate < cutoff Then
                ' Przed saldem pocz¹tkowym: tylko wynik sprzeda¿y (do szacunku podatku).
                If action = "Market sell" Or action = "Limit sell" Or action = "Stop sell" Then
                    loR.ListRows.Add.Range.Value = Array(tradeDate, displayName, Val(CsvField(fields, col, "Result")), rowKey)
                    keys(rowKey) = True
                    nReal = nReal + 1
                End If

            Else
                ' Po saldzie pocz¹tkowym: do rejestru transakcji.
                tx.Operation = vbNullString
                tx.DisplayName = displayName
                tx.Units = shares
                tx.Amount = Abs(total)
                tx.Fee = 0
                tx.IsNewInstrument = False
                Select Case action
                    Case "Market buy", "Limit buy", "Stop buy": tx.Operation = OP_BUY
                    Case "Market sell", "Limit sell", "Stop sell": tx.Operation = OP_SELL
                    ' Wydatki kart¹ = wyp³ata z portfela, cashback = wp³ata (pomniejsza wyp³aty).
                    Case "Deposit", "Spending cashback": tx.Operation = OP_DEPOSIT: tx.DisplayName = vbNullString
                    Case "Withdrawal", "Card debit": tx.Operation = OP_WITHDRAW: tx.DisplayName = vbNullString
                End Select
                If Len(tx.Operation) > 0 And tx.Amount > 0 Then
                    If (tx.Operation = OP_BUY Or tx.Operation = OP_SELL) And quoteRow = 0 Then
                        If tx.Operation = OP_SELL Then
                            errList = errList & vbCrLf & "- sprzeda¿ " & t212Name & " (brak w portfelu)"
                            nErr = nErr + 1
                            GoTo NextLine
                        End If
                        tx.Ticker = UCase$(ticker)
                        tx.DisplayName = t212Name
                        tx.AssetClass = IIf(Left$(isin, 2) = "IE" Or Left$(isin, 2) = "LU", "ETF-y", "Akcje")
                        tx.UnitPricePLN = tx.Amount / IIf(shares > 0, shares, 1)
                        ExchangeDetails T212Exchange(priceCur), tx.Ticker, tx.YahooSymbol, tx.CurrencyCode, tx.MarketPrefix, tx.ExchangeName
                        If FindPositionByName(wsP, tx.DisplayName) = 0 Then
                            RegisterInstrument wsP, wsQ, tx
                            nNew = nNew + 1
                        End If
                    End If
                    j = MatchManualLedgerRow(wsT, tx)
                    If j = 0 Then
                        j = NextLedgerRow(wsT)
                        If j > LAST_SUPPORTED_ROW Then Err.Raise vbObjectError + 2302, , "brak wolnych wierszy w rejestrze transakcji."
                        WriteLedgerRow wsT, j, tx
                    End If
                    wsT.Cells(j, "A").Value = tradeDate
                    wsT.Cells(j, "S").Value = rowKey
                    keys(rowKey) = True
                    nTx = nTx + 1
                End If
            End If
NextLine:
        Next i
NextFile:
    Next f

    FormatImportTables loH, loR
    Application.Calculate
    If nNew > 0 Then BuildDashboard
    BuildDividendsSheet
    SyncDividends True
    RecordPortfolioSnapshot
    RefreshDashboardVisuals
    Application.ScreenUpdating = True
    MsgBox "Import z Trading 212 zakoñczony." & vbCrLf & vbCrLf & _
           "Dywidendy: " & nDiv & vbCrLf & _
           "Wyniki sprzeda¿y sprzed " & Format$(cutoff, "dd.mm.yyyy") & ": " & nReal & vbCrLf & _
           "Transakcje i wp³aty od " & Format$(cutoff, "dd.mm.yyyy") & ": " & nTx & IIf(nNew > 0, " (nowe instrumenty: " & nNew & ")", "") & vbCrLf & _
           "Pominiête (ju¿ zaimportowane): " & nSkip & _
           IIf(nErr > 0, vbCrLf & vbCrLf & "Nie zaimportowano:" & errList, ""), vbInformation, "Portfel inwestycyjny"
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Import przerwany: " & Err.Description, vbExclamation, "Portfel inwestycyjny"
End Sub

' Dzieñ salda pocz¹tkowego: transakcje sprzed tej daty s¹ ju¿ w saldzie
' w arkuszu Transakcje (wiersze 7-29). Ustalany raz (pierwszy wpis historii portfela).
Public Function ImportCutoff() As Date
    Dim wsD As Worksheet
    Dim lo As ListObject
    Set wsD = ThisWorkbook.Worksheets(DATA_SHEET)
    If Not IsDate(wsD.Range("AB32").Value) Then
        wsD.Range("AB32").Value = Date
        Set lo = GetHistoryTable()
        If Not lo Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then
                If IsDate(lo.ListRows(1).Range.Cells(1, 1).Value) Then wsD.Range("AB32").Value = Int(CDbl(CDate(lo.ListRows(1).Range.Cells(1, 1).Value)))
            End If
        End If
        wsD.Range("AB32").NumberFormat = "dd.mm.yyyy"
    End If
    ImportCutoff = CDate(wsD.Range("AB32").Value)
End Function

Public Sub CollectKeys(ByVal keys As Object, ByVal lo As ListObject, ByVal keyCol As Long)
    Dim lr As ListRow
    For Each lr In lo.ListRows
        If Len(CStr(lr.Range.Cells(1, keyCol).Value)) > 0 Then keys(CStr(lr.Range.Cells(1, keyCol).Value)) = True
    Next lr
End Sub

Public Sub FormatImportTables(ByVal loH As ListObject, ByVal loR As ListObject)
    If Not loH.DataBodyRange Is Nothing Then
        loH.ListColumns(1).DataBodyRange.NumberFormat = "dd.mm.yyyy"
        loH.ListColumns(3).DataBodyRange.NumberFormat = "#,##0.00"
        loH.ListColumns(5).DataBodyRange.NumberFormat = "#,##0.00"
        With loH.Sort
            .SortFields.Clear
            .SortFields.Add Key:=loH.ListColumns(1).DataBodyRange, Order:=xlDescending
            .Header = xlYes
            .Apply
        End With
    End If
    If Not loR.DataBodyRange Is Nothing Then
        loR.ListColumns(1).DataBodyRange.NumberFormat = "dd.mm.yyyy"
        loR.ListColumns(3).DataBodyRange.NumberFormat = "#,##0.00"
    End If
End Sub

Public Function T212YahooSymbol(ByVal ticker As String, ByVal priceCur As String, ByVal isin As String) As String
    Select Case UCase$(priceCur)
        Case "GBX", "GBP": T212YahooSymbol = ticker & ".L"
        Case "EUR": T212YahooSymbol = ticker & ".DE"
        Case "PLN": T212YahooSymbol = ticker & ".WA"
        Case Else: T212YahooSymbol = Replace(ticker, ".", "-")
    End Select
End Function

Public Function T212Exchange(ByVal priceCur As String) As String
    Select Case UCase$(priceCur)
        Case "GBX", "GBP": T212Exchange = "LSE"
        Case "EUR": T212Exchange = "XETRA"
        Case "PLN": T212Exchange = "GPW"
        Case Else: T212Exchange = "NASDAQ"
    End Select
End Function

Public Function FxToPln(ByVal wsQ As Worksheet, ByVal currencyCode As String) As Double
    Dim r As Long
    FxToPln = 1
    For r = 7 To 9
        If UCase$(Left$(CStr(wsQ.Cells(r, "M").Value), 3)) = currencyCode And IsNumeric(wsQ.Cells(r, "N").Value) Then
            If NumVal(wsQ.Cells(r, "N").Value) > 0 Then FxToPln = CDbl(wsQ.Cells(r, "N").Value)
        End If
    Next r
End Function

Public Function ParseT212Date(ByVal s As String) As Date
    ' "2026-02-16 08:00:31+00:00" -> data (bez godziny).
    If Len(s) >= 10 Then ParseT212Date = DateSerial(Val(Mid$(s, 1, 4)), Val(Mid$(s, 6, 2)), Val(Mid$(s, 9, 2)))
End Function

Public Function CsvField(ByVal fields As Variant, ByVal col As Object, ByVal fieldName As String) As String
    If col.Exists(fieldName) Then
        If col(fieldName) <= UBound(fields) Then CsvField = Trim$(CStr(fields(col(fieldName))))
    End If
End Function

Public Function ReadUtf8File(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8File = st.ReadText
    st.Close
End Function

' Podzia³ wiersza CSV z obs³ug¹ pól w cudzys³owach.
Public Function SplitCsvLine(ByVal csvLine As String) As Variant
    Dim out() As String
    Dim n As Long, i As Long
    Dim ch As String, cur As String
    Dim inQuotes As Boolean
    ReDim out(0 To 0)
    For i = 1 To Len(csvLine)
        ch = Mid$(csvLine, i, 1)
        If inQuotes Then
            If ch = """" Then
                If Mid$(csvLine, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQuotes = False
                End If
            Else
                cur = cur & ch
            End If
        ElseIf ch = """" Then
            inQuotes = True
        ElseIf ch = "," Then
            out(n) = cur
            n = n + 1
            ReDim Preserve out(0 To n)
            cur = vbNullString
        Else
            cur = cur & ch
        End If
    Next i
    out(n) = cur
    SplitCsvLine = out
End Function

' ---------------------------------------------------------------------------
' IMPORT Z XTB (pliki .xlsx z historii rachunków: zwyk³y, IKE, IKZE)
' - dywidendy -> DywidendyHistoria (Ÿród³o "XTB" albo "XTB IKE"/"XTB IKZE"),
' - zamkniête pozycje z rachunku zwyk³ego sprzed salda -> WynikiSprzedazy,
' - kupno, sprzeda¿, wp³aty i wyp³aty zewnêtrzne po dniu salda -> rejestr.
' Przelewy miêdzy rachunkami XTB (IKE deposit / IKE return) s¹ pomijane.
' ---------------------------------------------------------------------------
Public Sub ImportXTB()
    Dim files As Variant, f As Variant
    Dim wb As Workbook, wsC As Worksheet, wsX As Worksheet
    Dim wsT As Worksheet, wsP As Worksheet, wsQ As Worksheet
    Dim loH As ListObject, loR As ListObject
    Dim keys As Object, whtMap As Object, colIx As Object
    Dim data As Variant
    Dim r As Long, hdrRow As Long, lastRow As Long, c As Long, j As Long
    Dim opType As String, xtbTicker As String, instrName As String, comment As String, product As String
    Dim amount As Double, units As Double, opTime As Date, tradeDate As Date, cutoff As Date
    Dim rowKey As String, displayName As String, quoteRow As Long, sourceName As String
    Dim nDiv As Long, nTx As Long, nReal As Long, nSkip As Long, nNew As Long, nMatched As Long
    Dim tx As PendingTx

    files = Application.GetOpenFilename("Eksport XTB (*.xlsx),*.xlsx", , "Wybierz pliki XLSX z XTB (rachunek, IKE, IKZE)", , True)
    If Not IsArray(files) Then Exit Sub

    On Error GoTo Failed
    Set wsT = ThisWorkbook.Worksheets("Transakcje")
    Set wsP = ThisWorkbook.Worksheets("Pozycje")
    Set wsQ = ThisWorkbook.Worksheets("Notowania")
    Set loH = FindTable(DIV_HISTORY_TABLE)
    Set loR = FindTable(REALIZED_TABLE)
    If loH Is Nothing Or loR Is Nothing Then Err.Raise vbObjectError + 2300, , "uruchom najpierw InstallPortfolioTools."
    cutoff = ImportCutoff()

    Set keys = CreateObject("Scripting.Dictionary")
    CollectKeys keys, loH, 6
    CollectKeys keys, loR, 4
    For r = FIRST_LEDGER_ROW To LAST_SUPPORTED_ROW
        If Len(CStr(wsT.Cells(r, "S").Value)) > 0 Then keys(CStr(wsT.Cells(r, "S").Value)) = True
    Next r

    Application.ScreenUpdating = False
    For Each f In files
        Set wb = Workbooks.Open(CStr(f), ReadOnly:=True, UpdateLinks:=False)

        ' ---- Operacje gotówkowe ----
        Set wsC = Nothing
        On Error Resume Next
        Set wsC = wb.Worksheets("Cash Operations")
        On Error GoTo Failed
        If Not wsC Is Nothing Then
            data = wsC.UsedRange.Value
            hdrRow = 0
            For r = 1 To UBound(data, 1)
                If CStr(data(r, 1)) = "Type" Then hdrRow = r: Exit For
            Next r
            If hdrRow > 0 Then
                Set colIx = HeaderIndex(data, hdrRow)
                ' Podatek u Ÿród³a od dywidend (wiersze z tickerem i s³owem "tax").
                Set whtMap = CreateObject("Scripting.Dictionary")
                For r = hdrRow + 1 To UBound(data, 1)
                    opType = XField(data, r, colIx, "Type")
                    xtbTicker = XField(data, r, colIx, "Ticker")
                    If Len(xtbTicker) > 0 And InStr(1, opType, "tax", vbTextCompare) > 0 And IsDate(data(r, colIx("Time"))) Then
                        rowKey = xtbTicker & "|" & Format$(Int(CDate(data(r, colIx("Time")))), "yyyy-mm-dd")
                        whtMap(rowKey) = whtMap(rowKey) + Abs(NumVal(data(r, colIx("Amount"))))
                    End If
                Next r

                For r = hdrRow + 1 To UBound(data, 1)
                    opType = XField(data, r, colIx, "Type")
                    If Len(opType) = 0 Or opType = "Total" Then GoTo NextOp
                    If Not IsDate(data(r, colIx("Time"))) Then GoTo NextOp
                    opTime = CDate(data(r, colIx("Time")))
                    tradeDate = Int(opTime)
                    xtbTicker = XField(data, r, colIx, "Ticker")
                    instrName = XField(data, r, colIx, "Instrument")
                    comment = XField(data, r, colIx, "Comment")
                    product = XField(data, r, colIx, "Product")
                    amount = NumVal(data(r, colIx("Amount")))
                    rowKey = "XTB|" & XField(data, r, colIx, "ID")
                    If keys.Exists(rowKey) Then
                        nSkip = nSkip + 1
                        GoTo NextOp
                    End If
                    sourceName = "XTB" & IIf(product = "IKE" Or product = "IKZE", " " & product, "")

                    displayName = vbNullString
                    quoteRow = 0
                    If Len(xtbTicker) > 0 Then
                        quoteRow = FindQuoteByYahooSymbol(wsQ, XtbYahooSymbol(xtbTicker))
                        If quoteRow > 0 Then displayName = CStr(wsQ.Cells(quoteRow, "A").Value) Else displayName = instrName
                    End If

                    tx.Operation = vbNullString
                    tx.IsNewInstrument = False
                    tx.Fee = 0
                    Select Case LCase$(opType)
                        Case "dividend", "divident"
                            loH.ListRows.Add.Range.Value = Array(tradeDate, displayName, Round(amount, 2), sourceName, _
                                Round(whtMap(xtbTicker & "|" & Format$(tradeDate, "yyyy-mm-dd")), 2), rowKey)
                            keys(rowKey) = True
                            nDiv = nDiv + 1
                            GoTo NextOp
                        Case "stock purchase", "stocks/etf purchase", "etf purchase"
                            tx.Operation = OP_BUY
                        Case "stock sell", "stock sale", "stocks/etf sale", "etf sale"
                            tx.Operation = OP_SELL
                        Case "deposit"
                            tx.Operation = OP_DEPOSIT
                        Case "withdrawal"
                            tx.Operation = OP_WITHDRAW
                    End Select
                    If Len(tx.Operation) = 0 Or tradeDate < cutoff Then GoTo NextOp

                    tx.Amount = Abs(amount)
                    tx.DisplayName = displayName
                    tx.Units = 0
                    If tx.Operation = OP_BUY Or tx.Operation = OP_SELL Then
                        tx.Units = XtbUnits(comment)
                        If tx.Units <= 0 Then GoTo NextOp
                        If quoteRow = 0 Then
                            If tx.Operation = OP_SELL Then GoTo NextOp
                            tx.Ticker = UCase$(Split(xtbTicker, ".")(0))
                            tx.DisplayName = instrName
                            tx.AssetClass = IIf(InStr(1, XField(data, r, colIx, "Category"), "ETF", vbTextCompare) > 0, "ETF-y", "Akcje")
                            tx.UnitPricePLN = tx.Amount / tx.Units
                            ExchangeDetails XtbExchange(xtbTicker), tx.Ticker, tx.YahooSymbol, tx.CurrencyCode, tx.MarketPrefix, tx.ExchangeName
                            If FindPositionByName(wsP, tx.DisplayName) = 0 Then
                                RegisterInstrument wsP, wsQ, tx
                                nNew = nNew + 1
                            End If
                        End If
                    Else
                        tx.DisplayName = vbNullString
                    End If

                    j = MatchManualLedgerRow(wsT, tx)
                    If j > 0 Then
                        nMatched = nMatched + 1
                    Else
                        j = NextLedgerRow(wsT)
                        If j > LAST_SUPPORTED_ROW Then Err.Raise vbObjectError + 2302, , "brak wolnych wierszy w rejestrze transakcji."
                        WriteLedgerRow wsT, j, tx
                        nTx = nTx + 1
                    End If
                    wsT.Cells(j, "A").Value = tradeDate
                    wsT.Cells(j, "S").Value = rowKey
                    keys(rowKey) = True
NextOp:
                Next r
            End If
        End If

        ' ---- Zamkniête pozycje (wynik do podatku; tylko rachunek zwyk³y) ----
        Set wsX = Nothing
        On Error Resume Next
        Set wsX = wb.Worksheets("Closed Positions")
        On Error GoTo Failed
        If Not wsX Is Nothing Then
            data = wsX.UsedRange.Value
            hdrRow = 0
            For r = 1 To UBound(data, 1)
                If CStr(data(r, 1)) = "Instrument" Then hdrRow = r: Exit For
            Next r
            If hdrRow > 0 Then
                Set colIx = HeaderIndex(data, hdrRow)
                For r = hdrRow + 1 To UBound(data, 1)
                    product = XField(data, r, colIx, "Product")
                    If Len(XField(data, r, colIx, "Ticker")) = 0 Or product = "IKE" Or product = "IKZE" Then GoTo NextClosed
                    If Not IsDate(data(r, colIx("Close Time (UTC)"))) Then GoTo NextClosed
                    tradeDate = Int(CDate(data(r, colIx("Close Time (UTC)"))))
                    If tradeDate >= cutoff Then GoTo NextClosed
                    rowKey = "XTB|closed|" & XField(data, r, colIx, "Ticker") & "|" & XField(data, r, colIx, "Close Time (UTC)") & "|" & XField(data, r, colIx, "Volume")
                    If keys.Exists(rowKey) Then
                        nSkip = nSkip + 1
                        GoTo NextClosed
                    End If
                    quoteRow = FindQuoteByYahooSymbol(wsQ, XtbYahooSymbol(XField(data, r, colIx, "Ticker")))
                    If quoteRow > 0 Then displayName = CStr(wsQ.Cells(quoteRow, "A").Value) Else displayName = XField(data, r, colIx, "Instrument")
                    loR.ListRows.Add.Range.Value = Array(tradeDate, displayName, NumVal(data(r, colIx("Profit/Loss"))), rowKey)
                    keys(rowKey) = True
                    nReal = nReal + 1
NextClosed:
                Next r
            End If
        End If
        wb.Close SaveChanges:=False
        Set wb = Nothing
    Next f

    FormatImportTables loH, loR
    Application.Calculate
    If nNew > 0 Then BuildDashboard
    BuildDividendsSheet
    SyncDividends True
    RecordPortfolioSnapshot
    RefreshDashboardVisuals
    Application.ScreenUpdating = True
    MsgBox "Import z XTB zakoñczony." & vbCrLf & vbCrLf & _
           "Dywidendy: " & nDiv & vbCrLf & _
           "Wyniki sprzeda¿y (rachunek zwyk³y) sprzed " & Format$(cutoff, "dd.mm.yyyy") & ": " & nReal & vbCrLf & _
           "Nowe wpisy w rejestrze od " & Format$(cutoff, "dd.mm.yyyy") & ": " & nTx & IIf(nNew > 0, " (nowe instrumenty: " & nNew & ")", "") & vbCrLf & _
           "Rozpoznane wpisy rêczne (bez dublowania): " & nMatched & vbCrLf & _
           "Pominiête (ju¿ zaimportowane): " & nSkip, vbInformation, "Portfel inwestycyjny"
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    MsgBox "Import przerwany: " & Err.Description, vbExclamation, "Portfel inwestycyjny"
End Sub

Public Function HeaderIndex(ByVal data As Variant, ByVal hdrRow As Long) As Object
    Dim d As Object
    Dim c As Long
    Set d = CreateObject("Scripting.Dictionary")
    For c = 1 To UBound(data, 2)
        If Len(CStr(data(hdrRow, c))) > 0 Then d(CStr(data(hdrRow, c))) = c
    Next c
    Set HeaderIndex = d
End Function

Public Function XField(ByVal data As Variant, ByVal r As Long, ByVal colIx As Object, ByVal fieldName As String) As String
    If colIx.Exists(fieldName) Then
        If Not IsError(data(r, colIx(fieldName))) Then XField = Trim$(CStr(data(r, colIx(fieldName))))
    End If
End Function

' "OPEN BUY NVDA.US 0.1354 @ 159.63", "OPEN BUY 3/3.2266 @ 61.98" -> liczba jednostek.
Public Function XtbUnits(ByVal comment As String) As Double
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(OPEN|CLOSE)\s+(BUY|SELL)\s+(?:[A-Z0-9._-]+\s+)?([0-9]+(?:\.[0-9]+)?)(?:/[0-9.]+)?\s*@"
    re.IgnoreCase = True
    If re.Test(comment) Then
        Set m = re.Execute(comment)(0)
        XtbUnits = Val(m.SubMatches(2))
    End If
End Function

Public Function XtbYahooSymbol(ByVal xtbTicker As String) As String
    Dim parts As Variant
    parts = Split(UCase$(xtbTicker), ".")
    If UBound(parts) < 1 Then
        XtbYahooSymbol = UCase$(xtbTicker)
        Exit Function
    End If
    Select Case parts(1)
        Case "PL": XtbYahooSymbol = parts(0) & ".WA"
        Case "US": XtbYahooSymbol = Replace(parts(0), "_", "-")
        Case "UK": XtbYahooSymbol = parts(0) & ".L"
        Case "DE": XtbYahooSymbol = parts(0) & ".DE"
        Case Else: XtbYahooSymbol = parts(0) & "." & parts(1)
    End Select
End Function

Public Function XtbExchange(ByVal xtbTicker As String) As String
    Select Case UCase$(Mid$(xtbTicker, InStrRev(xtbTicker, ".") + 1))
        Case "PL": XtbExchange = "GPW"
        Case "UK": XtbExchange = "LSE"
        Case "DE": XtbExchange = "XETRA"
        Case Else: XtbExchange = "NASDAQ"
    End Select
End Function

' Wpis dodany wczeœniej rêcznie (bez klucza importu) z t¹ sam¹ operacj¹,
' instrumentem i liczb¹ jednostek / kwot¹ - ¿eby import go nie zdublowa³.
Public Function MatchManualLedgerRow(ByVal wsT As Worksheet, ByRef tx As PendingTx) As Long
    Dim r As Long
    Dim e As Double
    For r = FIRST_LEDGER_ROW To LAST_SUPPORTED_ROW
        If Len(Trim$(CStr(wsT.Cells(r, "C").Value))) = 0 Then Exit Function
        If Len(CStr(wsT.Cells(r, "S").Value)) = 0 And CStr(wsT.Cells(r, "C").Value) = tx.Operation Then
            If StrComp(Trim$(CStr(wsT.Cells(r, "B").Value)), Trim$(tx.DisplayName), vbTextCompare) = 0 Then
                e = NumVal(wsT.Cells(r, "E").Value)
                If Abs(e - tx.Amount) <= 0.02 * tx.Amount + 0.05 Then
                    If tx.Operation = OP_BUY Or tx.Operation = OP_SELL Then
                        If Abs(NumVal(wsT.Cells(r, "D").Value) - tx.Units) < 0.0001 Then
                            MatchManualLedgerRow = r
                            Exit Function
                        End If
                    Else
                        MatchManualLedgerRow = r
                        Exit Function
                    End If
                End If
            End If
        End If
    Next r
End Function
