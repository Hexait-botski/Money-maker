Attribute VB_Name = "PortfelKursy"
Option Explicit

Public Const YAHOO_QUOTE_ROOT As String = "https://query1.finance.yahoo.com/v8/finance/chart/"
Public YahooQuoteLastError As String

#If VBA7 Then
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ---------------------------------------------------------------------------
' Modu³ "Portfel" - ca³a logika skoroszytu w jednym miejscu:
'  - pobieranie kursów z Yahoo (przycisk AKTUALIZUJ KURSY i przy otwarciu),
'  - arkusz "Nowe transakcje" i przycisk ZAPISZ TRANSAKCJE,
'  - zapis dziennej historii portfela,
'  - instalator InstallPortfolioTools.
' Transakcje s¹ dopisywane do arkusza Transakcje od wiersza 30.
' Data = dzisiaj, prowizja = 0. Wszystkie kwoty w PLN.
' ---------------------------------------------------------------------------

Public Const ENTRY_SHEET As String = "Nowe transakcje"
Public Const ANALYSIS_SHEET As String = "Analiza"
Public Const DASH_SHEET As String = "Dashboard"
Public Const DATA_SHEET As String = "Dane"
Public Const HISTORY_TABLE As String = "HistoriaPortfela"
Public Const DIVIDEND_SHEET As String = "Dywidendy"
Public Const DIV_HISTORY_TABLE As String = "DywidendyHistoria"
Public Const START_PANEL_ROW As Long = 30
Public Const DIV_FORECAST_TABLE As String = "DywidendyPrognoza"
Public Const DIV_AUTO_SOURCE As String = "automatycznie (szacunek)"
Public Const DIV_MANUAL_SOURCE As String = "otrzymana"
Public Const DIV_T212_SOURCE As String = "Trading 212"
Public Const REALIZED_TABLE As String = "WynikiSprzedazy"

' Kolory motywu (ciemny).
Public Const C_BG As Long = 1182986        ' #0a0d12 t³o
Public Const C_PANEL As Long = 1971729     ' #11161e panel
Public Const C_FG As Long = 15788775       ' #e7eaf0 tekst
Public Const C_MUTED As Long = 10983052    ' #8c96a7 tekst pomocniczy
Public Const C_BORDER As Long = 3286558    ' #1e2632 linie
Public Const C_OK As Long = 10081076       ' #34d399 zysk
Public Const C_BAD As Long = 7434744       ' #f87171 strata
Public Const C_WARN As Long = 2408443      ' #fbbf24
Public Const C_ACCENT As Long = 16426336   ' #60a5fa
Public Const C_ACCENT2 As Long = 16419751  ' #a78bfa
Public Const C_TRACK As Long = 2892058     ' #1a212c
Public Const C_HERO As Long = 4006675      ' #13233d
Public Const C_HEROLINE As Long = 5190690  ' #22344f
Public Const C_ROWLINE As Long = 2629143   ' #171e28
Public Const C_INPUT As Long = 2826264     ' #18202b
Public Const THEME_FONT As String = "Segoe UI"
Public Const ENTRY_HEADER_ROW As Long = 5
Public Const ENTRY_FIRST_ROW As Long = 6
Public Const ENTRY_LAST_ROW As Long = 55

Public Const OP_BUY As String = "Kupno"
Public Const OP_SELL As String = "Sprzeda¿"
Public Const OP_DEPOSIT As String = "Wp³ata"
Public Const OP_WITHDRAW As String = "Wyp³ata"
Public Const OP_DIVIDEND As String = "Dywidenda"

' Kolumny arkusza "Nowe transakcje".
Public Const COL_TYPE As Long = 1      ' A  Typ operacji
Public Const COL_NAME As Long = 2      ' B  Instrument / ticker
Public Const COL_UNITS As Long = 3     ' C  Liczba jednostek
Public Const COL_AMOUNT As Long = 4    ' D  Kwota w PLN (ca³oœæ)
Public Const COL_EXCH As Long = 5      ' E  Gie³da (tylko nowy ticker)
Public Const COL_CLASS As Long = 6     ' F  Klasa (tylko nowy ticker)
Public Const COL_STATUS As Long = 7    ' G  Status (wype³nia makro)
Public Const COL_LAST_INPUT As Long = 6

Public Const FIRST_POSITION_ROW As Long = 5
Public Const FIRST_QUOTE_ROW As Long = 7
Public Const QUOTE_ROW_OFFSET As Long = 2     ' Pozycje wiersz 5 = Notowania wiersz 7
Public Const FIRST_LEDGER_ROW As Long = 30
Public Const LAST_SUPPORTED_ROW As Long = 500

' Jedna przygotowana operacja.
Public Type PendingTx
    EntryRow As Long
    Operation As String
    DisplayName As String
    Units As Double
    Amount As Double
    Fee As Double
    IsNewInstrument As Boolean
    Ticker As String
    AssetClass As String
    YahooSymbol As String
    ExchangeName As String
    CurrencyCode As String
    MarketPrefix As String
    UnitPricePLN As Double
End Type

Public Sub TestYahooConnection()
    Dim testPrice As Double
    Dim testCurrency As String
    Dim testTime As Date

    If GetYahooQuote("AAPL", testPrice, testCurrency, testTime) Then
        MsgBox "Po³¹czenie dzia³a. AAPL: " & Format$(testPrice, "0.00") & " " & testCurrency & _
               " (" & Format$(testTime, "dd.mm.yyyy hh:mm") & ")", vbInformation, "Test Yahoo"
    Else
        MsgBox "Test Yahoo nie powiód³ siê:" & vbCrLf & YahooQuoteLastError, _
               vbExclamation, "Test Yahoo"
    End If
End Sub

' Zwraca cenê w walucie podanej przez Yahoo (np. "GBp" = pensy - przelicza wywo³uj¹cy).
Public Function GetYahooQuote(ByVal symbol As String, ByRef price As Double, _
                              ByRef quoteCurrency As String, ByRef quoteTime As Date) As Boolean
    Dim http As Object
    Dim body As String
    Dim p As Variant
    Dim c As String
    Dim t As Variant
    Dim urlSymbol As String

    YahooQuoteLastError = vbNullString
    On Error GoTo NoQuote
    urlSymbol = Replace(symbol, "=", "%3D")
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", YAHOO_QUOTE_ROOT & urlSymbol & "?range=1d&interval=1d", False
    http.SetRequestHeader "User-Agent", "Mozilla/5.0"
    http.SetRequestHeader "Accept", "application/json"
    http.SetTimeouts 5000, 5000, 8000, 8000
    http.Send
    If http.Status <> 200 Then
        YahooQuoteLastError = "HTTP " & CStr(http.Status) & " " & CStr(http.StatusText) & _
                              "; odpowiedŸ: " & Left$(CStr(http.ResponseText), 300)
        GoTo NoQuote
    End If

    body = CStr(http.ResponseText)
    p = YahooQuoteJsonNumber(body, "regularMarketPrice")
    c = YahooQuoteJsonString(body, "currency")
    t = YahooQuoteJsonNumber(body, "regularMarketTime")
    If IsEmpty(p) Or Len(c) = 0 Or IsEmpty(t) Then
        YahooQuoteLastError = "Serwer zwróci³ odpowiedŸ, której nie uda³o siê odczytaæ. Fragment: " & _
                              Left$(body, 300)
        GoTo NoQuote
    End If

    price = CDbl(p)
    quoteCurrency = c
    quoteTime = DateAdd("s", CDbl(t), DateSerial(1970, 1, 1))
    GetYahooQuote = (price > 0)
    Exit Function

NoQuote:
    If Len(YahooQuoteLastError) = 0 Then
        YahooQuoteLastError = "B³¹d po³¹czenia: " & CStr(Err.Number) & " - " & Err.Description
    End If
    GetYahooQuote = False
End Function

Public Function YahooQuoteJsonNumber(ByVal body As String, ByVal key As String) As Variant
    Dim re As Object
    Dim matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = """" & key & """" & "\s*:\s*(-?[0-9]+(\.[0-9]+)?)"
    re.Global = False
    If re.Test(body) Then
        Set matches = re.Execute(body)
        YahooQuoteJsonNumber = Val(matches(0).SubMatches(0))
    Else
        YahooQuoteJsonNumber = Empty
    End If
End Function

Public Function YahooQuoteJsonString(ByVal body As String, ByVal key As String) As String
    Dim re As Object
    Dim matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = """" & key & """" & "\s*:\s*""([^""]*)"""
    re.Global = False
    If re.Test(body) Then
        Set matches = re.Execute(body)
        YahooQuoteJsonString = matches(0).SubMatches(0)
    Else
        YahooQuoteJsonString = vbNullString
    End If
End Function

' Uruchamia siê samo przy otwarciu pliku (makra musz¹ byæ w³¹czone).
' Bez okienka z komunikatem - wynik widaæ w Notowania!A3.
Public Sub Auto_Open()
    On Error Resume Next
    UpdatePortfolioQuotes False
End Sub

' Przycisk "AKTUALIZUJ KURSY" na Dashboardzie.
Public Sub RefreshPortfolio()
    UpdatePortfolioQuotes True
End Sub

' Pobiera kursy wszystkich instrumentów i 3 kursy walut, potem zapisuje
' dzienny wpis w historii portfela. Zwraca liczbê pobranych kursów spó³ek.
Public Function UpdatePortfolioQuotes(ByVal showMessage As Boolean) As Long
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim symbol As String
    Dim expectedCurrency As String
    Dim quoteCurrency As String
    Dim price As Double
    Dim stamp As Date
    Dim okCount As Long
    Dim failCount As Long
    Dim failedList As String
    Dim fxSymbols As Variant
    Dim fxRows As Variant
    Dim fxPrice As Double
    Dim fxCurrency As String
    Dim fxStamp As Date
    Dim j As Long
    Dim lastQuoteRow As Long
    Dim totalSymbols As Long
    Dim inPence As Boolean

    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("Notowania")
    Application.ScreenUpdating = False

    lastQuoteRow = ws.Cells(ws.Rows.Count, "S").End(xlUp).Row
    If lastQuoteRow < 7 Then lastQuoteRow = 6

    For rowNum = 7 To lastQuoteRow
        symbol = Trim$(CStr(ws.Cells(rowNum, "S").Value))
        expectedCurrency = Trim$(CStr(ws.Cells(rowNum, "F").Value))
        ws.Cells(rowNum, "T").ClearContents
        ws.Cells(rowNum, "U").ClearContents

        If Len(symbol) > 0 Then
            totalSymbols = totalSymbols + 1
            Application.StatusBar = "Pobieranie kursów Yahoo: " & symbol & " (" & totalSymbols & ")"
            If GetYahooQuote(symbol, price, quoteCurrency, stamp) Then
                inPence = (quoteCurrency = "GBp" Or quoteCurrency = "GBX" Or quoteCurrency = "GBPX")
                If inPence Then
                    price = price / 100
                    quoteCurrency = "GBP"
                End If

                If StrComp(quoteCurrency, expectedCurrency, vbTextCompare) = 0 Then
                    ws.Cells(rowNum, "T").Value = price
                    ws.Cells(rowNum, "U").Value = stamp
                    ws.Cells(rowNum, "U").NumberFormat = "dd.mm.yyyy hh:mm"
                    ' STOCKHISTORY podaje spó³ki z Londynu w tej samej jednostce co gie³da,
                    ' wiêc zapamiêtujemy, czy kurs zapasowy trzeba podzieliæ przez 100.
                    If quoteCurrency = "GBP" Then ws.Cells(rowNum, "X").Value = IIf(inPence, 100, 1)
                    okCount = okCount + 1
                Else
                    failCount = failCount + 1
                    failedList = failedList & " " & symbol
                End If
            Else
                failCount = failCount + 1
                failedList = failedList & " " & symbol
            End If
            Sleep 180
        End If
    Next rowNum

    fxSymbols = Array("EURPLN=X", "USDPLN=X", "GBPPLN=X")
    fxRows = Array(7, 8, 9)
    For j = LBound(fxSymbols) To UBound(fxSymbols)
        ws.Cells(fxRows(j), "V").ClearContents
        ws.Cells(fxRows(j), "W").ClearContents
        Application.StatusBar = "Pobieranie kursów walut: " & fxSymbols(j)
        If GetYahooQuote(CStr(fxSymbols(j)), fxPrice, fxCurrency, fxStamp) Then
            If StrComp(fxCurrency, "PLN", vbTextCompare) = 0 Then
                ws.Cells(fxRows(j), "V").Value = fxPrice
                ws.Cells(fxRows(j), "W").Value = fxStamp
                ws.Cells(fxRows(j), "W").NumberFormat = "dd.mm.yyyy hh:mm"
            End If
        End If
        Sleep 180
    Next j

    Application.Calculate
    ws.Range("A3").Value = "Ostatnie pobranie Yahoo: " & Format$(Now, "dd.mm.yyyy hh:mm:ss") & _
                           " | kursy spó³ek: " & okCount & "/" & totalSymbols
    If okCount > 0 Then
        SyncDividends False
        Application.Calculate
        RecordPortfolioSnapshot
    End If
    RefreshDashboardVisuals

    Application.StatusBar = False
    Application.ScreenUpdating = True
    UpdatePortfolioQuotes = okCount

    If showMessage Then
        If failCount = 0 Then
            MsgBox "Aktualizacja zakoñczona: " & okCount & "/" & totalSymbols & " kursów.", _
                   vbInformation, "Portfel inwestycyjny"
        Else
            MsgBox "Pobrano " & okCount & "/" & totalSymbols & " kursów." & vbCrLf & _
                   "Bez kursu online:" & failedList & vbCrLf & _
                   "Dla nich u¿yto kursu zamkniêcia albo ceny zapasowej.", _
                   vbExclamation, "Portfel inwestycyjny"
        End If
    End If
    Exit Function

Failed:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If showMessage Then
        MsgBox "Aktualizacja nie powiod³a siê: " & Err.Description, vbExclamation, "Portfel inwestycyjny"
    End If
End Function

'' Dopisuje (raz dziennie) wartoœæ portfela i wp³aty netto do tabeli HistoriaPortfela
' (ukryty arkusz Dane). Kolejne odœwie¿enia tego samego dnia nadpisuj¹ dzisiejszy wpis.
Public Sub RecordPortfolioSnapshot(Optional ByVal fromCode As Boolean = False)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lr As ListRow
    Dim portfolioValue As Variant
    Dim netDeposits As Variant
    Dim lastDate As Variant
    Dim sameDay As Boolean

    On Error GoTo Done
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)
    Set lo = GetHistoryTable()
    If lo Is Nothing Then Exit Sub
    Application.Calculate
    portfolioValue = ws.Range("AB2").Value
    netDeposits = ws.Range("AB3").Value
    If IsError(portfolioValue) Or IsError(netDeposits) Then Exit Sub
    If Not IsNumeric(portfolioValue) Or Not IsNumeric(netDeposits) Then Exit Sub
    If CDbl(portfolioValue) <= 0 Then Exit Sub

    sameDay = False
    If lo.ListRows.Count > 0 Then
        Set lr = lo.ListRows(lo.ListRows.Count)
        lastDate = lr.Range.Cells(1, 1).Value
        If IsDate(lastDate) Then sameDay = (Int(CDbl(CDate(lastDate))) = CLng(Date))
        ' Pusty ostatni wiersz tabeli (np. œwie¿o utworzonej) jest wykorzystywany.
        If Len(Trim$(CStr(lr.Range.Cells(1, 2).Value))) = 0 Then sameDay = True
    End If
    If Not sameDay Then Set lr = lo.ListRows.Add

    lr.Range.Cells(1, 1).Value = Now
    lr.Range.Cells(1, 1).NumberFormat = "dd.mm.yyyy hh:mm"
    lr.Range.Cells(1, 2).Value = Round(CDbl(portfolioValue), 2)
    lr.Range.Cells(1, 3).Value = Round(CDbl(netDeposits), 2)
Done:
End Sub

' Liczba z komórki niezale¿nie od ustawieñ regionalnych (Val() ucina przy przecinku).
Public Function NumVal(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) And VarType(v) <> vbString Then
        NumVal = CDbl(v)
    Else
        NumVal = Val(Replace(Replace(CStr(v), " ", ""), ",", "."))
    End If
End Function
