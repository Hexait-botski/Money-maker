Attribute VB_Name = "Portfel"
Option Explicit

Private Const YAHOO_QUOTE_ROOT As String = "https://query1.finance.yahoo.com/v8/finance/chart/"
Public YahooQuoteLastError As String

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ---------------------------------------------------------------------------
' Modu≥ "Portfel" - ca≥a logika skoroszytu w jednym miejscu:
'  - pobieranie kursÛw z Yahoo (przycisk AKTUALIZUJ KURSY i przy otwarciu),
'  - arkusz "Nowe transakcje" i przycisk ZAPISZ TRANSAKCJE,
'  - zapis dziennej historii portfela,
'  - instalator InstallPortfolioTools.
' Transakcje sπ dopisywane do arkusza Transakcje od wiersza 30.
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
Private Const C_BG As Long = 1182986        ' #0a0d12 t≥o
Private Const C_PANEL As Long = 1971729     ' #11161e panel
Private Const C_FG As Long = 15788775       ' #e7eaf0 tekst
Private Const C_MUTED As Long = 10983052    ' #8c96a7 tekst pomocniczy
Private Const C_BORDER As Long = 3286558    ' #1e2632 linie
Private Const C_OK As Long = 10081076       ' #34d399 zysk
Private Const C_BAD As Long = 7434744       ' #f87171 strata
Private Const C_WARN As Long = 2408443      ' #fbbf24
Private Const C_ACCENT As Long = 16426336   ' #60a5fa
Private Const C_ACCENT2 As Long = 16419751  ' #a78bfa
Private Const C_TRACK As Long = 2892058     ' #1a212c
Private Const C_HERO As Long = 4006675      ' #13233d
Private Const C_HEROLINE As Long = 5190690  ' #22344f
Private Const C_ROWLINE As Long = 2629143   ' #171e28
Private Const C_INPUT As Long = 2826264     ' #18202b
Private Const THEME_FONT As String = "Segoe UI"
Public Const ENTRY_HEADER_ROW As Long = 5
Public Const ENTRY_FIRST_ROW As Long = 6
Public Const ENTRY_LAST_ROW As Long = 55

Public Const OP_BUY As String = "Kupno"
Public Const OP_SELL As String = "Sprzedaø"
Public Const OP_DEPOSIT As String = "Wp≥ata"
Public Const OP_WITHDRAW As String = "Wyp≥ata"
Public Const OP_DIVIDEND As String = "Dywidenda"

' Kolumny arkusza "Nowe transakcje".
Public Const COL_TYPE As Long = 1      ' A  Typ operacji
Public Const COL_NAME As Long = 2      ' B  Instrument / ticker
Public Const COL_UNITS As Long = 3     ' C  Liczba jednostek
Public Const COL_AMOUNT As Long = 4    ' D  Kwota w PLN (ca≥oúÊ)
Public Const COL_EXCH As Long = 5      ' E  Gie≥da (tylko nowy ticker)
Public Const COL_CLASS As Long = 6     ' F  Klasa (tylko nowy ticker)
Public Const COL_STATUS As Long = 7    ' G  Status (wype≥nia makro)
Public Const COL_LAST_INPUT As Long = 6

Private Const FIRST_POSITION_ROW As Long = 5
Private Const FIRST_QUOTE_ROW As Long = 7
Private Const QUOTE_ROW_OFFSET As Long = 2     ' Pozycje wiersz 5 = Notowania wiersz 7
Private Const FIRST_LEDGER_ROW As Long = 30
Private Const LAST_SUPPORTED_ROW As Long = 500

' Jedna przygotowana operacja.
Private Type PendingTx
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
        MsgBox "Po≥πczenie dzia≥a. AAPL: " & Format$(testPrice, "0.00") & " " & testCurrency & _
               " (" & Format$(testTime, "dd.mm.yyyy hh:mm") & ")", vbInformation, "Test Yahoo"
    Else
        MsgBox "Test Yahoo nie powiÛd≥ siÍ:" & vbCrLf & YahooQuoteLastError, _
               vbExclamation, "Test Yahoo"
    End If
End Sub

' Zwraca cenÍ w walucie podanej przez Yahoo (np. "GBp" = pensy - przelicza wywo≥ujπcy).
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
                              "; odpowiedü: " & Left$(CStr(http.ResponseText), 300)
        GoTo NoQuote
    End If

    body = CStr(http.ResponseText)
    p = YahooQuoteJsonNumber(body, "regularMarketPrice")
    c = YahooQuoteJsonString(body, "currency")
    t = YahooQuoteJsonNumber(body, "regularMarketTime")
    If IsEmpty(p) Or Len(c) = 0 Or IsEmpty(t) Then
        YahooQuoteLastError = "Serwer zwrÛci≥ odpowiedü, ktÛrej nie uda≥o siÍ odczytaÊ. Fragment: " & _
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
        YahooQuoteLastError = "B≥πd po≥πczenia: " & CStr(Err.Number) & " - " & Err.Description
    End If
    GetYahooQuote = False
End Function

Private Function YahooQuoteJsonNumber(ByVal body As String, ByVal key As String) As Variant
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

Private Function YahooQuoteJsonString(ByVal body As String, ByVal key As String) As String
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

' Uruchamia siÍ samo przy otwarciu pliku (makra muszπ byÊ w≥πczone).
' Bez okienka z komunikatem - wynik widaÊ w Notowania!A3.
Public Sub Auto_Open()
    On Error Resume Next
    UpdatePortfolioQuotes False
End Sub

' Przycisk "AKTUALIZUJ KURSY" na Dashboardzie.
Public Sub RefreshPortfolio()
    UpdatePortfolioQuotes True
End Sub

' Pobiera kursy wszystkich instrumentÛw i 3 kursy walut, potem zapisuje
' dzienny wpis w historii portfela. Zwraca liczbÍ pobranych kursÛw spÛ≥ek.
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
            Application.StatusBar = "Pobieranie kursÛw Yahoo: " & symbol & " (" & totalSymbols & ")"
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
                    ' STOCKHISTORY podaje spÛ≥ki z Londynu w tej samej jednostce co gie≥da,
                    ' wiÍc zapamiÍtujemy, czy kurs zapasowy trzeba podzieliÊ przez 100.
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
        Application.StatusBar = "Pobieranie kursÛw walut: " & fxSymbols(j)
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
                           " | kursy spÛ≥ek: " & okCount & "/" & totalSymbols
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
            MsgBox "Aktualizacja zakoÒczona: " & okCount & "/" & totalSymbols & " kursÛw.", _
                   vbInformation, "Portfel inwestycyjny"
        Else
            MsgBox "Pobrano " & okCount & "/" & totalSymbols & " kursÛw." & vbCrLf & _
                   "Bez kursu online:" & failedList & vbCrLf & _
                   "Dla nich uøyto kursu zamkniÍcia albo ceny zapasowej.", _
                   vbExclamation, "Portfel inwestycyjny"
        End If
    End If
    Exit Function

Failed:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If showMessage Then
        MsgBox "Aktualizacja nie powiod≥a siÍ: " & Err.Description, vbExclamation, "Portfel inwestycyjny"
    End If
End Function

'' Dopisuje (raz dziennie) wartoúÊ portfela i wp≥aty netto do tabeli HistoriaPortfela
' (ukryty arkusz Dane). Kolejne odúwieøenia tego samego dnia nadpisujπ dzisiejszy wpis.
Public Sub RecordPortfolioSnapshot()
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
        ' Pusty ostatni wiersz tabeli (np. úwieøo utworzonej) jest wykorzystywany.
        If Len(Trim$(CStr(lr.Range.Cells(1, 2).Value))) = 0 Then sameDay = True
    End If
    If Not sameDay Then Set lr = lo.ListRows.Add

    lr.Range.Cells(1, 1).Value = Now
    lr.Range.Cells(1, 1).NumberFormat = "dd.mm.yyyy hh:mm"
    lr.Range.Cells(1, 2).Value = Round(CDbl(portfolioValue), 2)
    lr.Range.Cells(1, 3).Value = Round(CDbl(netDeposits), 2)
Done:
End Sub

' Przycisk "DODAJ TRANSAKCJE" w arkuszach Dashboard i Transakcje.
Public Sub OpenTransactionEntry()
    With ThisWorkbook.Worksheets(ENTRY_SHEET)
        .Activate
        .Cells(ENTRY_FIRST_ROW, COL_TYPE).Select
    End With
End Sub

' Przycisk "WYCZYå∆".
Public Sub ClearNewTransactions()
    If MsgBox("WyczyúciÊ wszystkie wpisane, niezapisane wiersze?", vbQuestion + vbYesNo, _
              "Nowe transakcje") = vbYes Then
        ThisWorkbook.Worksheets(ENTRY_SHEET).Range( _
            ThisWorkbook.Worksheets(ENTRY_SHEET).Cells(ENTRY_FIRST_ROW, COL_TYPE), _
            ThisWorkbook.Worksheets(ENTRY_SHEET).Cells(ENTRY_LAST_ROW, COL_STATUS)).ClearContents
    End If
End Sub

' Przycisk "ZAPISZ TRANSAKCJE".
Public Sub SaveNewTransactions()
    Dim wsE As Worksheet, wsT As Worksheet, wsP As Worksheet, wsQ As Worksheet
    Dim txs() As PendingTx
    Dim holdings As Object
    Dim r As Long, i As Long, n As Long
    Dim errorCount As Long, newCount As Long
    Dim msg As String
    Dim ledgerRow As Long

    Set wsE = ThisWorkbook.Worksheets(ENTRY_SHEET)
    Set wsT = ThisWorkbook.Worksheets("Transakcje")
    Set wsP = ThisWorkbook.Worksheets("Pozycje")
    Set wsQ = ThisWorkbook.Worksheets("Notowania")

    For r = ENTRY_FIRST_ROW To ENTRY_LAST_ROW
        If RowHasInput(wsE, r) Then n = n + 1
    Next r
    If n = 0 Then
        MsgBox "Brak transakcji do zapisania. Wpisz je w wierszach " & ENTRY_FIRST_ROW & "-" & _
               ENTRY_LAST_ROW & ".", vbInformation, "Nowe transakcje"
        Exit Sub
    End If

    ledgerRow = NextLedgerRow(wsT)
    If ledgerRow + n - 1 > LAST_SUPPORTED_ROW Then
        MsgBox "W rejestrze transakcji brakuje wolnych wierszy (limit: wiersz " & LAST_SUPPORTED_ROW & ").", _
               vbExclamation, "Nowe transakcje"
        Exit Sub
    End If

    ReDim txs(1 To n)
    Set holdings = CreateObject("Scripting.Dictionary")
    holdings.CompareMode = vbTextCompare

    ' 1) Sprawdzenie wszystkich wierszy - nic nie jest jeszcze zapisywane.
    i = 0
    For r = ENTRY_FIRST_ROW To ENTRY_LAST_ROW
        If RowHasInput(wsE, r) Then
            i = i + 1
            txs(i).EntryRow = r
            msg = ValidateEntryRow(wsE, wsP, wsQ, r, holdings, txs(i))
            If Len(msg) > 0 Then
                wsE.Cells(r, COL_STATUS).Value = "B£•D: " & msg
                errorCount = errorCount + 1
            Else
                wsE.Cells(r, COL_STATUS).Value = "OK"
            End If
        End If
    Next r

    If errorCount > 0 Then
        MsgBox "Nic nie zosta≥o zapisane. Popraw wiersze oznaczone w kolumnie Status (" & _
               errorCount & ") i kliknij ZAPISZ ponownie.", vbExclamation, "Nowe transakcje"
        Exit Sub
    End If

    ' 2) Zapis.
    On Error GoTo Failed
    Application.ScreenUpdating = False
    For i = 1 To n
        If txs(i).IsNewInstrument Then
            If FindPositionByName(wsP, txs(i).DisplayName) = 0 Then
                RegisterInstrument wsP, wsQ, txs(i)
                newCount = newCount + 1
            End If
        End If
        ledgerRow = NextLedgerRow(wsT)
        WriteLedgerRow wsT, ledgerRow, txs(i)
    Next i

    If newCount > 0 Then
        BuildDashboard
        BuildDividendsSheet
    End If

    For i = 1 To n
        wsE.Range(wsE.Cells(txs(i).EntryRow, COL_TYPE), wsE.Cells(txs(i).EntryRow, COL_LAST_INPUT)).ClearContents
        wsE.Cells(txs(i).EntryRow, COL_STATUS).Value = "Zapisano " & Format$(Now, "dd.mm.yyyy hh:mm")
    Next i

    Application.Calculate
    Application.ScreenUpdating = True
    On Error GoTo 0

    If newCount > 0 Then
        ' Nowy instrument - od razu pobierz jego kurs (zapisze teø historiÍ portfela).
        UpdatePortfolioQuotes False
    Else
        RecordPortfolioSnapshot
        RefreshDashboardVisuals
    End If

    MsgBox "Zapisano " & n & " transakcji" & IIf(newCount > 0, " (nowe instrumenty: " & newCount & ")", "") & "." & _
           vbCrLf & "GotÛwka i zrealizowany wynik sπ liczone w arkuszu Transakcje.", _
           vbInformation, "Portfel inwestycyjny"
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Zapis przerwany: " & Err.Description & vbCrLf & _
           "Sprawdü koÒcÛwkÍ arkusza Transakcje - czÍúÊ wierszy mog≥a zostaÊ dopisana.", _
           vbCritical, "Nowe transakcje"
End Sub

' ---------------------------------------------------------------------------
' Sprawdzanie jednego wiersza. Zwraca pusty tekst, gdy wiersz jest poprawny.
' ---------------------------------------------------------------------------
Private Function ValidateEntryRow(ByVal wsE As Worksheet, ByVal wsP As Worksheet, ByVal wsQ As Worksheet, _
                                  ByVal r As Long, ByVal holdings As Object, ByRef tx As PendingTx) As String
    Dim nameInput As String, exchInput As String, classInput As String
    Dim available As Double
    Dim found As Boolean

    On Error GoTo Bad

    tx.Operation = NormalizeOperation(CStr(wsE.Cells(r, COL_TYPE).Value))
    nameInput = Trim$(CStr(wsE.Cells(r, COL_NAME).Value))
    exchInput = Trim$(CStr(wsE.Cells(r, COL_EXCH).Value))
    classInput = Trim$(CStr(wsE.Cells(r, COL_CLASS).Value))
    tx.Units = ReadNumber(wsE.Cells(r, COL_UNITS))
    tx.Amount = ReadNumber(wsE.Cells(r, COL_AMOUNT))
    tx.Fee = 0

    If Len(tx.Operation) = 0 Then
        ValidateEntryRow = "wybierz typ operacji (kolumna A)."
        Exit Function
    End If
    If tx.Amount <= 0 Then
        ValidateEntryRow = "podaj kwotÍ w PLN (kolumna D)."
        Exit Function
    End If

    ' Wp≥ata / Wyp≥ata: wystarczy kwota.
    If tx.Operation = OP_DEPOSIT Or tx.Operation = OP_WITHDRAW Then
        tx.DisplayName = vbNullString
        tx.Units = 0
        Exit Function
    End If

    If Len(nameInput) = 0 Then
        ValidateEntryRow = "wybierz instrument z listy albo wpisz ticker (kolumna B)."
        Exit Function
    End If
    found = ResolveInstrument(wsP, wsQ, nameInput, exchInput, tx)
    If Not found And holdings.Exists(UCase$(nameInput)) Then
        ' Nowy ticker kupiony wczeúniej w tej samej partii.
        tx.DisplayName = UCase$(nameInput)
        found = True
    End If

    ' Dywidenda: instrument + kwota, ktÛra wp≥ynÍ≥a na konto.
    If tx.Operation = OP_DIVIDEND Then
        tx.Units = 0
        If Not found Then ValidateEntryRow = "nie znaleziono """ & nameInput & """ w portfelu."
        Exit Function
    End If

    ' Kupno / Sprzedaø: liczba jednostek + kwota ca≥ej transakcji.
    If tx.Units <= 0 Then
        ValidateEntryRow = "podaj liczbÍ jednostek (kolumna C)."
        Exit Function
    End If
    tx.UnitPricePLN = tx.Amount / tx.Units

    If Not found Then
        If tx.Operation = OP_SELL Then
            ValidateEntryRow = "nie znaleziono """ & nameInput & """ w portfelu - nie ma czego sprzedaÊ."
            Exit Function
        End If
        If Len(exchInput) = 0 Or (classInput <> "Akcje" And classInput <> "ETF-y") Then
            ValidateEntryRow = """" & nameInput & """ to nowy ticker - wybierz gie≥dÍ (kolumna E) i klasÍ (kolumna F)."
            Exit Function
        End If
        tx.Ticker = UCase$(nameInput)
        tx.DisplayName = tx.Ticker
        tx.AssetClass = classInput
        tx.IsNewInstrument = True
        ExchangeDetails exchInput, tx.Ticker, tx.YahooSymbol, tx.CurrencyCode, tx.MarketPrefix, tx.ExchangeName
    End If

    If Not holdings.Exists(tx.DisplayName) Then
        holdings.Add tx.DisplayName, CurrentUnits(wsP, tx.DisplayName)
    End If
    available = CDbl(holdings(tx.DisplayName))
    If tx.Operation = OP_SELL Then
        If tx.Units > available + 0.00000001 Then
            ValidateEntryRow = "sprzedaø przekracza stan (posiadasz " & Format$(available, "0.0000") & ")."
            Exit Function
        End If
        holdings(tx.DisplayName) = available - tx.Units
    Else
        holdings(tx.DisplayName) = available + tx.Units
    End If
    Exit Function

Bad:
    ValidateEntryRow = Err.Description
End Function

' Szuka instrumentu: po nazwie z portfela, po tickerze + gie≥dzie albo po samym tickerze.
Private Function ResolveInstrument(ByVal wsP As Worksheet, ByVal wsQ As Worksheet, _
                                   ByVal nameInput As String, ByVal exchInput As String, _
                                   ByRef tx As PendingTx) As Boolean
    Dim posRow As Long, quoteRow As Long, r As Long, lastRow As Long
    Dim sym As String, cur As String, pre As String, exName As String
    Dim marketTicker As String, matchRow As Long, matchCount As Long

    posRow = FindPositionByName(wsP, nameInput)
    If posRow > 0 Then
        tx.DisplayName = CStr(wsP.Cells(posRow, "B").Value)
        ResolveInstrument = True
        Exit Function
    End If

    If Len(exchInput) > 0 Then
        ExchangeDetails exchInput, UCase$(nameInput), sym, cur, pre, exName
        quoteRow = FindQuoteByYahooSymbol(wsQ, sym)
        If quoteRow > 0 Then
            tx.DisplayName = CStr(wsQ.Cells(quoteRow, "A").Value)
            ResolveInstrument = True
        End If
        Exit Function
    End If

    ' Sam ticker bez gie≥dy, np. "AAPL" albo "PZU": porÛwnanie z kolumnπ B (XNAS:AAPL).
    lastRow = wsQ.Cells(wsQ.Rows.Count, "S").End(xlUp).Row
    For r = FIRST_QUOTE_ROW To lastRow
        marketTicker = CStr(wsQ.Cells(r, "B").Value)
        If InStr(marketTicker, ":") > 0 Then marketTicker = Mid$(marketTicker, InStr(marketTicker, ":") + 1)
        If StrComp(Trim$(marketTicker), nameInput, vbTextCompare) = 0 Then
            matchCount = matchCount + 1
            matchRow = r
        End If
    Next r
    If matchCount = 1 Then
        tx.DisplayName = CStr(wsQ.Cells(matchRow, "A").Value)
        ResolveInstrument = True
    ElseIf matchCount > 1 Then
        Err.Raise vbObjectError + 2111, , "ticker """ & nameInput & """ wystÍpuje na kilku gie≥dach - wybierz gie≥dÍ (kolumna C)."
    End If
End Function

Private Function NormalizeOperation(ByVal rawValue As String) As String
    Dim s As String
    s = LCase$(Trim$(rawValue))
    Select Case s
        Case "kupno", "k", "buy": NormalizeOperation = OP_BUY
        Case LCase$(OP_SELL), "sprzedaz", "s", "sell": NormalizeOperation = OP_SELL
        Case LCase$(OP_DEPOSIT), "wplata": NormalizeOperation = OP_DEPOSIT
        Case LCase$(OP_WITHDRAW), "wyplata": NormalizeOperation = OP_WITHDRAW
        Case "dywidenda", "dyw": NormalizeOperation = OP_DIVIDEND
        Case Else: NormalizeOperation = vbNullString
    End Select
End Function

Private Function RowHasInput(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim c As Long
    For c = COL_TYPE To COL_LAST_INPUT
        If Len(Trim$(CStr(ws.Cells(r, c).Value))) > 0 Then
            RowHasInput = True
            Exit Function
        End If
    Next c
End Function

' Liczba z komÛrki: pusta = 0; tekst typu "1 234,56" albo "12.5" jest zamieniany na liczbÍ.
Private Function ReadNumber(ByVal cell As Range) As Double
    Dim v As Variant
    v = cell.Value
    If IsError(v) Then Err.Raise vbObjectError + 2112, , "komÛrka " & cell.Address(False, False) & " zawiera b≥πd."
    If IsEmpty(v) Then
        ReadNumber = 0
    ElseIf VarType(v) = vbString Then
        If Len(Trim$(v)) = 0 Then
            ReadNumber = 0
        Else
            ReadNumber = ParseTransactionNumber(CStr(v))
        End If
    Else
        ReadNumber = CDbl(v)
    End If
End Function

Public Function ParseTransactionNumber(ByVal rawValue As String) As Double
    Dim s As String
    Dim decimalSeparator As String
    s = Trim$(Replace(rawValue, ChrW(160), ""))
    s = Replace(s, " ", "")
    s = Replace(s, "z≥", "", , , vbTextCompare)
    decimalSeparator = Application.International(xlDecimalSeparator)

    If InStr(s, ",") > 0 And InStr(s, ".") = 0 Then
        s = Replace(s, ",", decimalSeparator)
    ElseIf InStr(s, ".") > 0 And InStr(s, ",") = 0 Then
        s = Replace(s, ".", decimalSeparator)
    ElseIf InStr(s, ".") > 0 And InStr(s, ",") > 0 Then
        If InStrRev(s, ",") > InStrRev(s, ".") Then
            s = Replace(s, ".", "")
            s = Replace(s, ",", decimalSeparator)
        Else
            s = Replace(s, ",", "")
            s = Replace(s, ".", decimalSeparator)
        End If
    End If

    If Len(s) = 0 Or Not IsNumeric(s) Then Err.Raise vbObjectError + 2100, , "niepoprawna liczba: """ & rawValue & """."
    ParseTransactionNumber = CDbl(s)
End Function

Private Function CurrentUnits(ByVal wsP As Worksheet, ByVal displayName As String) As Double
    Dim posRow As Long
    Dim v As Variant
    posRow = FindPositionByName(wsP, displayName)
    If posRow = 0 Then Exit Function
    v = wsP.Cells(posRow, "D").Value
    If Not IsError(v) And IsNumeric(v) And Not IsEmpty(v) Then CurrentUnits = CDbl(v)
End Function

Private Sub WriteLedgerRow(ByVal wsT As Worksheet, ByVal logRow As Long, ByRef tx As PendingTx)
    With wsT
        .Cells(logRow, "A").Value = Date
        .Cells(logRow, "A").NumberFormat = "dd.mm.yyyy"
        If Len(tx.DisplayName) > 0 Then
            .Cells(logRow, "B").Value = tx.DisplayName
        Else
            .Cells(logRow, "B").ClearContents
        End If
        .Cells(logRow, "C").Value = tx.Operation
        If tx.Operation = OP_BUY Or tx.Operation = OP_SELL Then
            .Cells(logRow, "D").Value = tx.Units
        Else
            .Cells(logRow, "D").ClearContents
        End If
        .Cells(logRow, "E").Value = Round(tx.Amount, 2)
        .Cells(logRow, "F").Value = Round(tx.Fee, 2)
    End With
End Sub

Private Sub ExchangeDetails(ByVal exchangeInput As String, ByVal ticker As String, _
                            ByRef yahooSymbol As String, ByRef currencyCode As String, _
                            ByRef marketPrefix As String, ByRef exchangeName As String)
    Select Case UCase$(Trim$(exchangeInput))
        Case "GPW"
            exchangeName = "GPW": currencyCode = "PLN": marketPrefix = "XWAR:": yahooSymbol = ticker & ".WA"
        Case "NASDAQ"
            exchangeName = "NASDAQ": currencyCode = "USD": marketPrefix = "XNAS:": yahooSymbol = ticker
        Case "NYSE"
            exchangeName = "NYSE": currencyCode = "USD": marketPrefix = "XNYS:": yahooSymbol = ticker
        Case "XETRA"
            exchangeName = "Xetra": currencyCode = "EUR": marketPrefix = "XETR:": yahooSymbol = ticker & ".DE"
        Case "LSE"
            exchangeName = "LSE": currencyCode = "GBP": marketPrefix = "XLON:": yahooSymbol = ticker & ".L"
        Case Else
            Err.Raise vbObjectError + 2110, , "nieobs≥ugiwana gie≥da: " & exchangeInput & " (GPW, NASDAQ, NYSE, Xetra, LSE)."
    End Select
    ' Yahoo zapisuje klasy akcji z myúlnikiem, np. BRK.B -> BRK-B.
    If currencyCode = "USD" Then yahooSymbol = Replace(yahooSymbol, ".", "-")
End Sub

Private Function FindQuoteByYahooSymbol(ByVal ws As Worksheet, ByVal yahooSymbol As String) As Long
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, "S").End(xlUp).Row
    If lastRow < FIRST_QUOTE_ROW Then Exit Function
    For r = FIRST_QUOTE_ROW To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, "S").Value)), yahooSymbol, vbTextCompare) = 0 Then
            FindQuoteByYahooSymbol = r
            Exit Function
        End If
    Next r
End Function

Private Function FindPositionByName(ByVal ws As Worksheet, ByVal displayName As String) As Long
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If lastRow < FIRST_POSITION_ROW Then Exit Function
    For r = FIRST_POSITION_ROW To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, "B").Value)), Trim$(displayName), vbTextCompare) = 0 Then
            FindPositionByName = r
            Exit Function
        End If
    Next r
End Function

Private Function NextLedgerRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    For r = FIRST_LEDGER_ROW To LAST_SUPPORTED_ROW
        If Len(Trim$(CStr(ws.Cells(r, "A").Value))) = 0 And _
           Len(Trim$(CStr(ws.Cells(r, "B").Value))) = 0 And _
           Len(Trim$(CStr(ws.Cells(r, "C").Value))) = 0 Then
            NextLedgerRow = r
            Exit Function
        End If
    Next r
    NextLedgerRow = LAST_SUPPORTED_ROW + 1
End Function

' Dodaje nowy instrument: wiersz w Pozycje i odpowiadajπcy mu wiersz w Notowania
' (wiersz w Notowania = wiersz w Pozycje + 2, tak jak w istniejπcych danych).
Private Sub RegisterInstrument(ByVal wsP As Worksheet, ByVal wsQ As Worksheet, ByRef tx As PendingTx)
    Dim posRow As Long, quoteRow As Long
    Dim fxRate As Double

    posRow = wsP.Cells(wsP.Rows.Count, "B").End(xlUp).Row + 1
    If posRow < FIRST_POSITION_ROW + 1 Then Err.Raise vbObjectError + 2107, , "arkusz Pozycje nie ma øadnej pozycji wzorcowej."
    quoteRow = posRow + QUOTE_ROW_OFFSET
    If quoteRow > LAST_SUPPORTED_ROW Then Err.Raise vbObjectError + 2107, , "brak miejsca na nowy instrument."
    If Len(Trim$(CStr(wsQ.Cells(quoteRow, "S").Value))) > 0 Or Len(Trim$(CStr(wsQ.Cells(quoteRow, "B").Value))) > 0 Then
        Err.Raise vbObjectError + 2108, , "wiersz " & quoteRow & " w arkuszu Notowania jest zajÍty - arkusze Pozycje i Notowania nie sπ zgrane."
    End If

    ' Kopia poprzedniego wiersza przenosi formu≥y i formatowanie.
    wsP.Rows(posRow - 1).Copy Destination:=wsP.Rows(posRow)
    With wsP
        .Cells(posRow, "A").Value = tx.AssetClass
        .Cells(posRow, "B").Value = tx.DisplayName
        .Cells(posRow, "C").Value = "Nieustalone"
        .Cells(posRow, "G").Formula = "=Notowania!H" & quoteRow
    End With

    wsQ.Range("A" & (quoteRow - 1) & ":X" & (quoteRow - 1)).Copy Destination:=wsQ.Range("A" & quoteRow & ":X" & quoteRow)
    With wsQ
        .Cells(quoteRow, "A").Formula = "=Pozycje!B" & posRow
        .Cells(quoteRow, "B").Value = tx.MarketPrefix & tx.Ticker
        .Cells(quoteRow, "C").Value = tx.ExchangeName
        .Cells(quoteRow, "D").Formula = "=IF(Pozycje!D" & posRow & "="""","""",Pozycje!D" & posRow & ")"
        .Cells(quoteRow, "F").Value = tx.CurrencyCode
        .Cells(quoteRow, "I").Formula = "=Pozycje!F" & posRow
        .Cells(quoteRow, "S").Value = tx.YahooSymbol
        .Range("T" & quoteRow & ":W" & quoteRow).ClearContents
        .Cells(quoteRow, "X").Value = 1
        Select Case tx.CurrencyCode
            Case "PLN": fxRate = 1
            Case "EUR": fxRate = Val(.Range("N7").Value)
            Case "USD": fxRate = Val(.Range("N8").Value)
            Case "GBP": fxRate = Val(.Range("N9").Value)
        End Select
        ' Cena zapasowa w walucie instrumentu = cena zakupu w PLN / kurs waluty.
        If fxRate > 0 Then
            .Cells(quoteRow, "O").Value = tx.UnitPricePLN / fxRate
        Else
            .Cells(quoteRow, "O").ClearContents
        End If
    End With
End Sub

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
           "- import historii z Trading 212 (przycisk IMPORTUJ Z TRADING 212)." & vbCrLf & vbCrLf & _
           "Zapisz plik (Ctrl+S).", vbInformation, "Portfel inwestycyjny"
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Instalacja przerwana: " & Err.Description, vbExclamation, "Portfel inwestycyjny"
End Sub

Private Sub SetupEntrySheet()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ENTRY_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Transakcje"))
        ws.Name = ENTRY_SHEET
    Else
        ' Arkusz budowany od nowa (np. po wczeúniejszej wersji z innymi kolumnami).
        Do While ws.Shapes.Count > 0
            ws.Shapes(1).Delete
        Loop
        On Error Resume Next
        ws.Cells.Validation.Delete
        On Error GoTo 0
        ws.Cells.Clear
        ws.Columns.Hidden = False
    End If
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    With ws
        .Cells.Font.Name = "Arial"
        .Cells.Font.Size = 10

        .Range("A1").Value = "NOWE TRANSAKCJE"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Rows(1).RowHeight = 34
        .Range("A1").VerticalAlignment = xlCenter
        .Range("A2").Value = "Kaødy wiersz = jedna operacja. Wype≥nij øÛ≥te pola i kliknij ZAPISZ TRANSAKCJE. Data = dzisiaj, wszystkie kwoty w PLN."
        .Range("A3").Value = "Kupno / Sprzedaø: instrument, liczba jednostek i kwota ca≥ej transakcji.   Wp≥ata / Wyp≥ata: tylko kwota.   Dywidenda: instrument i kwota, ktÛra wp≥ynÍ≥a."
        .Range("A4").Value = "Gie≥dÍ i klasÍ wybierasz tylko przy pierwszym kupnie nowego tickera. Jeúli ktÛryú wiersz ma b≥πd, nic nie zostanie zapisane - opis w kolumnie Status."
        .Range("A2:A4").Font.Color = RGB(71, 85, 105)

        .Range(.Cells(ENTRY_HEADER_ROW, COL_TYPE), .Cells(ENTRY_HEADER_ROW, COL_STATUS)).Value = _
            Array("Typ operacji", "Instrument / ticker", "Liczba jednostek", "Kwota (PLN)", _
                  "Gie≥da (tylko nowy ticker)", "Klasa (tylko nowy ticker)", "Status")
        With .Range(.Cells(ENTRY_HEADER_ROW, COL_TYPE), .Cells(ENTRY_HEADER_ROW, COL_STATUS))
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(15, 23, 42)
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        .Rows(ENTRY_HEADER_ROW).RowHeight = 30

        ' Listy wyboru w ukrytych kolumnach L:N.
        .Range("L5:N5").Value = Array("Typy", "Gie≥dy", "Klasy")
        .Range("L6:L10").Value = Application.WorksheetFunction.Transpose(Array(OP_BUY, OP_SELL, OP_DEPOSIT, OP_WITHDRAW, OP_DIVIDEND))
        .Range("M6:M10").Value = Application.WorksheetFunction.Transpose(Array("GPW", "NASDAQ", "NYSE", "Xetra", "LSE"))
        .Range("N6:N7").Value = Application.WorksheetFunction.Transpose(Array("Akcje", "ETF-y"))
        .Columns("L:N").Hidden = True

        With .Range(.Cells(ENTRY_FIRST_ROW, COL_TYPE), .Cells(ENTRY_LAST_ROW, COL_LAST_INPUT))
            .Interior.Color = RGB(255, 247, 214)
            .Borders.Color = RGB(226, 232, 240)
        End With
        .Range(.Cells(ENTRY_FIRST_ROW, COL_STATUS), .Cells(ENTRY_LAST_ROW, COL_STATUS)).Font.Color = RGB(71, 85, 105)

        AddListValidation .Range(.Cells(ENTRY_FIRST_ROW, COL_TYPE), .Cells(ENTRY_LAST_ROW, COL_TYPE)), "=$L$6:$L$10", True
        AddListValidation .Range(.Cells(ENTRY_FIRST_ROW, COL_NAME), .Cells(ENTRY_LAST_ROW, COL_NAME)), "=Pozycje!$B$5:$B$500", False
        AddListValidation .Range(.Cells(ENTRY_FIRST_ROW, COL_EXCH), .Cells(ENTRY_LAST_ROW, COL_EXCH)), "=$M$6:$M$10", True
        AddListValidation .Range(.Cells(ENTRY_FIRST_ROW, COL_CLASS), .Cells(ENTRY_LAST_ROW, COL_CLASS)), "=$N$6:$N$7", True

        With .Range(.Cells(ENTRY_FIRST_ROW, COL_UNITS), .Cells(ENTRY_LAST_ROW, COL_AMOUNT)).Validation
            .Delete
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="0"
            .ErrorMessage = "Wpisz liczbÍ wiÍkszπ od zera."
        End With
        .Range(.Cells(ENTRY_FIRST_ROW, COL_UNITS), .Cells(ENTRY_LAST_ROW, COL_UNITS)).NumberFormat = "0.0000"
        .Range(.Cells(ENTRY_FIRST_ROW, COL_AMOUNT), .Cells(ENTRY_LAST_ROW, COL_AMOUNT)).NumberFormat = "#,##0.00"

        .Columns("A").ColumnWidth = 14
        .Columns("B").ColumnWidth = 30
        .Columns("C").ColumnWidth = 14
        .Columns("D").ColumnWidth = 14
        .Columns("E").ColumnWidth = 14
        .Columns("F").ColumnWidth = 14
        .Columns("G").ColumnWidth = 60
    End With

    ActiveWindow.FreezePanes = False
    ws.Range("A6").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub AddListValidation(ByVal target As Range, ByVal listFormula As String, ByVal strict As Boolean)
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listFormula
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = strict
        .ErrorMessage = "Wybierz wartoúÊ z listy."
    End With
End Sub

' Notowania: kurs zapasowy STOCKHISTORY dla spÛ≥ek z Londynu jest w pensach.
' Kolumna X przechowuje dzielnik (1 albo 100); formu≥a ceny go uwzglÍdnia.
Private Sub SetupQuotesSheet()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long

    Set ws = ThisWorkbook.Worksheets("Notowania")
    ws.Range("X6").Value = "Dzielnik STOCKHISTORY (100 = pensy)"
    ws.Range("X6").Font.Bold = True
    ws.Range("X6").WrapText = True

    lastRow = ws.Cells(ws.Rows.Count, "S").End(xlUp).Row
    For r = 7 To lastRow
        If Len(Trim$(CStr(ws.Cells(r, "S").Value))) > 0 Then
            If IsEmpty(ws.Cells(r, "X").Value) Or Not IsNumeric(ws.Cells(r, "X").Value) Then
                ws.Cells(r, "X").Value = GuessPriceDivisor(ws, r)
            End If
            ws.Cells(r, "E").Formula = "=IF(NOT(ISNUMBER(G" & r & ")),"""",IF(ISNUMBER(T" & r & "),T" & r & "*G" & r & _
                ",IF(ISNUMBER(P" & r & "),P" & r & "/IF(AND(ISNUMBER(X" & r & "),X" & r & ">0),X" & r & ",1)*G" & r & _
                ",IF(ISNUMBER(O" & r & "),O" & r & "*G" & r & ",""""))))"
            ws.Cells(r, "L").Formula = "=IF(ISNUMBER(T" & r & "),""Yahoo online"",IF(ISNUMBER(P" & r & "),""Kurs zamkniÍcia"",""Cena zapasowa""))"
        End If
    Next r

    ws.Range("A4").Value = "Kursy Yahoo mogπ byÊ opÛünione. Przy braku ceny online uøywany jest kurs zamkniÍcia (STOCKHISTORY), a potem cena zapasowa."
End Sub

Private Function GuessPriceDivisor(ByVal ws As Worksheet, ByVal r As Long) As Long
    Dim p As Variant, o As Variant
    GuessPriceDivisor = 1
    If UCase$(Trim$(CStr(ws.Cells(r, "F").Value))) <> "GBP" Then Exit Function
    p = ws.Cells(r, "P").Value
    o = ws.Cells(r, "O").Value
    If IsError(p) Or IsError(o) Then Exit Function
    If Not IsNumeric(p) Or Not IsNumeric(o) Or IsEmpty(p) Or IsEmpty(o) Then Exit Function
    If CDbl(o) <= 0 Then Exit Function
    If CDbl(p) / CDbl(o) > 50 Then GuessPriceDivisor = 100
End Function

Private Sub SetupButtons()
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
    AddButton ThisWorkbook.Worksheets(DIVIDEND_SHEET), "btnImportT212Div", "IMPORTUJ Z TRADING 212", _
              ThisWorkbook.Worksheets(DIVIDEND_SHEET).Range("E1").Left + 280, ThisWorkbook.Worksheets(DIVIDEND_SHEET).Range("B1").Top + 4, 170, "ImportTrading212", False
    AddButton ledger, "btnAddTransaction", "DODAJ TRANSAKCJE", ledger.Range("O2").Left, ledger.Range("O2").Top, 175, "OpenTransactionEntry"
    AddButton entry, "btnSaveTransactions", "ZAPISZ TRANSAKCJE", entry.Range("C1").Left, entry.Range("A1").Top + 2, 175, "SaveNewTransactions"
    AddButton entry, "btnClearTransactions", "WYCZYå∆", entry.Range("C1").Left + 185, entry.Range("A1").Top + 2, 90, "ClearNewTransactions"
End Sub

' Przycisk: g≥Ûwny (niebieski) albo pomocniczy (ciemny z obramowaniem).
Private Sub AddButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal caption As String, _
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

Private Sub SetupAnalysisSheet()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim labels As Variant
    Dim keepStarts As Object
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

Private Function PeriodTableFormula() As String
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
Private Function SetupHistoryColumns() As ListObject
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

Private Function ListColumnExists(ByVal lo As ListObject, ByVal colName As String) As Boolean
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then
            ListColumnExists = True
            Exit Function
        End If
    Next lc
End Function

Private Sub BuildAnalysisCharts(ByVal ws As Worksheet, ByVal lo As ListObject)
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

Private Sub PrepareLineChart(ByVal ch As Chart, ByVal title As String)
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

Private Sub AddLineSeries(ByVal ch As Chart, ByVal seriesName As String, ByVal lo As ListObject, _
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

Private Sub FormatChartAxes(ByVal ch As Chart, ByVal valueFormat As String)
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
Private Sub SetupDataSheet()
    Dim wsD As Worksheet
    Dim lo As ListObject
    Dim vals As Variant
    Dim nRows As Long
    Dim specs As Variant
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
    If Not lo.DataBodyRange Is Nothing Then
        lo.ListColumns(1).DataBodyRange.NumberFormat = "dd.mm.yyyy hh:mm"
        lo.ListColumns(2).DataBodyRange.NumberFormat = "#,##0.00"
        lo.ListColumns(3).DataBodyRange.NumberFormat = "#,##0.00"
    End If

    ' Kluczowe liczby: etykieta w AA, formu≥a w AB.
    specs = Array( _
        Array(2, "WartoúÊ portfela", "=SUM(Pozycje!$G$5:$G$500)+AB6", "#,##0.00 ""z≥"";-#,##0.00 ""z≥"""), _
        Array(3, "Wp≥aty netto", "=Transakcje!$R$7", "#,##0.00 ""z≥"";-#,##0.00 ""z≥"""), _
        Array(4, "Zysk / strata", "=AB2-AB3", "+#,##0.00 ""z≥"";-#,##0.00 ""z≥"";0.00 ""z≥"""), _
        Array(5, "Stopa zwrotu", "=IF(AB3=0,"""",AB4/AB3)", "+0.0%;-0.0%;0.0%"), _
        Array(6, "GotÛwka", "=Transakcje!$R$6", "#,##0.00 ""z≥"";-#,##0.00 ""z≥"""), _
        Array(7, "Udzia≥ gotÛwki", "=IF(AB2=0,"""",AB6/AB2)", "0.0%"), _
        Array(8, "Poprzedni odczyt - wartoúÊ", "=IFERROR(LET(daty," & HISTORY_TABLE & "[Data],wart," & HISTORY_TABLE & "[WartoúÊ portfela],ileW,ROWS(wart)," & _
              "pozW,IF(INT(INDEX(daty,ileW))=TODAY(),ileW-1,ileW),IF(pozW<1,"""",INDEX(wart,pozW))),"""")", "#,##0.00 ""z≥"""), _
        Array(9, "Poprzedni odczyt - data", "=IFERROR(LET(daty," & HISTORY_TABLE & "[Data],ileW,ROWS(daty)," & _
              "pozW,IF(INT(INDEX(daty,ileW))=TODAY(),ileW-1,ileW),IF(pozW<1,"""",INDEX(daty,pozW))),"""")", "dd.mm.yyyy hh:mm"), _
        Array(10, "Zmiana od odczytu", "=IFERROR(IF(AB8="""","""",AB2/AB8-1),"""")", ChrW(9650) & " 0.0%;" & ChrW(9660) & " 0.0%;0.0%"), _
        Array(11, "Kursy z", "=IFERROR(IF(COUNT(" & HISTORY_TABLE & "[Data])=0,"""",MAX(" & HISTORY_TABLE & "[Data])),"""")", "dd.mm.yyyy hh:mm"), _
        Array(12, "Najlepsza pozycja", "=IFERROR(INDEX(AH2#,1,1),"""")", "General"), _
        Array(13, "Najs≥absza pozycja", "=IFERROR(INDEX(AH2#,ROWS(AH2#),1),"""")", "General"), _
        Array(14, "Wynik najs≥abszej", "=IFERROR(INDEX(AH2#,ROWS(AH2#),2),0)", "#,##0.00"), _
        Array(15, "Liczba pozycji", "=SUMPRODUCT((Pozycje!$B$5:$B$500<>"""")*ISNUMBER(Pozycje!$H$5:$H$500))", "0"), _
        Array(16, "Zdanie podsumowania", "=""Portfel jest wart ""&FIXED(AB2,2)&"" z≥, czyli ""&IF(AB4>=0,FIXED(AB4,2)&"" z≥ ponad wp≥aty."",FIXED(-AB4,2)&"" z≥ poniøej wp≥at."")" & _
              "&IF(AB12="""","""","" NajwiÍcej zarobi≥a pozycja ""&AB12&IF(AB14<0,"", najwiÍcej traci ""&AB13&""."","".""))", "General"), _
        Array(17, "Tekst od odczytu", "=IF(ISNUMBER(AB9),""od odczytu ""&RIGHT(""0""&DAY(AB9),2)&"".""&RIGHT(""0""&MONTH(AB9),2)&"" ""&RIGHT(""0""&HOUR(AB9),2)&"":""&RIGHT(""0""&MINUTE(AB9),2),""brak wczeúniejszego odczytu"")", "General"), _
        Array(18, "Tekst kursy z", "=IF(ISNUMBER(AB11),""Kursy z ""&RIGHT(""0""&DAY(AB11),2)&"".""&RIGHT(""0""&MONTH(AB11),2)&"".""&YEAR(AB11)&"" ""&RIGHT(""0""&HOUR(AB11),2)&"":""&RIGHT(""0""&MINUTE(AB11),2),""Kursy: brak aktualizacji"")", "General"), _
        Array(19, "Tekst gotÛwki", "=IFERROR(FIXED(AB7*100,1)&""% portfela w gotÛwce"","""")", "General"), _
        Array(20, "Dywidendy netto - 12 mies.", "=IFERROR(SUMPRODUCT((INDEX(AP2#,0,1)>EDATE(TODAY(),-12))*INDEX(AP2#,0,5)),0)", "#,##0.00 ""z≥"""), _
        Array(21, "Dywidendy netto - ≥πcznie", "=IFERROR(SUM(INDEX(AP2#,0,5)),0)", "#,##0.00 ""z≥"""), _
        Array(22, "Dywidendy netto - bieøπcy rok", "=IFERROR(SUMPRODUCT((YEAR(INDEX(AP2#,0,1))=YEAR(TODAY()))*INDEX(AP2#,0,5)),0)", "#,##0.00 ""z≥"""), _
        Array(23, "Koszt pozycji wyp≥acajπcych", "=SUMIF(Pozycje!$C$5:$C$500,""Tak"",Pozycje!$F$5:$F$500)", "#,##0.00 ""z≥"""), _
        Array(24, "Stopa od kosztu (prognoza)", "=IFERROR(AB28/SUMIFS(" & DIV_FORECAST_TABLE & "[Koszt pozycji (PLN)]," & DIV_FORECAST_TABLE & "[Prognoza 12 mies. netto (PLN)],"">0""),0)", "0.00%"), _
        Array(25, "SpÛ≥ki wyp≥acajπce", "=IFERROR(COUNTIF(" & DIV_FORECAST_TABLE & "[Prognoza 12 mies. netto (PLN)],"">0""),0)", "0"), _
        Array(26, "Tekst dywidend", "=""otrzymane w 12 mies.: ""&FIXED(AB20,2)&"" z≥""", "General"), _
        Array(27, "Dywidendy - úrednio miesiÍcznie", "=AB20/12", "#,##0.00 ""z≥"""), _
        Array(28, "Prognoza dywidend 12 mies. (netto)", "=IFERROR(SUM(" & DIV_FORECAST_TABLE & "[Prognoza 12 mies. netto (PLN)]),0)", "#,##0.00 ""z≥"""), _
        Array(29, "Stopa dywidendy portfela", "=IFERROR(AB28/(AB2-AB6),0)", "0.00%"))
    wsD.Range("AA1:AB1").Value = Array("Kluczowe liczby (Dashboard)", "WartoúÊ")
    For i = LBound(specs) To UBound(specs)
        wsD.Cells(specs(i)(0), "AA").Value = specs(i)(1)
        wsD.Cells(specs(i)(0), "AB").NumberFormat = specs(i)(3)
    Next i

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
    For i = LBound(specs) To UBound(specs)
        wsD.Cells(specs(i)(0), "AB").Formula2 = specs(i)(2)
    Next i

    wsD.Columns("AA").ColumnWidth = 30
    wsD.Columns("AB").ColumnWidth = 22
    wsD.Columns("AD").ColumnWidth = 30
    wsD.Columns("AH").ColumnWidth = 30
    wsD.Columns("AK").ColumnWidth = 40
    wsD.Visible = xlSheetHidden
End Sub

Private Sub ThemeSheet(ByVal ws As Worksheet)
    With ws.Cells
        .Interior.Color = C_BG
        .Font.Name = THEME_FONT
        .Font.Size = 10
        .Font.Color = C_FG
    End With
End Sub

Private Sub ShowSheetLikeApp(ByVal ws As Worksheet, ByVal hideHeadings As Boolean)
    Dim prev As Object
    Set prev = ActiveSheet
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = Not hideHeadings
    If Not prev Is Nothing Then
        If prev.Visible = xlSheetVisible Then prev.Activate
    End If
End Sub

Private Function AddPanel(ByVal ws As Worksheet, ByVal shapeName As String, ByVal l As Double, ByVal t As Double, _
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

Private Function AddText(ByVal ws As Worksheet, ByVal shapeName As String, ByVal txt As String, _
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

Private Sub StyleChartDark(ByVal ch As Chart, ByVal title As String)
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
Private Sub AddStackBar(ByVal ws As Worksheet, ByVal chartName As String, ByVal l As Double, ByVal t As Double, _
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
        On Error GoTo 0
        .ChartArea.Format.Fill.Visible = msoTrue
        .ChartArea.Format.Fill.ForeColor.RGB = C_TRACK
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.Visible = msoFalse
        .PlotArea.InsideLeft = 0
        .PlotArea.InsideTop = 0
        .PlotArea.InsideWidth = w
        .PlotArea.InsideHeight = 14
    End With
End Sub

' Wiersz listy: kolorowy kwadrat, nazwa, wartoúÊ, udzia≥.
Private Sub AddLegendRow(ByVal ws As Worksheet, ByVal key As String, ByVal l As Double, ByVal t As Double, ByVal w As Double, _
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

Public Sub BuildDashboard()
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
    AddButton ws, "btnImportT212", "IMPORTUJ Z TRADING 212", x0 + 560, 102, 170, "ImportTrading212", False

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
Public Sub RefreshDashboardVisuals()
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

Private Sub SetSignColor(ByVal ws As Worksheet, ByVal shapeName As String, ByVal v As Variant, ByVal neutralColor As Long)
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

' ---------------------------------------------------------------------------
' ARKUSZ "DYWIDENDY" (+ dodatkowo szacunek podatku Belki)
' èrÛd≥o: wpisy typu Dywidenda w arkuszu Transakcje (lista w Dane!AP2#:
' data, spÛ≥ka, kwota, podatek, netto). Kwota z "Nowe transakcje" to netto.
' ---------------------------------------------------------------------------
Public Sub GoToDividends()
    ThisWorkbook.Worksheets(DIVIDEND_SHEET).Activate
End Sub

Public Sub BuildDividendsSheet()
    Dim ws As Worksheet, wsD As Worksheet
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

Private Function ReadStartPanel(ByVal ws As Worksheet) As Object
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
Private Sub BuildStartPanel(ByVal ws As Worksheet, ByVal keepStarts As Object)
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
        .Range("L" & START_PANEL_ROW - 1).Value = "Od kiedy masz akcje (dla automatu dywidend)"
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
    ws.Cells(r + 2, "L").Value = "Wpisz datÍ pierwszego zakupu; automat policzy dywidendy od tej daty (przycisk POBIERZ DYWIDENDY)."
    ws.Cells(r + 2, "L").Font.Size = 9
    ws.Cells(r + 2, "L").Font.Color = C_MUTED
    ws.Range(ws.Cells(START_PANEL_ROW - 1, "L"), ws.Cells(r + 2, "M")).Interior.Color = C_PANEL
    ws.Range(ws.Cells(START_PANEL_ROW + 1, "M"), ws.Cells(Application.Max(r, START_PANEL_ROW + 1), "M")).Interior.Color = C_INPUT
End Sub

' Dodatek: szacunek podatku Belki (19%) za wybrany rok.
Private Sub BuildBelkaPanel(ByVal ws As Worksheet, ByVal taxYear As Variant, ByVal rates As Variant)
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
            "doplata,IF(YEAR(dDaty)=rokP,IF(0.19*brutto-pobrany>0,0.19*brutto-pobrany,0),0),SUM(doplata)),0)"
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
            "Dywidendy z GPW sπ juø opodatkowane u ürÛd≥a. Do PIT-38 uøyj PIT-8C od brokera."
        .Range("L20:M27").Merge
        .Range("L20").WrapText = True
        .Range("L20").VerticalAlignment = xlTop
        .Range("L20").Font.Size = 9
        .Range("L20").Font.Color = C_MUTED
        .Range("L4:M27").Interior.Color = C_PANEL
    End With
End Sub

Private Sub StyleTableHeader(ByVal rng As Range)
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

Private Sub StyleTableBody(ByVal rng As Range)
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
    Dim loH As ListObject, loF As ListObject
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
    Set loF = FindTable(DIV_FORECAST_TABLE)
    If loH Is Nothing Or loF Is Nothing Then Exit Function

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
                fx = Val(CStr(wsQ.Cells(r, "G").Value))
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
                    unitsNow = Val(CStr(wsP.Cells(posRow, "D").Value))
                    costNow = Val(CStr(wsP.Cells(posRow, "F").Value))
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

    If Not loF.DataBodyRange Is Nothing Then loF.DataBodyRange.Delete
    For Each item In fRows
        loF.ListRows.Add.Range.Value = item
    Next item
    If Not loF.DataBodyRange Is Nothing Then
        loF.ListColumns(2).DataBodyRange.NumberFormat = "0.0000"
        loF.ListColumns(5).DataBodyRange.NumberFormat = "#,##0.00"
        loF.ListColumns(6).DataBodyRange.NumberFormat = "0.00%"
        loF.ListColumns(7).DataBodyRange.NumberFormat = "#,##0.00"
        loF.ListColumns(8).DataBodyRange.NumberFormat = "dd.mm.yyyy"
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

Private Function FindTable(ByVal tableName As String) As ListObject
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

Private Function FindPositionRow(ByVal wsP As Worksheet, ByVal nm As String) As Long
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
Private Function WithholdingRate(ByVal currencyCode As String) As Double
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
Private Function LastManualDividend(ByVal loH As ListObject, ByVal ledger As Variant, ByVal nm As String) As Date
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
Private Function UnitsHeldAt(ByVal nm As String, ByVal exDate As Date, ByVal initRows As Variant, _
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
Private Function EnsureTable(ByVal wsD As Worksheet, ByVal tableName As String, ByVal anchor As String, ByVal headers As Variant) As ListObject
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
Private Sub SetupDividendTables(ByVal wsD As Worksheet)
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

Private Sub EnsureColumns(ByVal lo As ListObject, ByVal headers As Variant)
    Dim h As Variant
    For Each h In headers
        If Not ListColumnExists(lo, CStr(h)) Then lo.ListColumns.Add.Name = CStr(h)
    Next h
End Sub

' Data zakupu pozycji z salda poczπtkowego - z panelu w arkuszu Dywidendy
' (kolumny L:M od wiersza START_PANEL_ROW). Brak daty = od poczπtku automatu.
Private Function StartDateFor(ByVal nm As String, ByVal fallback As Date) As Date
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

' ---------------------------------------------------------------------------
' IMPORT Z TRADING 212 (pliki CSV z historii konta)
' - dywidendy (ca≥a historia) -> tabela DywidendyHistoria, z podatkiem u ürÛd≥a,
' - wynik sprzedaøy sprzed salda poczπtkowego -> tabela WynikiSprzedazy (Belka),
' - kupno, sprzedaø, wp≥aty i wydatki kartπ PO dniu salda poczπtkowego
'   -> rejestr w arkuszu Transakcje (wczeúniejsze sπ juø w saldzie poczπtkowym).
' Kaødy wiersz ma klucz, wiÍc ponowny import tych samych plikÛw niczego nie dubluje.
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
    Dim total As Double, shares As Double, wht As Double, rate As Double
    Dim nDiv As Long, nTx As Long, nReal As Long, nSkip As Long, nNew As Long, nErr As Long
    Dim displayName As String, quoteRow As Long
    Dim tx As PendingTx
    Dim errList As String

    files = Application.GetOpenFilename("Eksport Trading 212 (*.csv),*.csv", , "Wybierz pliki CSV z Trading 212 (moøna kilka)", , True)
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
            Err.Raise vbObjectError + 2301, , "plik " & CStr(f) & " nie wyglπda na eksport z Trading 212."
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
                rate = Val(CsvField(fields, col, "Exchange rate"))
                If wht < 0 Then wht = 0
                loH.ListRows.Add.Range.Value = Array(tradeDate, displayName, Round(total, 2), DIV_T212_SOURCE, Round(wht * rate, 2), rowKey)
                keys(rowKey) = True
                nDiv = nDiv + 1

            ElseIf tradeDate < cutoff Then
                ' Przed saldem poczπtkowym: tylko wynik sprzedaøy (do szacunku podatku).
                If action = "Market sell" Or action = "Limit sell" Or action = "Stop sell" Then
                    loR.ListRows.Add.Range.Value = Array(tradeDate, displayName, Val(CsvField(fields, col, "Result")), rowKey)
                    keys(rowKey) = True
                    nReal = nReal + 1
                End If

            Else
                ' Po saldzie poczπtkowym: do rejestru transakcji.
                tx.Operation = vbNullString
                tx.DisplayName = displayName
                tx.Units = shares
                tx.Amount = Abs(total)
                tx.Fee = 0
                tx.IsNewInstrument = False
                Select Case action
                    Case "Market buy", "Limit buy", "Stop buy": tx.Operation = OP_BUY
                    Case "Market sell", "Limit sell", "Stop sell": tx.Operation = OP_SELL
                    Case "Deposit", "Spending cashback": tx.Operation = OP_DEPOSIT: tx.DisplayName = vbNullString
                    Case "Withdrawal", "Card debit": tx.Operation = OP_WITHDRAW: tx.DisplayName = vbNullString
                End Select
                If Len(tx.Operation) > 0 And tx.Amount > 0 Then
                    If (tx.Operation = OP_BUY Or tx.Operation = OP_SELL) And quoteRow = 0 Then
                        If tx.Operation = OP_SELL Then
                            errList = errList & vbCrLf & "- sprzedaø " & t212Name & " (brak w portfelu)"
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
                    j = NextLedgerRow(wsT)
                    If j > LAST_SUPPORTED_ROW Then Err.Raise vbObjectError + 2302, , "brak wolnych wierszy w rejestrze transakcji."
                    WriteLedgerRow wsT, j, tx
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
    MsgBox "Import z Trading 212 zakoÒczony." & vbCrLf & vbCrLf & _
           "Dywidendy: " & nDiv & vbCrLf & _
           "Wyniki sprzedaøy sprzed " & Format$(cutoff, "dd.mm.yyyy") & ": " & nReal & vbCrLf & _
           "Transakcje i wp≥aty od " & Format$(cutoff, "dd.mm.yyyy") & ": " & nTx & IIf(nNew > 0, " (nowe instrumenty: " & nNew & ")", "") & vbCrLf & _
           "PominiÍte (juø zaimportowane): " & nSkip & _
           IIf(nErr > 0, vbCrLf & vbCrLf & "Nie zaimportowano:" & errList, ""), vbInformation, "Portfel inwestycyjny"
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Import przerwany: " & Err.Description, vbExclamation, "Portfel inwestycyjny"
End Sub

' DzieÒ salda poczπtkowego: transakcje sprzed tej daty sπ juø w saldzie
' w arkuszu Transakcje (wiersze 7-29). Ustalany raz (pierwszy wpis historii portfela).
Private Function ImportCutoff() As Date
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

Private Sub CollectKeys(ByVal keys As Object, ByVal lo As ListObject, ByVal keyCol As Long)
    Dim lr As ListRow
    For Each lr In lo.ListRows
        If Len(CStr(lr.Range.Cells(1, keyCol).Value)) > 0 Then keys(CStr(lr.Range.Cells(1, keyCol).Value)) = True
    Next lr
End Sub

Private Sub FormatImportTables(ByVal loH As ListObject, ByVal loR As ListObject)
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

Private Function T212YahooSymbol(ByVal ticker As String, ByVal priceCur As String, ByVal isin As String) As String
    Select Case UCase$(priceCur)
        Case "GBX", "GBP": T212YahooSymbol = ticker & ".L"
        Case "EUR": T212YahooSymbol = ticker & ".DE"
        Case "PLN": T212YahooSymbol = ticker & ".WA"
        Case Else: T212YahooSymbol = Replace(ticker, ".", "-")
    End Select
End Function

Private Function T212Exchange(ByVal priceCur As String) As String
    Select Case UCase$(priceCur)
        Case "GBX", "GBP": T212Exchange = "LSE"
        Case "EUR": T212Exchange = "XETRA"
        Case "PLN": T212Exchange = "GPW"
        Case Else: T212Exchange = "NASDAQ"
    End Select
End Function

Private Function FxToPln(ByVal wsQ As Worksheet, ByVal currencyCode As String) As Double
    Dim r As Long
    FxToPln = 1
    For r = 7 To 9
        If UCase$(Left$(CStr(wsQ.Cells(r, "M").Value), 3)) = currencyCode And IsNumeric(wsQ.Cells(r, "N").Value) Then
            If Val(CStr(wsQ.Cells(r, "N").Value)) > 0 Then FxToPln = CDbl(wsQ.Cells(r, "N").Value)
        End If
    Next r
End Function

Private Function ParseT212Date(ByVal s As String) As Date
    ' "2026-02-16 08:00:31+00:00" -> data (bez godziny).
    If Len(s) >= 10 Then ParseT212Date = DateSerial(Val(Mid$(s, 1, 4)), Val(Mid$(s, 6, 2)), Val(Mid$(s, 9, 2)))
End Function

Private Function CsvField(ByVal fields As Variant, ByVal col As Object, ByVal fieldName As String) As String
    If col.Exists(fieldName) Then
        If col(fieldName) <= UBound(fields) Then CsvField = Trim$(CStr(fields(col(fieldName))))
    End If
End Function

Private Function ReadUtf8File(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8File = st.ReadText
    st.Close
End Function

' Podzia≥ wiersza CSV z obs≥ugπ pÛl w cudzys≥owach.
Private Function SplitCsvLine(ByVal line As String) As Variant
    Dim out() As String
    Dim n As Long, i As Long
    Dim ch As String, cur As String
    Dim inQuotes As Boolean
    ReDim out(0 To 0)
    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        If inQuotes Then
            If ch = """" Then
                If Mid$(line, i + 1, 1) = """" Then
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
