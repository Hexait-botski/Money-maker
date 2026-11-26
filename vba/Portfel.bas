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
    If okCount > 0 Then RecordPortfolioSnapshot

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

' Dopisuje (raz dziennie) wartoœæ portfela i wp³aty netto do tabeli HistoriaPortfela
' na Dashboardzie. Kolejne odœwie¿enia tego samego dnia nadpisuj¹ dzisiejszy wpis.
Public Sub RecordPortfolioSnapshot()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lr As ListRow
    Dim portfolioValue As Variant
    Dim netDeposits As Variant
    Dim lastDate As Variant
    Dim sameDay As Boolean

    On Error GoTo Done
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    Set lo = ws.ListObjects("HistoriaPortfela")
    portfolioValue = ws.Range("A6").Value
    netDeposits = ws.Range("D6").Value
    If IsError(portfolioValue) Or IsError(netDeposits) Then Exit Sub
    If Not IsNumeric(portfolioValue) Or Not IsNumeric(netDeposits) Then Exit Sub
    If CDbl(portfolioValue) <= 0 Then Exit Sub

    sameDay = False
    If lo.ListRows.Count > 0 Then
        Set lr = lo.ListRows(lo.ListRows.Count)
        lastDate = lr.Range.Cells(1, 1).Value
        If IsDate(lastDate) Then sameDay = (Int(CDbl(CDate(lastDate))) = CLng(Date))
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

' Przycisk "WYCZYŒÆ".
Public Sub ClearNewTransactions()
    If MsgBox("Wyczyœciæ wszystkie wpisane, niezapisane wiersze?", vbQuestion + vbYesNo, _
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
                wsE.Cells(r, COL_STATUS).Value = "B£¥D: " & msg
                errorCount = errorCount + 1
            Else
                wsE.Cells(r, COL_STATUS).Value = "OK"
            End If
        End If
    Next r

    If errorCount > 0 Then
        MsgBox "Nic nie zosta³o zapisane. Popraw wiersze oznaczone w kolumnie Status (" & _
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

    If newCount > 0 Then ApplyDashboardFormulas

    For i = 1 To n
        wsE.Range(wsE.Cells(txs(i).EntryRow, COL_TYPE), wsE.Cells(txs(i).EntryRow, COL_LAST_INPUT)).ClearContents
        wsE.Cells(txs(i).EntryRow, COL_STATUS).Value = "Zapisano " & Format$(Now, "dd.mm.yyyy hh:mm")
    Next i

    Application.Calculate
    Application.ScreenUpdating = True
    On Error GoTo 0

    If newCount > 0 Then
        ' Nowy instrument - od razu pobierz jego kurs (zapisze te¿ historiê portfela).
        UpdatePortfolioQuotes False
    Else
        RecordPortfolioSnapshot
    End If

    MsgBox "Zapisano " & n & " transakcji" & IIf(newCount > 0, " (nowe instrumenty: " & newCount & ")", "") & "." & _
           vbCrLf & "Gotówka i zrealizowany wynik s¹ liczone w arkuszu Transakcje.", _
           vbInformation, "Portfel inwestycyjny"
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Zapis przerwany: " & Err.Description & vbCrLf & _
           "SprawdŸ koñcówkê arkusza Transakcje - czêœæ wierszy mog³a zostaæ dopisana.", _
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
        ValidateEntryRow = "podaj kwotê w PLN (kolumna D)."
        Exit Function
    End If

    ' Wp³ata / Wyp³ata: wystarczy kwota.
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
        ' Nowy ticker kupiony wczeœniej w tej samej partii.
        tx.DisplayName = UCase$(nameInput)
        found = True
    End If

    ' Dywidenda: instrument + kwota, która wp³ynê³a na konto.
    If tx.Operation = OP_DIVIDEND Then
        tx.Units = 0
        If Not found Then ValidateEntryRow = "nie znaleziono """ & nameInput & """ w portfelu."
        Exit Function
    End If

    ' Kupno / Sprzeda¿: liczba jednostek + kwota ca³ej transakcji.
    If tx.Units <= 0 Then
        ValidateEntryRow = "podaj liczbê jednostek (kolumna C)."
        Exit Function
    End If
    tx.UnitPricePLN = tx.Amount / tx.Units

    If Not found Then
        If tx.Operation = OP_SELL Then
            ValidateEntryRow = "nie znaleziono """ & nameInput & """ w portfelu - nie ma czego sprzedaæ."
            Exit Function
        End If
        If Len(exchInput) = 0 Or (classInput <> "Akcje" And classInput <> "ETF-y") Then
            ValidateEntryRow = """" & nameInput & """ to nowy ticker - wybierz gie³dê (kolumna E) i klasê (kolumna F)."
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
            ValidateEntryRow = "sprzeda¿ przekracza stan (posiadasz " & Format$(available, "0.0000") & ")."
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

' Szuka instrumentu: po nazwie z portfela, po tickerze + gie³dzie albo po samym tickerze.
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

    ' Sam ticker bez gie³dy, np. "AAPL" albo "PZU": porównanie z kolumn¹ B (XNAS:AAPL).
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
        Err.Raise vbObjectError + 2111, , "ticker """ & nameInput & """ wystêpuje na kilku gie³dach - wybierz gie³dê (kolumna C)."
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

' Liczba z komórki: pusta = 0; tekst typu "1 234,56" albo "12.5" jest zamieniany na liczbê.
Private Function ReadNumber(ByVal cell As Range) As Double
    Dim v As Variant
    v = cell.Value
    If IsError(v) Then Err.Raise vbObjectError + 2112, , "komórka " & cell.Address(False, False) & " zawiera b³¹d."
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
    s = Replace(s, "z³", "", , , vbTextCompare)
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
            Err.Raise vbObjectError + 2110, , "nieobs³ugiwana gie³da: " & exchangeInput & " (GPW, NASDAQ, NYSE, Xetra, LSE)."
    End Select
    ' Yahoo zapisuje klasy akcji z myœlnikiem, np. BRK.B -> BRK-B.
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

' Dodaje nowy instrument: wiersz w Pozycje i odpowiadaj¹cy mu wiersz w Notowania
' (wiersz w Notowania = wiersz w Pozycje + 2, tak jak w istniej¹cych danych).
Private Sub RegisterInstrument(ByVal wsP As Worksheet, ByVal wsQ As Worksheet, ByRef tx As PendingTx)
    Dim posRow As Long, quoteRow As Long
    Dim fxRate As Double

    posRow = wsP.Cells(wsP.Rows.Count, "B").End(xlUp).Row + 1
    If posRow < FIRST_POSITION_ROW + 1 Then Err.Raise vbObjectError + 2107, , "arkusz Pozycje nie ma ¿adnej pozycji wzorcowej."
    quoteRow = posRow + QUOTE_ROW_OFFSET
    If quoteRow > LAST_SUPPORTED_ROW Then Err.Raise vbObjectError + 2107, , "brak miejsca na nowy instrument."
    If Len(Trim$(CStr(wsQ.Cells(quoteRow, "S").Value))) > 0 Or Len(Trim$(CStr(wsQ.Cells(quoteRow, "B").Value))) > 0 Then
        Err.Raise vbObjectError + 2108, , "wiersz " & quoteRow & " w arkuszu Notowania jest zajêty - arkusze Pozycje i Notowania nie s¹ zgrane."
    End If

    ' Kopia poprzedniego wiersza przenosi formu³y i formatowanie.
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
' Formu³y Dashboardu zale¿ne od liczby pozycji. Wywo³ywane przez instalator
' i po dodaniu nowego instrumentu.
' ---------------------------------------------------------------------------
Public Sub ApplyDashboardFormulas()
    Dim ws As Worksheet, wsP As Worksheet
    Dim r As Long, posRow As Long, lastPos As Long
    Dim helperRow As Long, oldHelperRow As Long
    Dim stockRows() As Long, stockCount As Long
    Dim a As Long, b As Long, tmp As Long
    Dim co As ChartObject

    Set ws = ThisWorkbook.Worksheets("Dashboard")
    Set wsP = ThisWorkbook.Worksheets("Pozycje")

    ' Podsumowania liczone z ca³ego zakresu pozycji (do wiersza 500).
    ws.Range("A6").Formula = "=SUM(Pozycje!G5:G500)+C8"
    ws.Range("B13").Formula = "=SUMIF(Pozycje!$A$5:$A$500,A13,Pozycje!$G$5:$G$500)"
    ws.Range("B14").Formula = "=SUMIF(Pozycje!$A$5:$A$500,A14,Pozycje!$G$5:$G$500)"
    ws.Range("E13").Formula = "=SUMIF(Pozycje!$A$5:$A$500,D13,Pozycje!$H$5:$H$500)"
    ws.Range("E14").Formula = "=SUMIF(Pozycje!$A$5:$A$500,D14,Pozycje!$H$5:$H$500)"
    ws.Range("F13").Formula = "=IFERROR(E13/SUMIF(Pozycje!$A$5:$A$500,D13,Pozycje!$F$5:$F$500),"""")"
    ws.Range("F14").Formula = "=IFERROR(E14/SUMIF(Pozycje!$A$5:$A$500,D14,Pozycje!$F$5:$F$500),"""")"

    ' Klucz sortowania w ukrytym arkuszu Pozycje (wartoœæ + drobny dodatek,
    ' ¿eby dwie pozycje o tej samej wartoœci nie dawa³y duplikatu).
    wsP.Range("J4").Value = "Klucz sortowania"
    wsP.Range("J5:J500").Formula = "=IF(OR(B5="""",NOT(ISNUMBER(G5))),"""",G5+ROW()/1000000000)"

    ' Szeœæ najwiêkszych pozycji - zawsze posortowane wed³ug aktualnej wartoœci.
    For r = 54 To 59
        ws.Cells(r, "A").Formula = "=IFERROR(INDEX(Pozycje!$B$5:$B$500,MATCH(LARGE(Pozycje!$J$5:$J$500,ROWS($A$54:A" & r & ")),Pozycje!$J$5:$J$500,0)),"""")"
        ws.Cells(r, "B").Formula = "=IF($A" & r & "="""","""",INDEX(Pozycje!$G$5:$G$500,MATCH($A" & r & ",Pozycje!$B$5:$B$500,0)))"
        ws.Cells(r, "C").Formula = "=IF($A" & r & "="""","""",INDEX(Pozycje!$H$5:$H$500,MATCH($A" & r & ",Pozycje!$B$5:$B$500,0)))"
        ws.Cells(r, "D").Formula = "=IF($A" & r & "="""","""",INDEX(Pozycje!$I$5:$I$500,MATCH($A" & r & ",Pozycje!$B$5:$B$500,0)))"
        ws.Cells(r, "E").Formula = "=IFERROR(B" & r & "/$A$6,"""")"
        ws.Cells(r, "F").Formula = "=IFERROR(IF(INDEX(Pozycje!$C$5:$C$500,MATCH($A" & r & ",Pozycje!$B$5:$B$500,0))=""Tak"",""Wyp³aca"",INDEX(Pozycje!$C$5:$C$500,MATCH($A" & r & ",Pozycje!$B$5:$B$500,0))),"""")"
    Next r

    ' Dane pomocnicze wykresu zysków/strat spó³ek (G40:I...), alfabetycznie.
    lastPos = wsP.Cells(wsP.Rows.Count, "B").End(xlUp).Row
    ReDim stockRows(1 To IIf(lastPos >= FIRST_POSITION_ROW, lastPos - FIRST_POSITION_ROW + 1, 1))
    For posRow = FIRST_POSITION_ROW To lastPos
        If CStr(wsP.Cells(posRow, "A").Value) = "Akcje" And Len(Trim$(CStr(wsP.Cells(posRow, "B").Value))) > 0 Then
            stockCount = stockCount + 1
            stockRows(stockCount) = posRow
        End If
    Next posRow
    For a = 1 To stockCount - 1
        For b = a + 1 To stockCount
            If StrComp(CStr(wsP.Cells(stockRows(b), "B").Value), CStr(wsP.Cells(stockRows(a), "B").Value), vbTextCompare) < 0 Then
                tmp = stockRows(a): stockRows(a) = stockRows(b): stockRows(b) = tmp
            End If
        Next b
    Next a

    oldHelperRow = Application.WorksheetFunction.Max( _
        ws.Cells(ws.Rows.Count, "G").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "H").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "I").End(xlUp).Row)
    If oldHelperRow < 39 Then oldHelperRow = 39

    helperRow = 39
    For a = 1 To stockCount
        helperRow = helperRow + 1
        If helperRow > 59 Then ws.Range("G59:I59").Copy Destination:=ws.Range("G" & helperRow & ":I" & helperRow)
        ws.Cells(helperRow, "G").Formula = "=Pozycje!B" & stockRows(a)
        ws.Cells(helperRow, "H").Formula = "=MAX(Pozycje!H" & stockRows(a) & ",0)"
        ws.Cells(helperRow, "I").Formula = "=MIN(Pozycje!H" & stockRows(a) & ",0)"
    Next a
    ' Usuniêcie starych, zdublowanych wierszy pomocniczych (np. G60:H62).
    If oldHelperRow > helperRow Then ws.Range("G" & (helperRow + 1) & ":I" & oldHelperRow).ClearContents

    If helperRow >= 40 Then
        On Error Resume Next
        Set co = ws.ChartObjects("WykresZyskuPozycji")
        On Error GoTo 0
        If Not co Is Nothing Then
            With co.Chart
                If .SeriesCollection.Count >= 1 Then
                    .SeriesCollection(1).XValues = ws.Range("G40:G" & helperRow)
                    .SeriesCollection(1).Values = ws.Range("H40:H" & helperRow)
                End If
                If .SeriesCollection.Count >= 2 Then
                    .SeriesCollection(2).XValues = ws.Range("G40:G" & helperRow)
                    .SeriesCollection(2).Values = ws.Range("I40:I" & helperRow)
                End If
            End With
        End If
    End If
End Sub

' ---------------------------------------------------------------------------
' Jednorazowa instalacja / naprawa skoroszytu. Mo¿na uruchamiaæ ponownie -
' nie kasuje transakcji ani wpisanych, niezapisanych wierszy.
' ---------------------------------------------------------------------------
Public Sub InstallPortfolioTools()
    On Error GoTo Failed
    Application.ScreenUpdating = False

    SetupEntrySheet
    SetupQuotesSheet
    ApplyDashboardFormulas
    SetupAnalysisSheet
    SetupButtons

    Application.Calculate
    ThisWorkbook.Worksheets(ANALYSIS_SHEET).Activate
    Application.ScreenUpdating = True
    MsgBox "Gotowe." & vbCrLf & vbCrLf & _
           "- nowy arkusz ""Nowe transakcje"" z przyciskiem ZAPISZ TRANSAKCJE," & vbCrLf & _
           "- dynamiczna lista najwiêkszych pozycji," & vbCrLf & _
           "- poprawka cen z Londynu podawanych w pensach," & vbCrLf & _
           "- historia portfela zapisywana przy ka¿dej aktualizacji kursów," & vbCrLf & _
           "- arkusz ""Analiza"": stopa zwrotu skorygowana o wp³aty i porównanie z benchmarkiem." & vbCrLf & vbCrLf & _
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
        ' Arkusz budowany od nowa (np. po wczeœniejszej wersji z innymi kolumnami).
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
        .Range("A2").Value = "Ka¿dy wiersz = jedna operacja. Wype³nij ¿ó³te pola i kliknij ZAPISZ TRANSAKCJE. Data = dzisiaj, wszystkie kwoty w PLN."
        .Range("A3").Value = "Kupno / Sprzeda¿: instrument, liczba jednostek i kwota ca³ej transakcji.   Wp³ata / Wyp³ata: tylko kwota.   Dywidenda: instrument i kwota, która wp³ynê³a."
        .Range("A4").Value = "Gie³dê i klasê wybierasz tylko przy pierwszym kupnie nowego tickera. Jeœli któryœ wiersz ma b³¹d, nic nie zostanie zapisane - opis w kolumnie Status."
        .Range("A2:A4").Font.Color = RGB(71, 85, 105)

        .Range(.Cells(ENTRY_HEADER_ROW, COL_TYPE), .Cells(ENTRY_HEADER_ROW, COL_STATUS)).Value = _
            Array("Typ operacji", "Instrument / ticker", "Liczba jednostek", "Kwota (PLN)", _
                  "Gie³da (tylko nowy ticker)", "Klasa (tylko nowy ticker)", "Status")
        With .Range(.Cells(ENTRY_HEADER_ROW, COL_TYPE), .Cells(ENTRY_HEADER_ROW, COL_STATUS))
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(15, 23, 42)
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        .Rows(ENTRY_HEADER_ROW).RowHeight = 30

        ' Listy wyboru w ukrytych kolumnach L:N.
        .Range("L5:N5").Value = Array("Typy", "Gie³dy", "Klasy")
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
            .ErrorMessage = "Wpisz liczbê wiêksz¹ od zera."
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
        .ErrorMessage = "Wybierz wartoœæ z listy."
    End With
End Sub

' Notowania: kurs zapasowy STOCKHISTORY dla spó³ek z Londynu jest w pensach.
' Kolumna X przechowuje dzielnik (1 albo 100); formu³a ceny go uwzglêdnia.
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
            ws.Cells(r, "L").Formula = "=IF(ISNUMBER(T" & r & "),""Yahoo online"",IF(ISNUMBER(P" & r & "),""Kurs zamkniêcia"",""Cena zapasowa""))"
        End If
    Next r

    ws.Range("A4").Value = "Kursy Yahoo mog¹ byæ opóŸnione. Przy braku ceny online u¿ywany jest kurs zamkniêcia (STOCKHISTORY), a potem cena zapasowa."
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
    Dim dash As Worksheet, ledger As Worksheet, entry As Worksheet
    Dim refreshLeft As Double, refreshTop As Double
    Dim shp As Shape

    Set dash = ThisWorkbook.Worksheets("Dashboard")
    Set ledger = ThisWorkbook.Worksheets("Transakcje")
    Set entry = ThisWorkbook.Worksheets(ENTRY_SHEET)

    ' Zachowaj po³o¿enie istniej¹cego przycisku odœwie¿ania.
    refreshLeft = dash.Range("K3").Left
    refreshTop = dash.Range("K3").Top
    On Error Resume Next
    Set shp = dash.Shapes("btnRefreshPortfolio")
    On Error GoTo 0
    If Not shp Is Nothing Then
        refreshLeft = shp.Left
        refreshTop = shp.Top
    End If

    AddButton dash, "btnRefreshPortfolio", "AKTUALIZUJ KURSY", refreshLeft, refreshTop, 150, "RefreshPortfolio"
    AddButton dash, "btnOpenEntryDashboard", "DODAJ TRANSAKCJE", refreshLeft, refreshTop + 36, 150, "OpenTransactionEntry"
    AddButton dash, "btnOpenAnalysis", "ANALIZA", refreshLeft, refreshTop + 72, 150, "GoToAnalysis"
    AddButton ThisWorkbook.Worksheets(ANALYSIS_SHEET), "btnRefreshAnalysis", "AKTUALIZUJ KURSY", _
              ThisWorkbook.Worksheets(ANALYSIS_SHEET).Range("E4").Left, ThisWorkbook.Worksheets(ANALYSIS_SHEET).Range("B1").Top + 2, 150, "RefreshPortfolio"
    AddButton ledger, "btnAddTransaction", "DODAJ TRANSAKCJE", ledger.Range("O2").Left, ledger.Range("O2").Top, 175, "OpenTransactionEntry"
    AddButton entry, "btnSaveTransactions", "ZAPISZ TRANSAKCJE", entry.Range("C1").Left, entry.Range("A1").Top + 2, 175, "SaveNewTransactions"
    AddButton entry, "btnClearTransactions", "WYCZYŒÆ", entry.Range("C1").Left + 185, entry.Range("A1").Top + 2, 90, "ClearNewTransactions"
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal caption As String, _
                      ByVal leftPos As Double, ByVal topPos As Double, ByVal widthPts As Double, _
                      ByVal macroName As String)
    Dim shp As Shape
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, widthPts, 30)
    With shp
        .Name = shapeName
        .Placement = xlFreeFloating
        .TextFrame.Characters.Text = caption
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .OnAction = macroName
        .Fill.ForeColor.RGB = RGB(245, 158, 11)
        .Line.Visible = msoFalse
        .TextFrame.Characters.Font.Name = "Arial"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(15, 23, 42)
    End With
End Sub

' ---------------------------------------------------------------------------
' ARKUSZ "ANALIZA"
' Stopa zwrotu skorygowana o wp³aty (metoda TWR: zwroty miêdzy kolejnymi
' wpisami historii, z wy³¹czeniem wp³at/wyp³at, mno¿one ³añcuchowo)
' i porównanie z benchmarkiem (domyœlnie VWCE, ceny z STOCKHISTORY).
' Dane: tabela HistoriaPortfela na Dashboardzie (ukryte kolumny O:Z).
' ---------------------------------------------------------------------------
Public Sub GoToAnalysis()
    ThisWorkbook.Worksheets(ANALYSIS_SHEET).Activate
    ThisWorkbook.Worksheets(ANALYSIS_SHEET).Range("A1").Select
End Sub

Private Sub SetupAnalysisSheet()
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
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    With ws
        .Cells.Font.Name = "Arial"
        .Cells.Font.Size = 10
        .Columns("A").ColumnWidth = 2
        .Columns("B").ColumnWidth = 22
        .Columns("C:H").ColumnWidth = 19

        .Range("B1").Value = "ANALIZA PORTFELA"
        .Range("B1").Font.Size = 16
        .Range("B1").Font.Bold = True
        .Rows(1).RowHeight = 34
        .Range("B1").VerticalAlignment = xlCenter
        .Range("B2").Value = "Stopa zwrotu skorygowana o wp³aty (TWR) i porównanie z benchmarkiem. Nowy wpis w historii powstaje przy ka¿dej aktualizacji kursów, raz dziennie."
        .Range("B2").Font.Color = RGB(71, 85, 105)

        ' Ustawienia
        .Range("B4").Value = "Okres w tabeli"
        .Range("C4").Value = "Miesi¹c"
        .Range("B5").Value = "Benchmark (ticker)"
        .Range("C5").Value = "XETR:VWCE"
        .Range("B6").Value = "Waluta benchmarku"
        .Range("C6").Value = "EUR"
        .Range("D5").Value = "np. XETR:VWCE, XNAS:QQQ, XLON:VUSA - zapis jak w STOCKHISTORY"
        .Range("D5:D6").Font.Color = RGB(100, 116, 139)
        .Range("D6").Value = "waluta notowañ benchmarku (PLN, EUR, USD, GBP)"
        .Range("B7").Value = "Historia od / do"
        .Range("C7").Formula2 = "=IFERROR(MIN(HistoriaPortfela[Data]),"""")"
        .Range("D7").Formula2 = "=IFERROR(MAX(HistoriaPortfela[Data]),"""")"
        .Range("C7:D7").NumberFormat = "dd.mm.yyyy"
        .Range("C7:D7").HorizontalAlignment = xlLeft
        .Range("B4:B7").Font.Bold = True
        .Range("C4:C6").Interior.Color = RGB(255, 247, 214)
        .Range("C4:C6").Borders.Color = RGB(226, 232, 240)
        .Range("Y5:Y7").Value = Application.WorksheetFunction.Transpose(Array("Dzieñ", "Miesi¹c", "Rok"))
        .Range("Z5:Z8").Value = Application.WorksheetFunction.Transpose(Array("PLN", "EUR", "USD", "GBP"))
        AddListValidation .Range("C4"), "=$Y$5:$Y$7", True
        AddListValidation .Range("C6"), "=$Z$5:$Z$8", True

        ' Ceny benchmarku i kursy walut (ukryte kolumny AA:AE)
        .Range("AA4").Value = "Benchmark: data, cena"
        .Range("AA5").Formula2 = "=IFERROR(STOCKHISTORY($C$5,INT(MIN(HistoriaPortfela[Data]))-10,TODAY(),0,0,0,1),""brak"")"
        .Range("AD4").Value = "Kurs waluty: data, kurs"
        .Range("AD5").Formula2 = "=IF($C$6=""PLN"",""PLN"",IFERROR(STOCKHISTORY($C$6&""/PLN"",INT(MIN(HistoriaPortfela[Data]))-10,TODAY(),0,0,0,1),""brak""))"
    End With

    ' Kolumny obliczeniowe w tabeli historii (wymagaj¹ arkusza Analiza).
    Set lo = SetupHistoryColumns()

    With ws
        ' WskaŸniki
        labels = Array("Wartoœæ portfela", "Stopa zwrotu portfela (TWR)", "Stopa zwrotu benchmarku", _
                       "Portfel - benchmark (pkt proc.)", "Gdyby wp³aty sz³y w benchmark", "Portfel - benchmark (PLN)")
        For i = 0 To 5
            .Cells(9, 2 + i).Value = labels(i)
        Next i
        With .Range("B9:G9")
            .Font.Size = 9
            .Font.Color = RGB(165, 178, 196)
            .Interior.Color = RGB(20, 36, 61)
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        .Rows(9).RowHeight = 28
        .Range("B10").Formula2 = "=IFERROR(INDEX(HistoriaPortfela[Wartoœæ portfela],ROWS(HistoriaPortfela[Wartoœæ portfela])),"""")"
        .Range("C10").Formula2 = "=IFERROR(INDEX(HistoriaPortfela[Stopa portfela],ROWS(HistoriaPortfela[Stopa portfela])),"""")"
        .Range("D10").Formula2 = "=IFERROR(INDEX(HistoriaPortfela[Stopa benchmarku],ROWS(HistoriaPortfela[Stopa benchmarku])),"""")"
        .Range("E10").Formula2 = "=IFERROR(C10-D10,"""")"
        .Range("F10").Formula2 = "=IFERROR(INDEX(HistoriaPortfela[Wartoœæ w benchmarku],ROWS(HistoriaPortfela[Wartoœæ w benchmarku])),"""")"
        .Range("G10").Formula2 = "=IFERROR(B10-F10,"""")"
        With .Range("B10:G10")
            .Font.Size = 15
            .Font.Bold = True
            .Font.Color = RGB(245, 160, 0)
            .Interior.Color = RGB(20, 36, 61)
            .HorizontalAlignment = xlLeft
        End With
        .Rows(10).RowHeight = 30
        .Range("B10,F10,G10").NumberFormat = "#,##0.00 ""z³"";[Red]-#,##0.00 ""z³"""
        .Range("C10:E10").NumberFormat = "0.00%;[Red]-0.00%"

        ' Tabela okresów
        .Range("B33").Value = "Wyniki w okresach (najnowsze na górze) - okres zmienisz w komórce C4"
        .Range("B33").Font.Bold = True
        .Range("B34:H34").Value = Array("Okres", "Wartoœæ na koniec", "Wp³aty netto w okresie", _
                                        "Zysk w okresie (PLN)", "Stopa zwrotu portfela", "Stopa zwrotu benchmarku", "Ró¿nica (pkt proc.)")
        With .Range("B34:H34")
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(15, 23, 42)
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        .Rows(34).RowHeight = 30
        .Range("B35").Formula2 = PeriodTableFormula()
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
        "keyArr,IF(okres=""Rok"",YEAR(dni),IF(okres=""Miesi¹c"",YEAR(dni)*100+MONTH(dni),dni))," & _
        "uniq,UNIQUE(keyArr),seq,SEQUENCE(ROWS(uniq))," & _
        "endIx,XMATCH(uniq,keyArr,0,-1)," & _
        "prvIx,IF(seq=1,1,INDEX(endIx,IF(seq=1,1,seq-1)))," & _
        "valA,HistoriaPortfela[Wartoœæ portfela],depA,HistoriaPortfela[Wp³aty netto]," & _
        "idxA,HistoriaPortfela[Indeks portfela],bixA,HistoriaPortfela[Indeks benchmarku]," & _
        "vEnd,INDEX(valA,endIx),vSt,INDEX(valA,prvIx),dEnd,INDEX(depA,endIx),dSt,INDEX(depA,prvIx)," & _
        "rPort,INDEX(idxA,endIx)/INDEX(idxA,prvIx)-1," & _
        "rBen,IFERROR(INDEX(bixA,endIx)/INDEX(bixA,prvIx)-1,"""")," & _
        "lbl,IF(okres=""Rok"",""""&uniq,IF(okres=""Miesi¹c"",INT(uniq/100)&""-""&RIGHT(""0""&MOD(uniq,100),2)," & _
            "YEAR(uniq)&""-""&RIGHT(""0""&MONTH(uniq),2)&""-""&RIGHT(""0""&DAY(uniq),2)))," & _
        "SORTBY(HSTACK(lbl,vEnd,dEnd-dSt,vEnd-vSt-(dEnd-dSt),rPort,rBen,IFERROR(rPort-rBen,"""")),uniq,-1))," & _
        """Brak danych w historii portfela"")"
End Function

' Dodaje do tabeli HistoriaPortfela kolumny obliczeniowe (wype³niaj¹ siê same
' dla ka¿dego nowego wpisu) i usuwa wpisy bez wartoœci portfela.
Private Function SetupHistoryColumns() As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim specs As Variant
    Dim firstRow As String

    Set ws = ThisWorkbook.Worksheets("Dashboard")
    Set lo = ws.ListObjects("HistoriaPortfela")

    For i = lo.ListRows.Count To 1 Step -1
        If Len(Trim$(CStr(lo.ListRows(i).Range.Cells(1, 2).Value))) = 0 Then lo.ListRows(i).Delete
    Next i

    firstRow = "ROW()=ROW(HistoriaPortfela[#Headers])+1"
    specs = Array( _
        Array("Przep³yw", "=IF(" & firstRow & ",0,[@[Wp³aty netto]]-OFFSET([@[Wp³aty netto]],-1,0))", "#,##0.00"), _
        Array("Zwrot okresu", "=IF(" & firstRow & ",0,IFERROR([@[Wartoœæ portfela]]/(OFFSET([@[Wartoœæ portfela]],-1,0)+[@Przep³yw])-1,0))", "0.00%"), _
        Array("Indeks portfela", "=IF(" & firstRow & ",1,OFFSET([@[Indeks portfela]],-1,0)*(1+[@[Zwrot okresu]]))", "0.0000"), _
        Array("Stopa portfela", "=[@[Indeks portfela]]-1", "0.00%"), _
        Array("Cena benchmarku (PLN)", "=IFERROR(XLOOKUP(INT([@Data]),INDEX(Analiza!$AA$5#,0,1),INDEX(Analiza!$AA$5#,0,2),,-1)*" & _
              "IF(Analiza!$C$6=""PLN"",1,XLOOKUP(INT([@Data]),INDEX(Analiza!$AD$5#,0,1),INDEX(Analiza!$AD$5#,0,2),,-1))," & _
              "IF(" & firstRow & ","""",OFFSET([@[Cena benchmarku (PLN)]],-1,0)))", "#,##0.00"), _
        Array("Indeks benchmarku", "=IFERROR([@[Cena benchmarku (PLN)]]/INDEX([Cena benchmarku (PLN)],1),"""")", "0.0000"), _
        Array("Stopa benchmarku", "=IFERROR([@[Indeks benchmarku]]-1,NA())", "0.00%"), _
        Array("Jednostki benchmarku", "=IFERROR(IF(" & firstRow & ",[@[Wartoœæ portfela]]/[@[Cena benchmarku (PLN)]]," & _
              "OFFSET([@[Jednostki benchmarku]],-1,0)+[@Przep³yw]/[@[Cena benchmarku (PLN)]]),"""")", "0.0000"), _
        Array("Wartoœæ w benchmarku", "=IFERROR([@[Jednostki benchmarku]]*[@[Cena benchmarku (PLN)]],NA())", "#,##0.00"))

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
    lo.Range.EntireColumn.Hidden = True
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
    PrepareLineChart co.Chart, "Wartoœæ portfela, wp³aty i benchmark (PLN)"
    AddLineSeries co.Chart, "Wartoœæ portfela", lo, "Wartoœæ portfela", RGB(245, 158, 11)
    AddLineSeries co.Chart, "Wp³aty netto", lo, "Wp³aty netto", RGB(100, 116, 139)
    AddLineSeries co.Chart, "Gdyby wp³aty sz³y w benchmark", lo, "Wartoœæ w benchmarku", RGB(39, 182, 232)
    FormatChartAxes co.Chart, "#,##0"

    Set co = ws.ChartObjects.Add(leftPos + w + 12, topPos, ws.Range("F12:H12").Width + 40, h)
    co.Name = "WykresStopyZwrotu"
    PrepareLineChart co.Chart, "Skumulowana stopa zwrotu (TWR) vs benchmark"
    AddLineSeries co.Chart, "Portfel", lo, "Stopa portfela", RGB(245, 158, 11)
    AddLineSeries co.Chart, "Benchmark", lo, "Stopa benchmarku", RGB(39, 182, 232)
    FormatChartAxes co.Chart, "0.0%"
End Sub

Private Sub PrepareLineChart(ByVal ch As Chart, ByVal title As String)
    ch.ChartType = xlLine
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
    ch.PlotVisibleOnly = False          ' dane historii s¹ w ukrytych kolumnach
    ch.DisplayBlanksAs = xlInterpolated
    ch.HasTitle = True
    ch.ChartTitle.Text = title
    ch.ChartTitle.Font.Size = 12
    ch.ChartTitle.Font.Bold = True
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    ch.ChartArea.Font.Name = "Arial"
    ch.ChartArea.Format.Line.Visible = msoFalse
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
    s.MarkerSize = 4
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
    ch.Axes(xlValue).HasMajorGridlines = True
    ch.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(226, 232, 240)
End Sub
