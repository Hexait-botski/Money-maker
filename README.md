# Money Maker

Dashboard do śledzenia osobistego portfela inwestycyjnego w Excelu (Microsoft 365, Windows).
Wszystkie kwoty w PLN.

Repozytorium zawiera **tylko kod makr**. Skoroszyt `Money Money.xlsm` z danymi
portfela nie jest tu przechowywany (patrz `.gitignore`).

## Zawartość

| Plik | Opis |
|---|---|
| `vba/Portfel.bas` | Cały kod VBA w jednym module: kursy z Yahoo, arkusz „Nowe transakcje”, historia portfela, instalator. |

`Portfel.bas` jest zapisany w kodowaniu **Windows-1250** z końcami linii CRLF, bo tak
importuje go edytor VBA. Na GitHubie polskie znaki w podglądzie mogą wyglądać źle,
ale pobrany plik (przycisk *Download raw file*) jest poprawny.

## Arkusze skoroszytu

- **Dashboard** (ciemny motyw): kafelki z wartością portfela, zyskiem, stopą zwrotu, wpłatami, gotówką i dywidendami z 12 miesięcy; wykres wartości na tle wpłat; skład portfela według klasy i waluty; zysk/strata każdej pozycji; tabela pozycji z filtrem (Wszystkie / Akcje / ETF-y) i wyszukiwarką.
- **Dywidendy**: otrzymane dywidendy (łącznie, w tym roku, ostatnie 12 miesięcy, średnio miesięcznie, stopa od kosztu), wykres z 24 miesięcy, tabela spółek i historia wypłat; z boku szacunek podatku Belki.
- **Analiza**: stopa zwrotu skorygowana o wpłaty (TWR), porównanie z benchmarkiem (domyślnie VWCE), wykresy i tabela wyników w okresach (dzień / miesiąc / rok).
- **Nowe transakcje**: wpisywanie operacji (kupno, sprzedaż, wpłata, wypłata, dywidenda) i przycisk ZAPISZ TRANSAKCJE.
- **Transakcje**: rejestr wszystkich operacji; liczy jednostki, koszt średni, gotówkę i zrealizowany wynik.
- **Notowania**: kursy instrumentów i walut (Yahoo, potem `STOCKHISTORY`, potem cena zapasowa).
- **Pozycje** (ukryty): stan każdej pozycji.
- **Dane** (ukryty): tabela `HistoriaPortfela` i liczby, z których korzysta Dashboard.

## Instalacja / aktualizacja makr

1. Zrób kopię zapasową `Money Money.xlsm`.
2. Otwórz plik, włącz makra, naciśnij **Alt+F11**.
3. Usuń wszystkie moduły z folderu *Modules* (prawy klik → *Remove* → „Nie” przy pytaniu o eksport).
4. **File → Import File…** → `vba/Portfel.bas`.
5. **Debug → Compile VBAProject** (bez komunikatów).
6. W Excelu **Alt+F8** → `InstallPortfolioTools` → *Uruchom*.
7. Zapisz plik (**Ctrl+S**).

Instalator można uruchamiać ponownie: nie usuwa transakcji ani historii. Dashboard i Analiza są
budowane od nowa, a przy pierwszym uruchomieniu tabela historii jest przenoszona z Dashboardu
do ukrytego arkusza Dane.

## Wpisywanie transakcji

Arkusz **Nowe transakcje**, jeden wiersz = jedna operacja, potem **ZAPISZ TRANSAKCJE**.

| Typ | Instrument | Liczba jednostek | Kwota (PLN) |
|---|---|---|---|
| Kupno | NVIDIA | 0,1620 | 99,99 |
| Sprzedaż | Apple | 0,2 | 250,00 |
| Wpłata | | | 500,00 |
| Dywidenda | Coca-Cola | | 4,12 |

- Kwota to zawsze suma w PLN (ile zapłaciłeś / dostałeś / ile wpłynęło).
- Przy pierwszym kupnie nowego tickera wybierz też giełdę (GPW, NASDAQ, NYSE, Xetra, LSE) i klasę (Akcje / ETF-y).
- Jeśli któryś wiersz ma błąd, nic nie jest zapisywane, a opis pojawia się w kolumnie *Status*.

## Kursy

- Pobierane przy otwarciu pliku (bez komunikatu) i przyciskiem **AKTUALIZUJ KURSY**.
- Źródło: nieoficjalne API Yahoo Finance (kursy mogą być opóźnione). Gdy Yahoo nie odpowie,
  używany jest kurs zamknięcia z `STOCKHISTORY`, a potem cena zapasowa.
- Spółki z Londynu notowane w pensach (np. National Grid) są przeliczane na GBP.
- Po każdej udanej aktualizacji do tabeli `HistoriaPortfela` (arkusz Dane) dopisywany jest wpis (jeden dziennie).

## Analiza

- **Stopa zwrotu portfela (TWR)**: zwroty między kolejnymi wpisami historii, z wyłączeniem
  wpłat i wypłat, mnożone łańcuchowo. Wpłata nie zawyża wyniku, a wypłata go nie zaniża.
- **Benchmark**: ticker w komórce `C5` (zapis jak w `STOCKHISTORY`, np. `XETR:VWCE`) i jego waluta w `C6`.
  Ceny i kursy walut pobiera `STOCKHISTORY` (kurs zamknięcia z danego dnia lub ostatniego dnia notowań).
- **„Gdyby wpłaty szły w benchmark”**: na starcie historii cała wartość portfela trafia do benchmarku,
  a każda kolejna wpłata/wypłata kupuje/sprzedaje jego jednostki po cenie z tego dnia.
- **Okres w tabeli** zmieniasz w komórce `C4` (Dzień / Miesiąc / Rok).
- Analiza obejmuje okres od pierwszego wpisu w historii (19.11.2026); wcześniejszej historii wartości nie ma.

## Dywidendy i podatek Belki

- **Automatycznie**: raz dziennie (przy aktualizacji kursów) makro pobiera z Yahoo dywidendy na akcję
  z ostatnich 5 lat i liczy: dywidenda na akcję × liczba jednostek w dniu ustalenia prawa (ex-date) × kurs waluty,
  minus podatek u źródła. Wyniki trafiają do tabeli `DywidendyHistoria` (arkusz Dane) ze źródłem
  „automatycznie (szacunek)”. To tylko statystyki: gotówka się nie zmienia.
- **Wypłaty otrzymane** (wpisane przez użytkownika albo typ Dywidenda w rejestrze) mają pierwszeństwo:
  automat liczy tylko dywidendy z ex-date po ostatniej wpisanej wypłacie danej spółki.
- **Pozycje z salda początkowego** nie mają daty zakupu; panel „Od kiedy masz akcje” w arkuszu Dywidendy
  mówi automatowi, od kiedy liczyć ich dywidendy wstecz. Bez daty automat liczy od dnia instalacji.
- **Prognoza na 12 miesięcy** = dywidendy na akcję z ostatnich 12 miesięcy × obecna liczba jednostek, netto.
  Z niej liczone są stopa dywidendy portfela i stopa od kosztu.
- Kursy walut do przeliczenia dywidend są bieżące (szacunek).
- **Podatek Belki (szacunek, dodatek)**: 19% od zrealizowanego wyniku ze sprzedaży w wybranym roku oraz dopłata
  do dywidend zagranicznych. Kwotę brutto i podatek pobrany u źródła wylicza się ze stawek w tabeli założeń
  (domyślnie PLN 19%, USD 15%, GBP 0%, EUR 0%; można je zmienić). Urząd liczy koszt metodą FIFO i po kursach NBP,
  więc wynik to tylko orientacja; do PIT-38 służy PIT-8C od brokera.

## Import z Trading 212

Przycisk **IMPORTUJ Z TRADING 212** (Dashboard, Dywidendy) wczytuje pliki CSV z historii konta
(Trading 212 → Historia → Eksportuj; można zaznaczyć kilka plików naraz):

- **dywidendy** z całej historii → tabela `DywidendyHistoria` (kwota netto i podatek pobrany u źródła),
- **wynik sprzedaży** sprzed dnia salda początkowego → tabela `WynikiSprzedazy` (do szacunku podatku Belki),
- **zakupy, sprzedaże, wpłaty i wydatki kartą** po dniu salda początkowego → rejestr w arkuszu Transakcje
  (wcześniejsze operacje są już w saldzie początkowym). Nowe spółki są dodawane automatycznie.

Każdy zaimportowany wiersz ma klucz (kolumna S w Transakcjach, kolumna „Klucz” w tabelach), więc ponowny
import tych samych plików niczego nie dubluje. Dzień salda początkowego jest zapisany w `Dane!AB32`.
