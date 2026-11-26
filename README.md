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

- **Dashboard**: wartość portfela, wpłaty netto, wynik, skład portfela, wykres zysków/strat spółek, największe pozycje, historia wartości.
- **Nowe transakcje**: wpisywanie operacji (kupno, sprzedaż, wpłata, wypłata, dywidenda) i przycisk ZAPISZ TRANSAKCJE.
- **Transakcje**: rejestr wszystkich operacji; liczy jednostki, koszt średni, gotówkę i zrealizowany wynik.
- **Notowania**: kursy instrumentów i walut (Yahoo, potem `STOCKHISTORY`, potem cena zapasowa).
- **Pozycje** (ukryty): stan każdej pozycji.

## Instalacja / aktualizacja makr

1. Zrób kopię zapasową `Money Money.xlsm`.
2. Otwórz plik, włącz makra, naciśnij **Alt+F11**.
3. Usuń wszystkie moduły z folderu *Modules* (prawy klik → *Remove* → „Nie” przy pytaniu o eksport).
4. **File → Import File…** → `vba/Portfel.bas`.
5. **Debug → Compile VBAProject** (bez komunikatów).
6. W Excelu **Alt+F8** → `InstallPortfolioTools` → *Uruchom*.
7. Zapisz plik (**Ctrl+S**).

Instalator można uruchamiać ponownie: nie usuwa transakcji.

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
- Po każdej udanej aktualizacji do tabeli `HistoriaPortfela` na Dashboardzie dopisywany jest wpis (jeden dziennie).
