RegDoktor 11 - przeglądarka rejestru i centrum napraw Windows 11
==================================================================

URUCHOMIENIE
  1. Rozpakuj cały folder (np. na Pulpit).
     Jeśli plik ZIP był pobrany z internetu: PPM na ZIP > Właściwości > zaznacz "Odblokuj" > OK,
     dopiero potem rozpakuj (inaczej Windows może pytać o zgodę przy każdym uruchomieniu).
  2. Kliknij dwukrotnie "Uruchom RegDoktor.bat" i zgódź się na uprawnienia administratora.
     (Jeśli SmartScreen pokaże "System Windows ochronił ten komputer": Więcej informacji > Uruchom mimo to.)
  Program to jeden skrypt PowerShell (RegDoktor.ps1) - nic nie instaluje, nie wymaga internetu
  (poza naprawami, które same z niego korzystają: DISM, Windows Update, Defender, test sieci).

CO ZAWIERA
  * Naprawy i optymalizacja - kolorowe kafelki w 9 kategoriach:
      System Windows | Sieć i Internet | Dysk i wolne miejsce | Pamięć RAM | Wydajność i rozruch
      Eksplorator, Start i wygląd | Urządzenia, dźwięk i druk | Bezpieczeństwo | Rejestr
    Pole "Szukaj naprawy" filtruje kafelki (np. "dns", "drukarka", "menu start").
  * Przeglądarka rejestru - drzewo kluczy, wartości, dodawanie/edycja/usuwanie, wyszukiwanie,
    eksport i import .reg, kopiowanie ścieżki, otwieranie w Regedit.
  * Osierocone wpisy - skanowanie 13 kategorii (wpisy po odinstalowanych programach, brakujące
    pliki, puste klucze...). Każdy wpis ma poziom ryzyka; domyślnie zaznaczone są tylko bezpieczne.
  * Defragmentacja rejestru - analiza realnego zysku i kompaktowanie gałęzi przy restarcie.
  * Autostart - włączanie/wyłączanie programów startowych (jak w Menedżerze zadań).

BEZPIECZEŃSTWO
  * Każda zmiana w rejestrze jest poprzedzona kopią .reg:  Dokumenty\RegDoktor\Kopie zapasowe
    Przywracanie: zakładka "Osierocone wpisy" > "Przywróć z kopii..." (albo dwuklik na pliku .reg).
  * Przy ryzykownych operacjach program sam tworzy punkt przywracania systemu
    (przełącznik w nagłówku okna).
  * Najważniejsze klucze systemowe są zablokowane przed usunięciem w przeglądarce.
  * Dziennik wszystkich działań: Dokumenty\RegDoktor\Dziennik_RRRR-MM.log

UCZCIWIE O "CZYSZCZENIU" I "PRZYSPIESZANIU" REJESTRU
  Microsoft nie zaleca agresywnych "czyścicieli rejestru" - pojedyncze osierocone wpisy nie spowalniają
  Windows. Dlatego RegDoktor usuwa wyłącznie wpisy, które wskazują na nieistniejące pliki lub są
  całkowicie puste, i zawsze robi kopię. Realne przyspieszenie dają raczej: autostart, plan zasilania,
  wolne miejsce na dysku, TRIM/defragmentacja dysku i naprawa uszkodzonych plików systemowych (DISM + SFC).
  Windows od wersji 8 sam kompaktuje pliki rejestru, gdy zysk jest duży - zakładka
  "Defragmentacja rejestru" najpierw pokazuje, ile realnie można odzyskać.

WYMAGANIA
  Windows 10 / 11 (64-bit), Windows PowerShell 5.1 (wbudowany), konto z uprawnieniami administratora.
  Zmiany w HKEY_CURRENT_USER dotyczą konta, na którym zatwierdzono uprawnienia administratora.
