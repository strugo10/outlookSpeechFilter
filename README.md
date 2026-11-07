# Outlook Speech Filter

Dodatek dla czytnika ekranu **NVDA**, który wycisza uciążliwe i rozwlekłe komunikaty o statusie wiadomości (stan przeczytania, załączniki, odpowiedzi, ważność, nadawcy zewnętrzni) w programie **Microsoft Outlook Classic** (wersja klasyczna dla systemu Windows: 2013, 2016, 2019, 2021, Microsoft 365), zastępując je dyskretnymi sygnałami dźwiękowymi (earconami). Zawiera także zaawansowany nawigator wątków wiadomości i skróty szybkiej nawigacji.

---

## ✨ Kluczowe możliwości dodatku

1. **Wyciszanie gadatliwości i zastępowanie dźwiękami (Earcony):**
   - Zastępuje czytanie statusów (*„przeczytana”*, *„nieprzeczytana”*, *„ma załącznik”*, *„odpowiedziałeś”*, *„wysoka ważność”*, *„zewnętrzna”*) 10 krótkimi, przyjemnymi i dyskretnymi dźwiękami.
   - Płynna regulacja głośności dźwięków w ustawieniach dodatku z możliwością odsłuchu każdego sygnału.
   - **Dedykowane okno konfiguracji dźwięków:** przejrzyste okno dialogowe z możliwością włączania i odsłuchu każdego z 11 sygnałów zdarzeń z osobna.
   - **Stan częściowego wyboru:** główne pole wyboru odtwarzania dźwięków natywnie informuje o stanie „częściowo zaznaczone”, gdy tylko wybrane dźwięki są aktywne.

2. **Zaawansowany Nawigator Wątków (NVDA+F7 / NVDA+Alt+F7):**
   - Przeglądanie całego wątku konwersacji mailowej w przystępnym oknie dialogowym.
   - Drzewiasta lub chronologiczna struktura wiadomości w wątku.
   - Szybkie przeskakiwanie do wybranego maila lub cytatu z bezpośrednim ustawieniem fokusu w oknie Outlooka.
   - Czyste kopiowanie treści wybranego maila skrótem <kbd>Alt + K</kbd> bez cytatów i zbędnych nagłówków.

3. **Błyskawiczne przeskakiwanie między mailami (NVDA+Alt+Strzałki):**
   - Skróty klawiszowe do przejścia do kolejnej lub poprzedniej wiadomości w wątku bez konieczności zamykania bieżącego okna.

4. **Raport stanu wiadomości na żądanie (NVDA+Shift+Z):**
   - Gdy potrzebujesz pełnego, szczegółowego podsumowania stanu bieżącej wiadomości, jeden skrót odczyta wszystkie jej atrybuty.
   - Pełna obsługa podsumowania zaznaczenia wielu wiadomości na liście maili.

5. **Funkcja „Podziel się dodatkiem ze znajomymi”:**
   - Gotowe okno dialogowe z linkami do oficjalnego wydania na GitHubie oraz gotową treścią e-maila polecającego.

6. **Interaktywny samouczek audio w pomocy HTML:**
   - Wbudowane odtwarzacze HTML5 dla każdego dźwięku oraz symulator sekwencji dźwiękowych wprost w przeglądarce.

7. **Dostępność i wielojęzyczność:**
   - Zaprojektowany z zachowaniem standardów dostępności cyfrowej (**WCAG 2.1 / 2.2 AA**).
   - Pełna obsługa języka polskiego i angielskiego (automatyczne dopasowanie do języka NVDA).
   - Pomoc kontekstowa pod klawiszem `F1` oraz `NVDA+Ctrl+Shift+H`.

---

## 💻 Wymagania systemowe

- **Czytnik ekranu:** NVDA 2024.1 – 2026.1+
- **Klient poczty:** Microsoft Outlook Classic (wersje 2013, 2016, 2019, 2021 oraz Microsoft 365 dla Windows; dodatek nie jest przeznaczony dla nowego webowego klienta „New Outlook”).
- **System operacyjny:** Microsoft Windows 10 / 11

---

## ⌨️ Najważniejsze skróty klawiszowe

| Skrót klawiszowy | Działanie |
| :--- | :--- |
| **NVDA+F7** | Otwiera okno dialogowe Nawigatora Wątku dla bieżącej wiadomości |
| **NVDA+Alt+F7** | Szybki podgląd informacji o wątku |
| **NVDA+Ctrl+Shift+T** | Włączenie / wyłączenie nawigatora wątków |
| **NVDA+Alt+Strzałka w dół** | Przejście do następnej wiadomości w wątku |
| **NVDA+Alt+Strzałka w górę** | Przejście do poprzedniej wiadomości w wątku |
| **NVDA+Shift+Z** | Raport stanu wiadomości na żądanie (załączniki, status, ważność itp.) |
| **NVDA+Ctrl+Shift+O** | Szybkie włączenie / wyłączenie wyciszania zbędnych komunikatów mowy |
| **NVDA+Ctrl+Shift+S** | Szybkie włączenie / wyłączenie odtwarzania dyskretnych dźwięków |
| **NVDA+Ctrl+Shift+H** | Pomoc kontekstowa i spis skrótów klawiszowych |
| **F1** (w oknach dodatku) | Pomoc kontekstowa w danym oknie |

---

## 🌐 Oficjalne repozytorium i zgłoszenia

- **Strona projektu na GitHubie:** [https://github.com/strugo10/outlookSpeechFilter](https://github.com/strugo10/outlookSpeechFilter)
- **Zgłaszanie uwag i błędów (GitHub Issues):** [https://github.com/strugo10/outlookSpeechFilter/issues](https://github.com/strugo10/outlookSpeechFilter/issues)
- **Pobieranie najnowszej wersji instalatora:** [https://github.com/strugo10/outlookSpeechFilter/releases/latest](https://github.com/strugo10/outlookSpeechFilter/releases/latest)

---

## 🛠️ Budowanie paczki instalacyjnej (.nvda-addon)

Aby zbudować plik instalacyjny dodatku, uruchom w PowerShellu skrypt:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File build.ps1
```

---

## 👤 Autor

- **Sławomir Strugarek** (przy wsparciu asystenta Antigravity / Google DeepMind)
- Licencja: GNU General Public License v2 (lub nowsza)
