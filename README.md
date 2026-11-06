# Outlook Speech Filter

Dodatek dla czytnika ekranu **NVDA**, który wycisza uciążliwe i rozwlekłe komunikaty o statusie wiadomości (stan przeczytania, załączniki, odpowiedzi, ważność, nadawcy zewnętrzni) w programie **Microsoft Outlook Classic**, zastępując je dyskretnymi sygnałami dźwiękowymi (earconami). Zawiera także zaawansowany nawigator wątków wiadomości i skróty szybkiej nawigacji.

---

## ✨ Kluczowe możliwości dodatku

1. **Wyciszanie gadatliwości i zastępowanie dźwiękami (Earcony):**
   - Zastępuje czytanie statusów (*„przeczytana”*, *„nieprzeczytana”*, *„ma załącznik”*, *„odpowiedziałeś”*, *„wysoka ważność”*, *„zewnętrzna”*) 10 krótkimi, przyjemnymi i dyskretnymi dźwiękami.
   - Płynna regulacja głośności dźwięków w ustawieniach dodatku z możliwością odsłuchu każdego sygnału.
   - Pełna swoboda wyboru: możesz zdecydować, które komunikaty mają być zastąpione dźwiękiem, które wyciszone całkowicie, a które czytane słownie.

2. **Zaawansowany Nawigator Wątków (NVDA+F7 / NVDA+Alt+F7):**
   - Przeglądanie całego wątku konwersacji mailowej w przystępnym oknie dialogowym.
   - Drzewiasta lub chronologiczna struktura wiadomości w wątku.
   - Szybkie przeskakiwanie do wybranego maila lub cytatu z bezpośrednim ustawieniem fokusu w oknie Outlooka.

3. **Błyskawiczne przeskakiwanie między mailami (NVDA+Alt+Strzałki):**
   - Skróty klawiszowe do przejścia do kolejnej lub poprzedniej wiadomości w wątku bez konieczności zamykania bieżącego okna.

4. **Raport stanu wiadomości na żądanie (NVDA+Shift+Z):**
   - Gdy potrzebujesz pełnego, szczegółowego podsumowania stanu bieżącej wiadomości, jeden skrót odczyta wszystkie jej atrybuty.

5. **Dostępność i wielojęzyczność:**
   - Zaprojektowany z zachowaniem standardów dostępności cyfrowej (**WCAG 2.1 / 2.2 AA**).
   - Pełna obsługa języka polskiego i angielskiego (automatyczne dopasowanie do języka NVDA).
   - Pomoc kontekstowa pod klawiszem `F1` oraz `NVDA+Ctrl+Shift+H`.

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
| **F1** (w oknach dodatku) | Pomoc kontekstowa |

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
