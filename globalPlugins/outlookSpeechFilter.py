# -*- coding: utf-8 -*-
# NVDA Add-on: Outlook Speech Filter
# Author: Sławomir Strugarek z pomocą Antigravity (Google DeepMind)
# Description: Controls and mutes unwanted message statuses (attachments, read/unread, replied/forwarded, etc.)
# in Microsoft Outlook's message list.

from __future__ import annotations
import array
import ctypes
import json
import locale
import os
import platform
import re
import shutil
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
import wave
import weakref
from comtypes import COMError
import comtypes.automation
from comtypes.hresult import S_OK

import addonHandler
import api
import config
try:
	from config.configFlags import ReportTableHeaders
except (ImportError, AttributeError):
	ReportTableHeaders = None
import controlTypes
import globalPluginHandler
import globalVars
import gui
import gui.guiHelper
import gui.settingsDialogs
from gui.settingsDialogs import SettingsPanel
from logHandler import log
import NVDAHelper
import nvwave
from scriptHandler import script
import ui
import UIAHandler
import winsound
import wx

try:
	import nvdaBuiltin.appModules.outlook as builtinOutlook
except ImportError:
	builtinOutlook = None

addonHandler.initTranslation()

CONFIG_SECTION = "outlookSpeechFilter"
OLD_CONFIG_SECTION = "outlookMessageCleaner"
ADDON_NAME = "Outlook Speech Filter"
ADDON_VERSION = "1.18.2"
ADDON_BUILD = "20261103.2"
ADDON_BUILD_DATE = "2026-11-03"
ADDON_AUTHOR = "Sławomir Strugarek (z pomocą Antigravity)"
ADDON_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SOUNDS_DIR = os.path.join(ADDON_DIR, "sounds")
SOUNDS_CACHE_DIR = os.path.join(tempfile.gettempdir(), "nvda_osf_sounds")
DOC_PATH = os.path.join(ADDON_DIR, "doc.html")
DOC_PATH_PL = os.path.join(ADDON_DIR, "doc_pl.html")
DOC_PATH_EN = os.path.join(ADDON_DIR, "doc_en.html")
DEFAULT_FEEDBACK_FORM_URL_PL = "https://docs.google.com/forms/d/e/1FAIpQLSfylH_KLg3W2hNkck3tJ7SHj_zAExvSs5HsUwbw78Faeml64A/viewform"
DEFAULT_FEEDBACK_FORM_URL_EN = "https://docs.google.com/forms/d/e/1FAIpQLSfVFKRn-5EDsp0cjU0fl8H0WVOdKxDcQjhHPhY9eDzG6YF_KA/viewform"
DEFAULT_FEEDBACK_FORM_URL = DEFAULT_FEEDBACK_FORM_URL_PL

confspec = {
	"enabled": "boolean(default=True)",
	"threadNavigationEnabled": "boolean(default=True)",
	"language": "string(default='auto')",
	"soundTheme": "string(default='associative')",
	"feedbackFormUrl": "string(default='')",
	"useSounds": "boolean(default=True)",
	"soundVolume": "integer(default=80, min=0, max=100)",
	"soundUnread": "boolean(default=True)",
	"soundFlags": "boolean(default=False)",
	"soundAttachments": "boolean(default=True)",
	"soundInlineImages": "boolean(default=True)",
	"soundReplied": "boolean(default=False)",
	"soundForwarded": "boolean(default=False)",
	"soundImportanceHigh": "boolean(default=True)",
	"soundImportanceLow": "boolean(default=False)",
	"soundMeetingRequest": "boolean(default=True)",
	"soundCategories": "boolean(default=False)",
	"soundExternal": "boolean(default=False)",
	# Kompatybilność wsteczna
	"soundForwardReply": "boolean(default=False)",
	"soundImportance": "boolean(default=True)",
	"reportUnread": "boolean(default=True)",
	"reportFlags": "boolean(default=False)",
	"reportAttachments": "boolean(default=True)",
	"reportInlineImages": "boolean(default=False)",
	"reportReplied": "boolean(default=False)",
	"reportForwarded": "boolean(default=False)",
	"reportImportanceHigh": "boolean(default=True)",
	"reportImportanceLow": "boolean(default=False)",
	"reportForwardReply": "boolean(default=False)",
	"reportImportance": "boolean(default=True)",
	"reportMeetingRequest": "boolean(default=True)",
	"reportCategories": "boolean(default=False)",
	"reportExternal": "boolean(default=False)",
	"autoCheckUpdates": "boolean(default=True)",
	"updateSourceUrl": r"string(default='https://raw.githubusercontent.com/strugo10/outlookSpeechFilter/main/update_info.json')",
	"lastUpdateCheckTime": "float(default=0.0)",
}

# Kanoniczna kolejność statusów (dźwięków oraz mowy)
CANONICAL_STATUS_ORDER = (
	"unread",
	"flag",
	"attachment",
	"inline_image",
	"replied",
	"forwarded",
	"importance_high",
	"importance_low",
	"meeting",
	"category",
	"external",
)

# Sekwencja wszystkich dźwięków do pełnego testu głośności przy suwaku
TEST_SOUNDS_SEQUENCE = (
	"unread",
	"flag",
	"attachment",
	"inline_image",
	"replied",
	"forwarded",
	"importance_high",
	"importance_low",
	"meeting",
	"category",
	"external",
)

def _to_bool(val, default=True):
	"""
	Bezpieczna konwersja dowolnej wartości (bool, str, int) na typ bool.
	Chroni przed błędem wxPython: CheckBox.SetValue(): argument 1 has unexpected type 'str'.
	"""
	if isinstance(val, bool):
		return val
	if isinstance(val, str):
		v = val.strip().lower()
		if v in ("true", "1", "yes", "on", "t", "prawda"):
			return True
		if v in ("false", "0", "no", "off", "f", "fałsz", "falsz"):
			return False
	if isinstance(val, (int, float)):
		return bool(val)
	return bool(default)


def _to_int(val, default=80, min_val=0, max_val=100):
	"""
	Bezpieczna konwersja wartości na int w dopuszczalnym zakresie [min_val, max_val].
	"""
	try:
		iv = int(val)
		return max(min_val, min(max_val, iv))
	except (ValueError, TypeError):
		return default


# ==============================================================================
# OBSŁUGA WIELOJĘZYCZNOŚCI (POLSKI / ENGLISH / AUTOMATYCZNY)
# ==============================================================================

def get_effective_language():
	"""
	Zwraca efektywny kod języka ('pl' lub 'en'):
	- Jeśli użytkownik wybrał w ustawieniach 'pl' -> 'pl'
	- Jeśli użytkownik wybrał w ustawieniach 'en' -> 'en'
	- Jeśli 'auto': sprawdza język NVDA (languageHandler.getLanguage()). Jeśli polski -> 'pl', inaczej 'en'.
	"""
	try:
		if "_activeSettingsPanelRef" in globals() and _activeSettingsPanelRef:
			panel = _activeSettingsPanelRef()
			if panel and hasattr(panel, "languageChoice") and hasattr(panel, "langKeys"):
				idx = panel.languageChoice.GetSelection()
				if 0 <= idx < len(panel.langKeys):
					chosen = panel.langKeys[idx]
					if chosen in ("pl", "en"):
						return chosen
	except Exception:
		pass
	try:
		confLang = config.conf.get(CONFIG_SECTION, {}).get("language", "auto")
	except Exception:
		confLang = "auto"
	if confLang in ("pl", "en"):
		return confLang
	try:
		import languageHandler
		nvdaLang = languageHandler.getLanguage() or ""
		if nvdaLang.lower().startswith("pl"):
			return "pl"
		if nvdaLang.lower() == "windows" or "default" in nvdaLang.lower():
			import locale
			loc = locale.getdefaultlocale()[0] or ""
			if loc.lower().startswith("pl"):
				return "pl"
	except Exception:
		pass
	try:
		import locale
		loc = locale.getdefaultlocale()[0] or ""
		if loc.lower().startswith("pl"):
			return "pl"
	except Exception:
		pass
	return "en"


def get_doc_path():
	"""Zwraca ścieżkę do pliku dokumentacji odpowiedniej dla aktywnego języka."""
	lang = get_effective_language()
	if lang == "en" and os.path.isfile(DOC_PATH_EN):
		return DOC_PATH_EN
	if lang == "pl" and os.path.isfile(DOC_PATH_PL):
		return DOC_PATH_PL
	if os.path.isfile(DOC_PATH):
		return DOC_PATH
	return DOC_PATH_PL if os.path.isfile(DOC_PATH_PL) else DOC_PATH_EN


def showContextHelp(context: str = "settings", parent=None):
	"""
	Wyświetla kontekstową pomoc i spis skrótów klawiszowych dla konkretnego okna
	w natywnym, czytelnym i w pełni dostępnym buforze przeglądania NVDA (ui.browseableMessage).
	"""
	lang = get_effective_language()
	if lang == "en":
		if context == "contact":
			title = "Help: Contact Add-on Author Dialog"
			text = (
				"OUTLOOK SPEECH FILTER — CONTACT DIALOG HELP\n"
				+ "=" * 60 + "\n\n"
				"OPTIONS AVAILABLE IN CONTACT DIALOG:\n"
				"- Open feedback form in web browser:\n"
				"  Opens a 100% anonymous Google Form. No login required, no email addresses\n"
				"  or personal data collected (fully GDPR compliant).\n"
				"- Copy report to clipboard:\n"
				"  Copies a complete diagnostic report to your clipboard containing NVDA version,\n"
				"  Windows version, add-on settings, and Outlook integration state.\n"
				"  You can paste it (Ctrl+V) into the feedback form or forum message.\n"
				"- View report:\n"
				"  Opens an accessible viewer showing the full contents of the generated technical report.\n"
				"- Save report to Desktop:\n"
				"  Saves the report as a text file (.txt) directly onto your Windows Desktop.\n\n"
				"KEYBOARD SHORTCUTS:\n"
				"  F1              - Show this contextual help\n"
				"  Escape / Alt+F4 - Close contact dialog\n"
			)
		elif context == "thread":
			title = "Help: Outlook Email Thread Navigator"
			text = (
				"OUTLOOK EMAIL THREAD — CONVERSATION STRUCTURE HELP\n"
				+ "=" * 60 + "\n\n"
				"NAVIGATION IN THE THREAD DIALOG:\n"
				"- Down / Up Arrow: move between individual messages in the thread\n"
				"- Enter: jump immediately to selected message in Outlook reading window\n"
				"- Alt+K: copy clean body of selected message to clipboard (without quoted history)\n"
				"- Typing in search box: instant real-time filtering by sender, date, or body snippet\n"
				"- Escape: close dialog without changing position\n"
				"- F1: display this help again\n\n"
				"KEYBOARD SHORTCUTS WHILE READING EMAIL:\n"
				"  NVDA+F7 / NVDA+Alt+F7      - Open this thread structure dialog\n"
				"  NVDA+Alt+Down Arrow        - Jump to next (older) message in thread\n"
				"  NVDA+Alt+Up Arrow          - Jump to previous (newer) message in thread\n"
				"  NVDA+Alt+Home              - Jump to newest message (top of thread)\n"
				"  NVDA+Alt+End               - Jump to oldest message (first inquiry)\n"
			)
		elif context in ("outlook", "messages", "list"):
			title = "Help: Outlook Speech Filter — Outlook Usage"
			text = (
				"OUTLOOK SPEECH FILTER — OUTLOOK USAGE HELP\n"
				+ "=" * 60 + "\n\n"
				"HOW FILTERING WORKS IN THE MAILBOX:\n"
				"- While moving through emails with Arrow keys, the add-on eliminates repetitive\n"
				"  speech prefixes ('unread, flagged, has attachment...').\n"
				"- NVDA immediately announces essential details: sender, subject, and date/time.\n"
				"- Muted attributes are signaled instantly in the background via subtle,\n"
				"  pleasant earcon sounds.\n"
				"- In grouped views (e.g. grouped by flag or category), redundant phrases\n"
				"  ('Group by: Collapsed: Categories...') are stripped, speaking the flag color\n"
				"  or category name first.\n\n"
				"EMAIL THREAD NAVIGATOR (IN EMAIL BODY & ON MESSAGE LIST):\n"
				"  NVDA+F7 / NVDA+Alt+F7      - Open email thread structure list (jump with Enter, copy with Alt+K)\n"
				"  NVDA+Alt+Down Arrow        - Jump to next (older) email in conversation thread\n"
				"  NVDA+Alt+Up Arrow          - Jump to previous (newer) email in conversation thread\n"
				"  NVDA+Alt+Home              - Jump to newest email (top of thread)\n"
				"  NVDA+Alt+End               - Jump to oldest email (origin of discussion)\n\n"
				"OTHER SHORTCUTS IN OUTLOOK:\n"
				"  NVDA+Shift+Z / NVDA+Ctrl+Shift+Z - Speak all statuses of current message on demand\n"
				"  NVDA+Ctrl+Shift+O               - Quickly toggle speech muting on or off\n"
				"  NVDA+Ctrl+Shift+S               - Quickly toggle earcon sound playback on or off\n"
				"  NVDA+Ctrl+Shift+H               - Show this contextual help and shortcuts reference\n\n"
				"EARCON SOUND PLAYBACK ORDER & CORRESPONDING SPEECH:\n"
				"   1. Unread status        - clear, high ping chime (880 Hz)\n"
				"   2. Message flag         - percussive pin tap (1500 Hz)\n"
				"   3. File attachment      - paperclip snap / marimba\n"
				"   4. Inline graphic / pic - camera shutter / vibraphone\n"
				"   5. Replied              - descending two-tone\n"
				"   6. Forwarded            - ascending three-tone\n"
				"   7. High importance      - two sharp alert tones (1000 Hz)\n"
				"   8. Low importance       - soft low tone (280 Hz)\n"
				"   9. Meeting request      - harmonious calendar chord (C5)\n"
				"  10. Message category     - vibrant mark (700 Hz)\n"
				"  11. External message     - double low warning tone\n\n"
				"Note: All shortcuts work with either Insert or Caps Lock as NVDA modifier key.\n"
				"You can customize them in: NVDA Menu -> Preferences -> Input Gestures.\n"
			)
		else:  # settings
			title = "Help: Outlook Speech Filter Settings"
			text = (
				"OUTLOOK SPEECH FILTER — SETTINGS HELP\n"
				+ "=" * 60 + "\n\n"
				"AVAILABLE OPTIONS & SETTINGS:\n"
				"- Interface Language:\n"
				"  Choose add-on language: Automatic (matches NVDA/Windows), Polish, English.\n"
				"  Switching language instantly updates all labels in the open settings panel.\n"
				"- Mute unwanted message status announcements:\n"
				"  Mutes verbose prefixes on the message list, allowing NVDA to read\n"
				"  sender, subject, and date/time first.\n"
				"- Play subtle earcons instead of muted speech announcements:\n"
				"  Master switch for audio cue playback. Replaces muted speech messages\n"
				"  with brief, intuitive earcons in the background.\n\n"
				"SOUND THEMES:\n"
				"- Associative (Default): Realistic everyday sounds (paperclip click for attachment,\n"
				"  camera shutter for inline graphic, chime for unread message).\n"
				"- Minimalist: Modern, subtle synthesized tones of precise frequencies.\n"
				"- Instrumental: Pleasant acoustic instruments (marimba, vibraphone, chimes).\n\n"
				"SOUND CHECKLIST & VOLUME CONTROL:\n"
				"- Bulk buttons: 'Select all sounds' and 'Deselect all sounds'.\n"
				"  Deselecting all sounds automatically collapses sound playback, and re-checking\n"
				"  the master box instantly restores previous selections!\n"
				"- Individual sounds: Each status has its own checkbox and full-volume test button.\n"
				"- Volume slider: Adjust from 0% to 100%. Use Up/Down arrows (1%), PageUp/PageDown (10%),\n"
				"  Home (100%), End (0%). Changes apply in real time — test immediately via Alt+Tab.\n"
				"- 'Test volume' button: Sequentially plays all sounds at current slider volume.\n\n"
				"UTILITY BUTTONS:\n"
				"- Contextual Help (F1): Displays this help buffer.\n"
				"- User Guide / Tutorial: Opens full HTML documentation in default web browser.\n"
				"- Contact add-on author: Opens anonymous feedback dialog and diagnostic report.\n"
				"- Configure shortcut gestures: Opens NVDA's Input Gestures dialog.\n"
				"- Restore default settings: Safely resets configuration to factory defaults.\n\n"
				"GLOBAL KEYBOARD SHORTCUTS IN OUTLOOK:\n"
				"  NVDA+Ctrl+Shift+O               - Toggle speech muting on and off\n"
				"  NVDA+Ctrl+Shift+S               - Toggle earcon sounds on and off\n"
				"  NVDA+Shift+Z / NVDA+Ctrl+Shift+Z - Speak all statuses of current message on demand\n"
				"  NVDA+Ctrl+Shift+H               - Show contextual help buffer\n"
				"  F1                              - Show contextual help in settings and dialogs\n"
			)
	else:
		if context == "contact":
			title = "Pomoc: Kontakt z autorem dodatku"
			text = (
				"OUTLOOK FILTR MOWY — POMOC OKNA KONTAKTU\n"
				+ "=" * 60 + "\n\n"
				"OPCJE DOSTĘPNE W OKNIE KONTAKTU:\n"
				"- Otwórz formularz zgłoszeniowy w przeglądarce:\n"
				"  Otwiera w 100% anonimowy formularz online (Google Forms). Formularz nie wymaga\n"
				"  logowania, nie gromadzi adresów e-mail ani żadnych danych osobowych (RODO).\n"
				"- Kopiuj raport do schowka:\n"
				"  Kopiuje do schowka kompletny raport diagnostyczny dodatku zawierający wersję NVDA,\n"
				"  wersję systemu Windows, konfigurację dodatku i stan integracji z Outlookiem.\n"
				"  Możesz wkleić ten raport (Ctrl+V) w formularzu zgłoszeniowym lub w wiadomości.\n"
				"- Wyświetl podgląd raportu:\n"
				"  Otwiera czytelne okno podglądu z pełną treścią wygenerowanego raportu technicznego.\n"
				"- Zapisz raport na Pulpicie:\n"
				"  Zapisuje raport jako plik tekstowy (.txt) bezpośrednio na Twoim Pulpicie Windows.\n\n"
				"SKRÓTY KLAWISZOWE:\n"
				"  F1              - Wyświetl tę pomoc kontekstową\n"
				"  Escape / Alt+F4 - Zamknij okno kontaktu\n"
			)
		elif context == "thread":
			title = "Pomoc: Nawigator wątków wiadomości Outlook"
			text = (
				"WĄTEK WIADOMOŚCI OUTLOOK — STRUKTURA ROZMOWY\n"
				+ "=" * 60 + "\n\n"
				"KLAWISZE NAWIGACYJNE W TYM OKNIE:\n"
				"- Strzałka w dół / w górę: poruszanie się po poszczególnych mailach w wątku\n"
				"- Enter: natychmiastowe przejście do wybranego maila w oknie wiadomości Outlooka\n"
				"- Alt+K: skopiowanie do schowka wyłącznie treści wybranego maila (bez cytatów i historii)\n"
				"- Wpisanie tekstu w polu wyszukiwania: natychmiastowe filtrowanie po nadawcy, dacie lub treści\n"
				"- Escape: zamknięcie okna bez zmiany pozycji kursora\n"
				"- F1: ponowne wyświetlenie tej pomocy\n\n"
				"SKRÓTY KLAWISZOWE PODCZAS CZYTANIA MAILA:\n"
				"  NVDA+F7 / NVDA+Alt+F7      - Otwarcie tego okna struktury wątku wiadomości\n"
				"  NVDA+Alt+Strzałka w dół    - Skok do następnego (starszego) maila w wątku\n"
				"  NVDA+Alt+Strzałka w górę   - Skok do poprzedniego (nowszego) maila w wątku\n"
				"  NVDA+Alt+Home              - Skok do najnowszego maila (góra wątku)\n"
				"  NVDA+Alt+End               - Skok do najstarszego maila (początek korespondencji)\n"
			)
		elif context in ("outlook", "messages", "list"):
			title = "Pomoc: Outlook Filtr Mowy — Praca w programie Outlook"
			text = (
				"OUTLOOK FILTR MOWY — POMOC W PROGRAMIE OUTLOOK\n"
				+ "=" * 60 + "\n\n"
				"JAK DZIAŁA FILTROWANIE W SKRZYNCE POCZTOWEJ:\n"
				"- Podczas nawigacji strzałkami po liście wiadomości, dodatek eliminuje powtarzające\n"
				"  się komunikaty mowy („nieprzeczytana, z flagą, ma załącznik...”).\n"
				"- NVDA od razu wypowiada najważniejsze informacje: nadawcę, temat oraz datę/godzinę.\n"
				"- Wyciszone cechy wiadomości są natychmiast sygnalizowane krótkimi, przyjemnymi\n"
				"  dźwiękami (earconami) odtwarzanymi w tle.\n"
				"- W widokach z włączonym grupowaniem (np. według flag lub kategorii), dodatek\n"
				"  eliminuje tasiemcowe frazy („Grupuj według: Zwinięte: Kategorie...”) i podaje\n"
				"  kolor flagi lub kategorię jako pierwszy element.\n\n"
				"NAWIGATOR WĄTKÓW WIADOMOŚCI (W TREŚCI MAILA I NA LIŚCIE):\n"
				"  NVDA+F7 / NVDA+Alt+F7      - Otwórz okno listy maili w wątku (przejście Enter, kopiowanie Alt+K)\n"
				"  NVDA+Alt+Strzałka w dół    - Skok do następnego (starszego) maila w wątku korespondencji\n"
				"  NVDA+Alt+Strzałka w górę   - Skok do poprzedniego (nowszego) maila w wątku korespondencji\n"
				"  NVDA+Alt+Home              - Skok do najnowszego maila (góra wątku)\n"
				"  NVDA+Alt+End               - Skok do najstarszego maila (początek dyskusji)\n\n"
				"POZOSTAŁE SKRÓTY KLAWISZOWE W OUTLOOKU:\n"
				"  NVDA+Shift+Z / NVDA+Ctrl+Shift+Z - Odczytaj na żądanie wszystkie statusy wiadomości\n"
				"  NVDA+Ctrl+Shift+O               - Szybkie włączenie / wyłączenie filtrowania mowy\n"
				"  NVDA+Ctrl+Shift+S               - Szybkie włączenie / wyłączenie dźwięków statusów\n"
				"  NVDA+Ctrl+Shift+H               - Wyświetl tę pomoc kontekstową i spis skrótów\n\n"
				"KOLEJNOŚĆ ODTWARZANIA DŹWIĘKÓW I ODPOWIADAJĄCA IM MOWA:\n"
				"   1. Nieprzeczytana       - czysty, wysoki dzwonek (880 Hz)\n"
				"   2. Flaga wiadomości     - krótkie perkusyjne kliknięcie (1500 Hz)\n"
				"   3. Załącznik plikowy    - pstryknięcie spinacza biurowego / marimba\n"
				"   4. Wklejona grafika     - migawka aparatu fotograficznego / wibrafon\n"
				"   5. Odpowiedziano        - dwuton opadający powrotny\n"
				"   6. Przesłano dalej      - trzyton wznoszący w świat\n"
				"   7. Wysoki priorytet     - dwa szybkie ostre tony alarmowe (1000 Hz)\n"
				"   8. Niski priorytet      - spokojny niski pomruk (280 Hz)\n"
				"   9. Żądanie spotkania    - harmonijny dzwonek terminarza (akord C5)\n"
				"  10. Kategoria wiadomości - wibrujący znacznik kolorowej etykiety (700 Hz)\n"
				"  11. Wiadomość zewnętrzna - podwójny niski ton ostrzegawczy\n\n"
				"Uwaga: Wszystkie skróty działają zarówno z klawiszem Insert, jak i Caps Lock\n"
				"jako klawiszem NVDA. Możesz je dostosować w Menu NVDA -> Preferencje -> Zdarzenia wejścia.\n"
			)
		else:  # settings
			title = "Pomoc: Ustawienia dodatku Outlook Speech Filter"
			text = (
				"OUTLOOK FILTR MOWY — POMOC USTAWIEŃ DODATKU\n"
				+ "=" * 60 + "\n\n"
				"GŁÓWNE OPCJE I PRZEŁĄCZNIKI:\n"
				"- Język interfejsu dodatku:\n"
				"  Wybór języka (Automatyczny, Polski, English). Zmiana języka natychmiast\n"
				"  aktualizuje wszystkie etykiety, przyciski i opisy w otwartym oknie.\n"
				"- Włącz filtrowanie mowy o wiadomościach:\n"
				"  Wycisza uciążliwe prefiksy mowy na liście maili w programie Outlook,\n"
				"  pozwalając NVDA czytać od razu nadawcę, temat oraz datę/godzinę.\n"
				"- Odtwarzaj dyskretne dźwięki zamiast wyciszonych komunikatów mowy:\n"
				"  Główny przełącznik sygnałów dźwiękowych (earconów). Zastępuje wyciszone\n"
				"  komunikaty mowy krótkimi, intuicyjnymi dźwiękami w tle.\n\n"
				"MOTYWY DŹWIĘKOWE:\n"
				"- Skojarzeniowy (Domyślny): Realistyczne dźwięki przedmiotów (pstryknięcie spinacza\n"
				"  dla załącznika, migawka aparatu dla wklejonej grafiki, dzwonek dla nowej wiadomości).\n"
				"- Minimalistyczny: Nowoczesne, czyste syntezy tonowe o ściśle dobranych częstotliwościach.\n"
				"- Instrumentalny: Przyjemne, organiczne instrumenty akustyczne (marimba, wibrafon, dzwonki).\n\n"
				"LISTA DŹWIĘKÓW I REGULACJA GŁOŚNOŚCI:\n"
				"- Przyciski zbiorcze: 'Zaznacz wszystkie dźwięki' oraz 'Odznacz wszystkie dźwięki'.\n"
				"  Odznaczenie wszystkich dźwięków automatycznie zwija sekcję dźwiękową,\n"
				"  a ponowne zaznaczenie opcji nadrzędnej natychmiast przywraca poprzedni zestaw!\n"
				"- Niezależne dźwięki: Każdy status ma własne pole wyboru i przycisk odsłuchu 100%.\n"
				"- Suwak głośności: Precyzyjna regulacja (0-100%). Obsługa klawiaturą:\n"
				"  Strzałki (1%), PageUp/PageDown (10%), Home (100%), End (0%).\n"
				"  Działa w czasie rzeczywistym — możesz testować głośność w Outlooku bez klikania OK!\n"
				"- Przycisk 'Przetestuj głośność': Odtwarza sekwencyjnie wszystkie dźwięki dodatku.\n\n"
				"PRZYCISKI NARZĘDZIOWE:\n"
				"- Pomoc kontekstowa (F1): Wyświetla niniejsze okno pomocy.\n"
				"- Instrukcja obsługi / Samouczek: Otwiera pełną dokumentację HTML w przeglądarce.\n"
				"- Kontakt z autorem: Otwiera anonimowy formularz zgłoszeniowy i generator raportu.\n"
				"- Skonfiguruj skróty klawiszowe: Przechodzi do konfiguracji gestów w Zdarzeniach wejścia NVDA.\n"
				"- Przywróć ustawienia domyślne: Bezpiecznie przywraca fabryczną konfigurację dodatku.\n\n"
				"GLOBALNE SKRÓTY KLAWISZOWE W PROGRAMIE OUTLOOK:\n"
				"  NVDA+Ctrl+Shift+O               - Szybkie włączenie / wyłączenie filtrowania mowy\n"
				"  NVDA+Ctrl+Shift+S               - Szybkie włączenie / wyłączenie dźwięków statusów\n"
				"  NVDA+Shift+Z / NVDA+Ctrl+Shift+Z - Odczytanie mową na żądanie wszystkich statusów\n"
				"  NVDA+Ctrl+Shift+H               - Wyświetlenie pomocy kontekstowej dodatku\n"
				"  F1                              - Wyświetlenie tej pomocy kontekstowej w oknie ustawień i kontaktu\n"
			)
	ui.browseableMessage(text, title=title)


TRANSLATIONS = {
	"TITLE": {
		"pl": "Outlook - Filtr mowy (Outlook - Speech Filter)",
		"en": "Outlook - Speech Filter",
	},
	"LANG_LABEL": {
		"pl": "Język interfejsu dodatku:",
		"en": "Add-on interface language:",
	},
	"LANG_AUTO": {
		"pl": "Automatyczny (zgodny z NVDA i systemem)",
		"en": "Automatic (match NVDA and system)",
	},
	"LANG_PL": {
		"pl": "Polski (Polish)",
		"en": "Polski (Polish)",
	},
	"LANG_EN": {
		"pl": "English (angielski)",
		"en": "English",
	},
	"LANG_CHANGED": {
		"pl": "Język dodatku zmieniony na: {lang}",
		"en": "Add-on language changed to: {lang}",
	},
	"ENABLE_FILTER": {
		"pl": "Włącz wyciszanie zbędnych informacji o wiadomościach",
		"en": "Mute unwanted message status announcements",
	},
	"PARENT_SOUNDS_LABEL": {
		"pl": "Odtwarzaj dyskretne dźwięki zamiast wyciszonych komunikatów mowy",
		"en": "Play subtle earcons instead of muted speech announcements",
	},
	"SOUND_EXPL_TEXT": {
		"pl": "Poniżej możesz włączyć lub wyłączyć odtwarzanie dźwięku dla każdego zdarzenia z osobna oraz odsłuchać jego brzmienie:",
		"en": "Below you can enable or disable earcons for each event individually and preview their sound:",
	},
	"SELECT_ALL_SOUNDS": {
		"pl": "Zaznacz wszystkie dźwięki",
		"en": "Select all sounds",
	},
	"DESELECT_ALL_SOUNDS": {
		"pl": "Odznacz wszystkie dźwięki",
		"en": "Deselect all sounds",
	},
	"SOUND_UNREAD": {
		"pl": "Dźwięk stanu nieprzeczytania",
		"en": "Unread message status sound",
	},
	"PLAY_UNREAD": {
		"pl": "Odtwórz dźwięk: Nieprzeczytana",
		"en": "Preview sound: Unread",
	},
	"SOUND_FLAG": {
		"pl": "Dźwięk flagi wiadomości (oflagowana)",
		"en": "Flagged message sound (flagged)",
	},
	"PLAY_FLAG": {
		"pl": "Odtwórz dźwięk: Flaga",
		"en": "Preview sound: Flag",
	},
	"SOUND_ATTACHMENT": {
		"pl": "Dźwięk załącznika (ma załącznik)",
		"en": "Attachment sound (has attachment)",
	},
	"PLAY_ATTACHMENT": {
		"pl": "Odtwórz dźwięk: Załącznik",
		"en": "Preview sound: Attachment",
	},
	"SOUND_INLINE_IMAGE": {
		"pl": "Dźwięk wklejonej grafiki / zrzutu ekranu w treści",
		"en": "Inline graphic / screenshot sound",
	},
	"PLAY_INLINE_IMAGE": {
		"pl": "Odtwórz dźwięk: Wklejona grafika",
		"en": "Preview sound: Inline graphic",
	},
	"SOUND_THEME_LABEL": {
		"pl": "Motyw dźwięków statusów:",
		"en": "Status sounds theme:",
	},
	"THEME_ASSOCIATIVE": {
		"pl": "Skojarzeniowe (spinacz, migawka aparatu, dzwonek) — Domyślny",
		"en": "Associative (paperclip, shutter, chime) — Default",
	},
	"THEME_MINIMAL": {
		"pl": "Minimalistyczne (tony pojedyncze i podwójne)",
		"en": "Minimalist (single and double electronic tones)",
	},
	"THEME_INSTRUMENTAL": {
		"pl": "Instrumentalne (marimba, wibrafon, dzwonki)",
		"en": "Instrumental (marimba, vibraphone, bells)",
	},
	"SOUND_REPLIED": {
		"pl": "Dźwięk odpowiedzi na wiadomość (odpowiedziano)",
		"en": "Replied message sound (replied)",
	},
	"PLAY_REPLIED": {
		"pl": "Odtwórz dźwięk: Odpowiedź",
		"en": "Preview sound: Replied",
	},
	"SOUND_FORWARDED": {
		"pl": "Dźwięk przesłania wiadomości dalej (przesłano dalej)",
		"en": "Forwarded message sound (forwarded)",
	},
	"PLAY_FORWARDED": {
		"pl": "Odtwórz dźwięk: Przesłano dalej",
		"en": "Preview sound: Forwarded",
	},
	"SOUND_IMPORTANCE_HIGH": {
		"pl": "Dźwięk wysokiego priorytetu wiadomości",
		"en": "High importance sound",
	},
	"PLAY_IMPORTANCE_HIGH": {
		"pl": "Odtwórz dźwięk: Wysoki priorytet",
		"en": "Preview sound: High importance",
	},
	"SOUND_IMPORTANCE_LOW": {
		"pl": "Dźwięk niskiego priorytetu wiadomości",
		"en": "Low importance sound",
	},
	"PLAY_IMPORTANCE_LOW": {
		"pl": "Odtwórz dźwięk: Niski priorytet",
		"en": "Preview sound: Low importance",
	},
	"SOUND_MEETING": {
		"pl": "Dźwięk żądania spotkania (terminarz)",
		"en": "Meeting request sound (calendar)",
	},
	"PLAY_MEETING": {
		"pl": "Odtwórz dźwięk: Spotkanie",
		"en": "Preview sound: Meeting request",
	},
	"SOUND_CATEGORY": {
		"pl": "Dźwięk kategorii wiadomości",
		"en": "Message category sound",
	},
	"PLAY_CATEGORY": {
		"pl": "Odtwórz dźwięk: Kategoria",
		"en": "Preview sound: Category",
	},
	"SOUND_EXTERNAL": {
		"pl": "Dźwięk wiadomości zewnętrznej (np. 'zewnętrzna')",
		"en": "External message sound (e.g. 'external')",
	},
	"PLAY_EXTERNAL": {
		"pl": "Odtwórz dźwięk: Zewnętrzna",
		"en": "Preview sound: External",
	},
	"VOL_LABEL": {
		"pl": "Głośność dźwięków (%):",
		"en": "Sound volume (%):",
	},
	"SLIDER_NAME": {
		"pl": "Głośność dźwięków",
		"en": "Sound volume",
	},
	"TEST_VOL_BTN": {
		"pl": "Przetestuj głośność",
		"en": "Test volume",
	},
	"HELP_BTN": {
		"pl": "Otwórz pomoc i samouczek dźwięków...",
		"en": "Open Help and Sound Tutorial...",
	},
	"CONTEXT_HELP_BTN": {
		"pl": "Pomoc kontekstowa (F1)...",
		"en": "Contextual Help (F1)...",
	},
	"BTN_HELP_CONTEXT": {
		"pl": "Pomoc (F1)...",
		"en": "Help (F1)...",
	},
	"GESTURES_BTN": {
		"pl": "Skonfiguruj skróty klawiszowe...",
		"en": "Configure shortcut gestures...",
	},
	"CONTACT_BTN": {
		"pl": "Kontakt z autorem dodatku...",
		"en": "Contact add-on author...",
	},
	"RESTORE_DEFAULTS_BTN": {
		"pl": "Przywróć ustawienia domyślne...",
		"en": "Restore default settings...",
	},
	"RESTORE_CONFIRM_TITLE": {
		"pl": "Przywracanie ustawień domyślnych",
		"en": "Restore default settings",
	},
	"RESTORE_CONFIRM_MSG": {
		"pl": "Czy na pewno chcesz przywrócić wszystkie ustawienia dodatku Outlook Speech Filter do wartości domyślnych?",
		"en": "Are you sure you want to restore all Outlook Speech Filter settings to factory defaults?",
	},
	"CONTACT_MENU_BUG": {
		"pl": "Błąd w działaniu dodatku (zgłoszenie problemu)...",
		"en": "Bug report (report an issue)...",
	},
	"CONTACT_MENU_LANG": {
		"pl": "Prośba o nową wersję językową dodatku...",
		"en": "Request a new language translation...",
	},
	"CONTACT_MENU_FEATURE": {
		"pl": "Pomysły na dodatkowe funkcje (sugestia)...",
		"en": "Feature request (ideas / suggestions)...",
	},
	"CONTACT_MENU_OTHER": {
		"pl": "Inne zapytanie lub opinia o dodatku...",
		"en": "Other question or feedback...",
	},
	"CONTACT_TOPIC_BUG": {
		"pl": "Błąd w działaniu dodatku",
		"en": "Bug report",
	},
	"CONTACT_TOPIC_LANG": {
		"pl": "Prośba o nową wersję językową",
		"en": "Translation request",
	},
	"CONTACT_TOPIC_FEATURE": {
		"pl": "Pomysły na dodatkowe funkcje",
		"en": "Feature request",
	},
	"CONTACT_TOPIC_OTHER": {
		"pl": "Inne zapytanie lub opinia",
		"en": "Other inquiry or feedback",
	},
	"PARENT_ALL_DISABLED": {
		"pl": " (wszystkie wyłączone)",
		"en": " (all disabled)",
	},
	"PARENT_PARTIAL": {
		"pl": " (częściowo wybrane)",
		"en": " (partially selected)",
	},
	"SOUNDS_ALL_SELECTED": {
		"pl": "Wszystkie dźwięki zaznaczone",
		"en": "All sounds selected",
	},
	"SOUNDS_ALL_DESELECTED": {
		"pl": "Wszystkie dźwięki zostały odznaczone. Wyłączono opcję odtwarzania dźwięków.",
		"en": "All sounds deselected. Sound playback option disabled.",
	},
	"SOUNDS_VOL_MUTED": {
		"pl": "Głośność 0% (dźwięki wyciszone)",
		"en": "Volume 0% (sounds muted)",
	},
	"SOUNDS_STATE_RESTORED": {
		"pl": "Przywrócono stan zaznaczenia dźwięków",
		"en": "Sound selection state restored",
	},
	"HELP_NOT_FOUND": {
		"pl": "Plik pomocy nie został znaleziony.",
		"en": "Help file not found.",
	},
	"DEFAULTS_RESTORED": {
		"pl": "Przywrócono domyślne ustawienia dodatku",
		"en": "Default add-on settings restored",
	},
	"PLAY_GENERIC": {
		"pl": "Odtwórz dźwięk",
		"en": "Preview sound",
	},
	"CONTACT_DLG_TITLE": {
		"pl": "Kontakt z autorem i zgłoszenia: Sławomir Strugarek",
		"en": "Contact Author & Feedback: Sławomir Strugarek",
	},
	"CONTACT_DLG_INFO": {
		"pl": "Zgłoszenia i opinie obsługiwane są za pośrednictwem bezpiecznego formularza online (Google Forms).\nFormularz jest w 100% anonimowy – nie gromadzi adresów e-mail ani żadnych danych osobowych (pełna zgodność z RODO).",
		"en": "Feedback and bug reports are handled via a secure online form (Google Forms).\nThe form is 100% anonymous — it does not collect email addresses or any personal data (fully GDPR compliant).",
	},
	"CONTACT_TOPIC_LABEL": {
		"pl": "Kategoria zgłoszenia:",
		"en": "Report category:",
	},
	"CONTACT_BTN_OPEN_FORM": {
		"pl": "Otwórz formularz zgłoszeniowy w przeglądarce...",
		"en": "Open feedback form in web browser...",
	},
	"CONTACT_REPORT_SECTION_LABEL": {
		"pl": "Raport diagnostyczny (zalecany przy zgłaszaniu problemów technicznych):",
		"en": "Diagnostic report (recommended for technical troubleshooting):",
	},
	"CONTACT_BTN_COPY_REPORT": {
		"pl": "Kopiuj raport do schowka",
		"en": "Copy report to clipboard",
	},
	"CONTACT_REPORT_COPIED_MSG": {
		"pl": "Raport diagnostyczny został pomyślnie skopiowany do schowka. Możesz go teraz wkleić w formularzu (skrótem Ctrl+V).",
		"en": "Diagnostic report has been successfully copied to clipboard. You can now paste it into the form (Ctrl+V).",
	},
	"CONTACT_BTN_VIEW_REPORT": {
		"pl": "Podgląd raportu...",
		"en": "View report...",
	},
	"CONTACT_BTN_SAVE_REPORT": {
		"pl": "Zapisz raport na Pulpicie...",
		"en": "Save report to Desktop...",
	},
	"CONTACT_REPORT_SAVED_MSG": {
		"pl": "Raport diagnostyczny został pomyślnie zapisany na Pulpicie:\n{path}",
		"en": "Diagnostic report has been successfully saved to your Desktop:\n{path}",
	},
	"CONTACT_DESKTOP_SAVED_TITLE": {
		"pl": "Zapisano na Pulpicie",
		"en": "Saved to Desktop",
	},
	"REPORT_VIEW_TITLE": {
		"pl": "Podgląd raportu diagnostycznego",
		"en": "Diagnostic Report Preview",
	},
	"BTN_CLOSE": {
		"pl": "Zamknij",
		"en": "Close",
	},
	"STATUS_UNREAD": {
		"pl": "nieprzeczytana",
		"en": "unread",
	},
	"STATUS_FLAG": {
		"pl": "oflagowana",
		"en": "flagged",
	},
	"STATUS_ATTACHMENT_SINGLE": {
		"pl": "załącznik",
		"en": "attachment",
	},
	"STATUS_ATTACHMENTS_MULTIPLE": {
		"pl": "załączniki: {count}",
		"en": "attachments: {count}",
	},
	"STATUS_INLINE_IMAGE": {
		"pl": "wklejona grafika",
		"en": "inline graphic",
	},
	"STATUS_INLINE_IMAGES_MULTIPLE": {
		"pl": "wklejone grafiki: {count}",
		"en": "inline graphics: {count}",
	},
	"REPORT_INLINE_IMAGE_CHECK": {
		"pl": "Wklejone grafiki i zrzuty ekranu w treści wiadomości",
		"en": "Inline graphics and screenshots in message body",
	},
	"STATUS_REPLIED": {
		"pl": "odpowiedziano",
		"en": "replied",
	},
	"STATUS_REPLIED_ALL": {
		"pl": "odpowiedziano wszystkim",
		"en": "replied to all",
	},
	"STATUS_FORWARDED": {
		"pl": "przesłano dalej",
		"en": "forwarded",
	},
	"STATUS_IMPORTANCE_HIGH": {
		"pl": "wysoki priorytet",
		"en": "high importance",
	},
	"STATUS_IMPORTANCE_LOW": {
		"pl": "niski priorytet",
		"en": "low importance",
	},
	"STATUS_MEETING": {
		"pl": "żądanie spotkania",
		"en": "meeting request",
	},
	"STATUS_CATEGORY": {
		"pl": "kategorie: {cat}",
		"en": "categories: {cat}",
	},
	"STATUS_EXTERNAL": {
		"pl": "wiadomość zewnętrzna",
		"en": "external message",
	},
	"STATUS_NO_ADDITIONAL": {
		"pl": "Wiadomość przeczytana, brak dodatkowych statusów",
		"en": "Message read, no additional statuses",
	},
	"STATUS_NO_SELECTION": {
		"pl": "Nie zaznaczono żadnej wiadomości w Outlooku",
		"en": "No message selected in Outlook",
	},
	"STATUS_NO_ACTIVE_ITEM": {
		"pl": "Brak aktywnego elementu",
		"en": "No active element",
	},
	"STATUS_ONLY_OUTLOOK": {
		"pl": "To polecenie działa tylko w programie Microsoft Outlook",
		"en": "This command works only in Microsoft Outlook",
	},
	"TOGGLE_FILTER_ON": {
		"pl": "Filtrowanie mowy w Outlooku: włączone (komunikaty wyciszone)",
		"en": "Outlook Speech Filter: enabled (speech announcements muted)",
	},
	"TOGGLE_FILTER_OFF": {
		"pl": "Filtrowanie mowy w Outlooku: wyłączone (komunikaty mowy włączone)",
		"en": "Outlook Speech Filter: disabled (speech announcements unmuted)",
	},
	"TOGGLE_SOUNDS_ON": {
		"pl": "Dźwięki statusów w Outlooku: włączone",
		"en": "Outlook status sounds: enabled",
	},
	"TOGGLE_SOUNDS_OFF": {
		"pl": "Dźwięki statusów w Outlooku: wyłączone",
		"en": "Outlook status sounds: disabled",
	},
	"GESTURE_DESC_TOGGLE_FILTER": {
		"pl": "Przełącza filtrowanie mowy o wiadomościach w skrzynce programu Outlook",
		"en": "Toggles message speech filtering in Microsoft Outlook",
	},
	"GESTURE_DESC_TOGGLE_SOUNDS": {
		"pl": "Przełącza odtwarzanie dźwięków statusów w skrzynce programu Outlook",
		"en": "Toggles status sounds in Microsoft Outlook",
	},
	"GESTURE_DESC_REPORT_STATUSES": {
		"pl": "Odczytuje na żądanie wszystkie statusy aktualnej wiadomości w Outlooku",
		"en": "Reports all statuses of the current message in Outlook on demand",
	},
	"GESTURE_DESC_CONTEXT_HELP": {
		"pl": "Wyświetla pomoc kontekstową i spis skrótów klawiszowych dodatku Outlook Speech Filter",
		"en": "Shows contextual help and keyboard shortcuts for Outlook Speech Filter",
	},
	"ADDON_INFO_LINE": {
		"pl": "{name} wersja {ver} (kompilacja {build} z dnia {date})\nAutorzy: {author}",
		"en": "{name} version {ver} (build {build}, {date})\nAuthors: {author}",
	},
	"THREAD_DIALOG_TITLE": {
		"pl": "Wątek rozmowy — lista wiadomości (NVDA+F7)",
		"en": "Conversation Thread — Message List (NVDA+F7)",
	},
	"THREAD_FILTER_LABEL": {
		"pl": "&Filtruj wiadomości w wątku (wpisz nazwisko, treść lub datę):",
		"en": "&Filter messages in thread (type name, body, or date):",
	},
	"THREAD_FILTER_HINT": {
		"pl": "Wpisz np. nazwisko, słowo z treści lub datę...",
		"en": "Type sender name, body snippet, or date...",
	},
	"THREAD_LIST_LABEL": {
		"pl": "&Wiadomości w wątku rozmowy (od najnowszej do najstarszej):",
		"en": "&Messages in conversation thread (newest to oldest):",
	},
	"THREAD_BTN_GO": {
		"pl": "&Przejdź do wiadomości (Enter)",
		"en": "&Jump to message (Enter)",
	},
	"THREAD_BTN_COPY": {
		"pl": "&Kopiuj treść wiadomości (Alt+K)",
		"en": "&Copy message body (Alt+K)",
	},
	"THREAD_NO_MESSAGES_OR_DOC": {
		"pl": "Nie udało się odczytać treści wiadomości z programu Outlook.",
		"en": "Could not read message content from Outlook.",
	},
	"THREAD_ONLY_ONE_MESSAGE": {
		"pl": "Ta wiadomość nie zawiera wcześniejszego wątku rozmowy (tylko 1 wiadomość).",
		"en": "This email does not contain a previous conversation thread (single message only).",
	},
	"THREAD_LAST_MESSAGE": {
		"pl": "To jest najstarsza wiadomość w wątku.",
		"en": "This is the oldest message in the thread.",
	},
	"THREAD_FIRST_MESSAGE": {
		"pl": "To jest najnowsza wiadomość w wątku.",
		"en": "This is the newest message in the thread.",
	},
	"THREAD_CANNOT_NAVIGATE": {
		"pl": "Nie można przejść do wiadomości — brak dostępu do okna edytora.",
		"en": "Cannot navigate to message — no access to editor window.",
	},
	"THREAD_NAV_ERROR": {
		"pl": "Błąd podczas nawigacji do wiadomości.",
		"en": "Error while navigating to message.",
	},
	"THREAD_UNKNOWN_SENDER": {
		"pl": "Nieznany nadawca",
		"en": "Unknown sender",
	},
	"THREAD_TOP_MESSAGE": {
		"pl": "Bieżąca wiadomość (najnowsza)",
		"en": "Current message (newest)",
	},
	"THREAD_MSG_ANNOUNCE": {
		"pl": "Wiadomość {index} z {total}: {sender}{date}",
		"en": "Message {index} of {total}: {sender}{date}",
	},
	"THREAD_COPIED_TO_CLIPBOARD": {
		"pl": "Skopiowano treść wiadomości ({sender}) do schowka.",
		"en": "Copied message body ({sender}) to clipboard.",
	},
	"BTN_CANCEL": {
		"pl": "&Anuluj (Esc)",
		"en": "&Cancel (Esc)",
	},
	"GESTURE_DESC_THREAD_STRUCTURE": {
		"pl": "Wyświetla listę wiadomości w wątku bieżącej korespondencji w Outlooku (okno struktury wątku)",
		"en": "Displays the list of messages in current Outlook email thread (structure dialog)",
	},
	"GESTURE_DESC_THREAD_NEXT": {
		"pl": "Przechodzi do następnej (starszej) wiadomości w wątku korespondencji w Outlooku",
		"en": "Jumps to the next (older) message in Outlook conversation thread",
	},
	"GESTURE_DESC_THREAD_PREV": {
		"pl": "Przechodzi do poprzedniej (nowszej) wiadomości w wątku korespondencji w Outlooku",
		"en": "Jumps to the previous (newer) message in Outlook conversation thread",
	},
	"GESTURE_DESC_THREAD_FIRST": {
		"pl": "Przechodzi do najnowszej (pierwszej) wiadomości w wątku korespondencji w Outlooku",
		"en": "Jumps to the newest (first) message in Outlook conversation thread",
	},
	"GESTURE_DESC_THREAD_LAST": {
		"pl": "Przechodzi do najstarszej (ostatniej) wiadomości w wątku korespondencji w Outlooku",
		"en": "Jumps to the oldest (last) message in Outlook conversation thread",
	},
	"ENABLE_THREAD_NAV": {
		"pl": "Włącz nawigator wątków wiadomości w programie Outlook (NVDA+F7 oraz NVDA+Alt+Strzałki)",
		"en": "Enable conversation thread navigator in Outlook (NVDA+F7 and NVDA+Alt+Arrows)",
	},
	"TOGGLE_THREAD_NAV_ON": {
		"pl": "Nawigator wątków w Outlooku: włączony",
		"en": "Outlook thread navigator: enabled",
	},
	"TOGGLE_THREAD_NAV_OFF": {
		"pl": "Nawigator wątków w Outlooku: wyłączony",
		"en": "Outlook thread navigator: disabled",
	},
	"THREAD_NAV_DISABLED": {
		"pl": "Nawigator wątków wiadomości jest wyłączony w ustawieniach dodatku.",
		"en": "Conversation thread navigator is disabled in add-on settings.",
	},
	"GESTURE_DESC_TOGGLE_THREAD_NAV": {
		"pl": "Przełącza nawigator wątków wiadomości w programie Outlook",
		"en": "Toggles conversation thread navigator in Microsoft Outlook",
	},
}

REVERSE_MAP = {
	"Włącz wyciszanie zbędnych informacji o wiadomościach": "ENABLE_FILTER",
	"Odtwarzaj dyskretne dźwięki zamiast wyciszonych komunikatów mowy": "PARENT_SOUNDS_LABEL",
	"Poniżej możesz włączyć lub wyłączyć odtwarzanie dźwięku dla każdego zdarzenia z osobna oraz odsłuchać jego brzmienie:": "SOUND_EXPL_TEXT",
	"Zaznacz wszystkie dźwięki": "SELECT_ALL_SOUNDS",
	"Odznacz wszystkie dźwięki": "DESELECT_ALL_SOUNDS",
	"nieprzeczytana": "STATUS_UNREAD",
	"oflagowana": "STATUS_FLAG",
	"załącznik": "STATUS_ATTACHMENT_SINGLE",
	"odpowiedziano": "STATUS_REPLIED",
	"odpowiedziano wszystkim": "STATUS_REPLIED_ALL",
	"przesłano dalej": "STATUS_FORWARDED",
	"wysoki priorytet": "STATUS_IMPORTANCE_HIGH",
	"niski priorytet": "STATUS_IMPORTANCE_LOW",
	"żądanie spotkania": "STATUS_MEETING",
	"wiadomość zewnętrzna": "STATUS_EXTERNAL",
	"Brak aktywnego elementu": "STATUS_NO_ACTIVE_ITEM",
	"To polecenie działa tylko w programie Microsoft Outlook": "STATUS_ONLY_OUTLOOK",
	"Wiadomość przeczytana, brak dodatkowych statusów": "STATUS_NO_ADDITIONAL",
	"Nie zaznaczono żadnej wiadomości w Outlooku": "STATUS_NO_SELECTION",
	"Outlook - Filtr mowy": "TITLE",
	"Głośność dźwięków (%):": "VOL_LABEL",
	"Głośność dźwięków": "SLIDER_NAME",
	"Przetestuj głośność": "TEST_VOL_BTN",
	"Otwórz pomoc i samouczek dźwięków...": "HELP_BTN",
	"Pomoc kontekstowa (F1)...": "CONTEXT_HELP_BTN",
	"Pomoc (F1)...": "BTN_HELP_CONTEXT",
	"Skonfiguruj skróty klawiszowe...": "GESTURES_BTN",
	"Przywróć ustawienia domyślne...": "RESTORE_DEFAULTS_BTN",
	"Filtrowanie mowy w Outlooku: włączone (komunikaty wyciszone)": "TOGGLE_FILTER_ON",
	"Filtrowanie mowy w Outlooku: wyłączone (komunikaty mowy włączone)": "TOGGLE_FILTER_OFF",
	"Dźwięki statusów w Outlooku: włączone": "TOGGLE_SOUNDS_ON",
	"Dźwięki statusów w Outlooku: wyłączone": "TOGGLE_SOUNDS_OFF",
	"Przełącza filtrowanie mowy o wiadomościach w skrzynce programu Outlook": "GESTURE_DESC_TOGGLE_FILTER",
	"Przełącza odtwarzanie dźwięków statusów w skrzynce programu Outlook": "GESTURE_DESC_TOGGLE_SOUNDS",
	"Odczytuje na żądanie wszystkie statusy aktualnej wiadomości w Outlooku": "GESTURE_DESC_REPORT_STATUSES",
	"Wyświetla pomoc kontekstową i spis skrótów klawiszowych dodatku Outlook Speech Filter": "GESTURE_DESC_CONTEXT_HELP",
}

def _t(key, **kwargs):
	lang = get_effective_language()
	realKey = REVERSE_MAP.get(key, key)
	val = None
	if realKey in TRANSLATIONS:
		val = TRANSLATIONS[realKey].get(lang) or TRANSLATIONS[realKey].get("en")
	if val is None:
		val = str(key)
	if kwargs:
		try:
			val = val.format(**kwargs)
		except Exception:
			pass
	return val

def _(text):
	return _t(text)


class AccessibleMessageDialog(wx.Dialog):
	"""
	W pełni dostępne okno komunikatów i instrukcji dla czytnika ekranu NVDA.
	Fokus jest automatycznie ustawiany w wielowierszowym polu tekstowym tylko do odczytu (wx.TextCtrl).
	Dzięki temu użytkownik programu NVDA od razu słyszy treść komunikatu,
	może swobodnie czytać go linijka po linijce strzałkami w górę i w dół,
	skopiować tekst (Ctrl+C), a także zamknąć okno klawiszem Enter lub Escape.
	"""

	def __init__(self, parent, message: str, title: str = "Informacja", buttonLabel: str = "&Zamknij"):
		super().__init__(
			parent or gui.mainFrame,
			title=title,
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self.SetMinSize((560, 320))
		mainSizer = wx.BoxSizer(wx.VERTICAL)

		self.textCtrl = wx.TextCtrl(
			self,
			value=message,
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2 | wx.BORDER_THEME,
			size=(640, 340),
		)
		self.textCtrl.SetName(title)
		mainSizer.Add(self.textCtrl, 1, wx.ALL | wx.EXPAND, 12)

		btnSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.btnClose = wx.Button(self, wx.ID_OK, label=buttonLabel)
		self.btnClose.SetDefault()
		btnSizer.Add(self.btnClose, 0, wx.ALL, 0)
		mainSizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.RIGHT | wx.BOTTOM, 12)

		self.SetSizerAndFit(mainSizer)
		self.CenterOnScreen()

		self.SetAffirmativeId(wx.ID_OK)
		self.SetEscapeId(wx.ID_OK)

		def _focusText():
			try:
				self.textCtrl.SetFocus()
				self.textCtrl.SetInsertionPoint(0)
			except Exception:
				pass

		wx.CallAfter(_focusText)


class AccessibleConfirmDialog(wx.Dialog):
	"""
	W pełni dostępne okno potwierdzenia (Tak / Nie) dla czytnika ekranu NVDA.
	Fokus jest automatycznie ustawiany w polu tekstowym z pytaniem/ostrzeżeniem,
	co pozwala użytkownikowi NVDA od razu zapoznać się z treścią, a następnie
	przejść klawiszem Tab do przycisków 'Tak' lub 'Nie' (lub użyć Alt+T / Alt+N bądź Esc).
	"""

	def __init__(self, parent, message: str, title: str = "Potwierdzenie", defaultYes: bool = False):
		super().__init__(
			parent or gui.mainFrame,
			title=title,
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self.SetMinSize((540, 280))
		mainSizer = wx.BoxSizer(wx.VERTICAL)

		self.textCtrl = wx.TextCtrl(
			self,
			value=message,
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2 | wx.BORDER_THEME,
			size=(620, 220),
		)
		self.textCtrl.SetName(title)
		mainSizer.Add(self.textCtrl, 1, wx.ALL | wx.EXPAND, 12)

		btnSizer = wx.BoxSizer(wx.HORIZONTAL)
		isPl = (get_effective_language() == "pl")
		yesLabel = "&Tak" if isPl else "&Yes"
		noLabel = "&Nie" if isPl else "&No"

		self.btnYes = wx.Button(self, wx.ID_YES, label=yesLabel)
		self.btnNo = wx.Button(self, wx.ID_NO, label=noLabel)

		if defaultYes:
			self.btnYes.SetDefault()
		else:
			self.btnNo.SetDefault()

		btnSizer.Add(self.btnYes, 0, wx.RIGHT, 8)
		btnSizer.Add(self.btnNo, 0, 0)
		mainSizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.RIGHT | wx.BOTTOM, 12)

		self.SetSizerAndFit(mainSizer)
		self.CenterOnScreen()

		self.SetAffirmativeId(wx.ID_YES)
		self.SetEscapeId(wx.ID_NO)

		def _focusText():
			try:
				self.textCtrl.SetFocus()
				self.textCtrl.SetInsertionPoint(0)
			except Exception:
				pass

		wx.CallAfter(_focusText)


def showAccessibleMessageBox(message: str, title: str = "Informacja", parent=None, buttonLabel: str = "&Zamknij"):
	"""Wyświetla w pełni dostępne okno komunikatu z fokusem na tekście."""
	if not wx.IsMainThread():
		wx.CallAfter(showAccessibleMessageBox, message, title, parent, buttonLabel)
		return
	gui.mainFrame.prePopup()
	try:
		dlg = AccessibleMessageDialog(parent, message, title=title, buttonLabel=buttonLabel)
		dlg.ShowModal()
		dlg.Destroy()
	finally:
		gui.mainFrame.postPopup()


def showAccessibleConfirmBox(message: str, title: str = "Potwierdzenie", parent=None, defaultYes: bool = False) -> bool:
	"""Wyświetla w pełni dostępne okno pytania Tak/Nie z fokusem na tekście pytania."""
	gui.mainFrame.prePopup()
	try:
		dlg = AccessibleConfirmDialog(parent, message, title=title, defaultYes=defaultYes)
		res = dlg.ShowModal()
		dlg.Destroy()
		return (res == wx.ID_YES)
	finally:
		gui.mainFrame.postPopup()


def generate_diagnostic_report():
	"""
	Generuje szczegółowy raport diagnostyczny środowiska (Windows, NVDA, Outlook, ustawienia dodatku).
	Raport nie gromadzi treści wiadomości e-mail ani haseł.
	"""
	lines = []
	lines.append("=" * 60)
	lines.append("OUTLOOK SPEECH FILTER - RAPORT DIAGNOSTYCZNY / DIAGNOSTIC REPORT")
	lines.append("=" * 60)
	lines.append(f"Data wygenerowania / Timestamp: {time.strftime('%Y-%m-%d %H:%M:%S')}")
	lines.append(f"Dodatek / Add-on: {ADDON_NAME} v{ADDON_VERSION} (Build {ADDON_BUILD})")
	lines.append(f"Skonfigurowany język / Configured language: {config.conf.get(CONFIG_SECTION, {}).get('language', 'auto')} (aktywny: {get_effective_language()})")
	lines.append("")
	lines.append("--- ŚRODOWISKO SYSTEMOWE / SYSTEM ENVIRONMENT ---")
	try:
		lines.append(f"System operacyjny / OS: {platform.platform()} ({platform.architecture()[0]})")
		lines.append(f"Wersja Windows / Windows version: {sys.getwindowsversion()}")
		lines.append(f"Python: {sys.version.split()[0]} ({sys.executable})")
	except Exception as e:
		lines.append(f"Błąd odczytu systemu: {e}")
	try:
		import versionInfo
		lines.append(f"NVDA Version: {versionInfo.version} ({versionInfo.version_year}.{versionInfo.version_major})")
	except Exception as e:
		lines.append(f"NVDA Version: błąd odczytu ({e})")
	try:
		lines.append(f"NVDA Config Path: {globalVars.appArgs.configPath}")
	except Exception:
		pass
	lines.append("")
	lines.append("--- STAN PROGRAMU MICROSOFT OUTLOOK ---")
	try:
		import psutil
		procs = [p for p in psutil.process_iter(['name']) if p.info.get('name') and 'outlook' in p.info['name'].lower()]
		lines.append(f"Proces Outlook uruchomiony / Outlook process running: {'TAK / YES' if procs else 'NIE / NO'} (liczba procesów: {len(procs)})")
	except Exception:
		lines.append("psutil niedostępny do weryfikacji procesów")
	try:
		focus = api.getFocusObject()
		appMod = getattr(focus, "appModule", None)
		lines.append(f"Aktualny focus appModule: {getattr(appMod, 'appName', None)}")
		if appMod and getattr(appMod, "appName", "").lower() == "outlook":
			nativeOm = getattr(appMod, "nativeOm", None)
			if nativeOm:
				try:
					lines.append(f"Outlook COM nativeOm Version: {nativeOm.Version}")
				except Exception as e:
					lines.append(f"Outlook COM nativeOm Version error: {e}")
	except Exception as e:
		lines.append(f"Błąd sprawdzania focus/appModule: {e}")
	lines.append("")
	lines.append("--- KONFIGURACJA DODATKU / ADD-ON CONFIGURATION ---")
	if CONFIG_SECTION in config.conf:
		for k, v in sorted(config.conf[CONFIG_SECTION].items()):
			lines.append(f"  {k} = {v}")
	else:
		lines.append("  Brak sekcji konfiguracyjnej w nvda.ini (używane wartości domyślne)")
	lines.append("")
	lines.append("--- PLIKI DŹWIĘKOWE / SOUND FILES ---")
	for sound_name in CANONICAL_STATUS_ORDER:
		wav_path = os.path.join(SOUNDS_DIR, f"{sound_name}.wav")
		lines.append(f"  {sound_name}.wav: {'ISTNIEJE / EXISTS (' + str(os.path.getsize(wav_path)) + ' B)' if os.path.isfile(wav_path) else 'BRAK / MISSING'}")
	lines.append("")
	lines.append("--- WYCIĄG Z LOGU NVDA / NVDA LOG EXCERPT ---")
	try:
		logFile = getattr(getattr(globalVars, "appArgs", None), "logFileName", None)
		if logFile and os.path.isfile(logFile):
			with open(logFile, "r", encoding="utf-8", errors="replace") as f:
				all_lines = f.readlines()
			filtered = [l.strip() for l in all_lines if any(k in l.lower() for k in ("outlook", "speechfilter", "messagecleaner", "error", "exception", "traceback"))]
			for fl in filtered[-50:]:
				lines.append(f"  {fl}")
		else:
			lines.append("  Brak dostępu do pliku logu NVDA")
	except Exception as e:
		lines.append(f"  Błąd odczytu logu NVDA: {e}")
	lines.append("=" * 60)
	lines.append("KONIEC RAPORTU / END OF REPORT")
	lines.append("=" * 60)
	return "\n".join(lines)


def copy_text_to_clipboard(text):
	"""Kopiuje podany tekst do schowka systemowego."""
	try:
		import api
		if hasattr(api, "copyToClip"):
			api.copyToClip(text)
			return True
	except Exception:
		pass
	try:
		if wx.TheClipboard.Open():
			wx.TheClipboard.SetData(wx.TextDataObject(text))
			wx.TheClipboard.Close()
			return True
	except Exception:
		pass
	return False


class ContactAuthorDialog(wx.Dialog):
	"""
	Dostępne okno dialogowe kontaktu z autorem i zgłaszania uwag.
	Przekierowuje do w 100% anonimowego formularza online (Google Forms),
	bez gromadzenia adresów e-mail ani żadnych danych osobowych (RODO).
	"""
	def __init__(self, parent, topicKey="", topicTitle=""):
		self.topicKey = topicKey
		self.topicTitle = topicTitle
		dlgTitle = _t("CONTACT_DLG_TITLE")
		super().__init__(parent, title=dlgTitle, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, size=(540, 370))
		self.reportContent = None

		mainSizer = wx.BoxSizer(wx.VERTICAL)

		# 1. Informacja o anonimowości i RODO
		infoText = wx.StaticText(self, label=_t("CONTACT_DLG_INFO"))
		infoText.Wrap(500)
		mainSizer.Add(infoText, flag=wx.ALL, border=10)

		# 2. Główny przycisk: Otwórz formularz w przeglądarce
		openBtnBox = wx.BoxSizer(wx.VERTICAL)
		self.openFormBtn = wx.Button(self, label=_t("CONTACT_BTN_OPEN_FORM"))
		self.openFormBtn.Bind(wx.EVT_BUTTON, self.onOpenForm)
		self.openFormBtn.SetDefault()
		openBtnBox.Add(self.openFormBtn, flag=wx.EXPAND | wx.BOTTOM, border=10)
		mainSizer.Add(openBtnBox, flag=wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, border=10)

		# 3. Sekcja raportu diagnostycznego
		reportSectionLabel = wx.StaticText(self, label=_t("CONTACT_REPORT_SECTION_LABEL"))
		mainSizer.Add(reportSectionLabel, flag=wx.LEFT | wx.RIGHT | wx.BOTTOM, border=10)

		reportBtnBox = wx.BoxSizer(wx.HORIZONTAL)
		self.copyReportBtn = wx.Button(self, label=_t("CONTACT_BTN_COPY_REPORT"))
		self.copyReportBtn.Bind(wx.EVT_BUTTON, self.onCopyReport)
		reportBtnBox.Add(self.copyReportBtn, flag=wx.RIGHT, border=10)

		self.viewReportBtn = wx.Button(self, label=_t("CONTACT_BTN_VIEW_REPORT"))
		self.viewReportBtn.Bind(wx.EVT_BUTTON, self.onViewReport)
		reportBtnBox.Add(self.viewReportBtn, flag=wx.RIGHT, border=10)

		self.saveReportBtn = wx.Button(self, label=_t("CONTACT_BTN_SAVE_REPORT"))
		self.saveReportBtn.Bind(wx.EVT_BUTTON, self.onSaveReport)
		reportBtnBox.Add(self.saveReportBtn)

		mainSizer.Add(reportBtnBox, flag=wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, border=10)

		# 4. Przyciski: Pomoc (F1) i Zamknij
		closeSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.helpBtn = wx.Button(self, label=_t("BTN_HELP_CONTEXT"))
		self.helpBtn.Bind(wx.EVT_BUTTON, lambda evt: showContextHelp("contact", parent=self))
		closeSizer.Add(self.helpBtn, flag=wx.RIGHT, border=10)

		self.closeBtn = wx.Button(self, wx.ID_CANCEL, label=_t("BTN_CLOSE"))
		closeSizer.Add(self.closeBtn)
		mainSizer.Add(closeSizer, flag=wx.ALIGN_RIGHT | wx.ALL, border=10)

		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)

		self.SetSizer(mainSizer)
		self.Layout()
		self.CentreOnParent()

	def onCharHook(self, event):
		if event.GetKeyCode() == wx.WXK_F1:
			showContextHelp("contact", parent=self)
			return
		event.Skip()

	def getReport(self):
		if not self.reportContent:
			self.reportContent = generate_diagnostic_report()
		return self.reportContent

	def onOpenForm(self, evt):
		customUrl = config.conf.get(CONFIG_SECTION, {}).get("feedbackFormUrl", "").strip()
		if customUrl:
			url = customUrl
		else:
			lang = get_effective_language()
			if lang == "en" and DEFAULT_FEEDBACK_FORM_URL_EN:
				url = DEFAULT_FEEDBACK_FORM_URL_EN
			else:
				url = DEFAULT_FEEDBACK_FORM_URL_PL
		try:
			os.startfile(url)
		except Exception:
			try:
				import webbrowser
				webbrowser.open(url)
			except Exception as e:
				log.warning(f"outlookSpeechFilter: nie można otworzyć formularza: {e}")

	def onCopyReport(self, evt):
		rep = self.getReport()
		ok = copy_text_to_clipboard(rep)
		if ok:
			ui.message(_t("CONTACT_REPORT_COPIED_MSG"))
		else:
			ui.message("Nie udało się skopiować raportu do schowka.")

	def onViewReport(self, evt):
		rep = self.getReport()
		dlg = wx.Dialog(self, title=_t("REPORT_VIEW_TITLE"), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, size=(600, 480))
		s = wx.BoxSizer(wx.VERTICAL)
		tc = wx.TextCtrl(dlg, value=rep, style=wx.TE_MULTILINE | wx.TE_READONLY)
		tc.SetName(_t("REPORT_VIEW_TITLE"))
		s.Add(tc, proportion=1, flag=wx.EXPAND | wx.ALL, border=10)
		closeBtn = wx.Button(dlg, wx.ID_OK, label=_t("BTN_CLOSE"))
		s.Add(closeBtn, flag=wx.ALIGN_RIGHT | wx.RIGHT | wx.BOTTOM, border=10)
		dlg.SetSizer(s)
		dlg.CentreOnParent()
		dlg.ShowModal()
		dlg.Destroy()

	def onSaveReport(self, evt):
		rep = self.getReport()
		desktop = os.path.join(os.path.expanduser("~"), "Desktop")
		fn = f"Raport_OutlookSpeechFilter_{time.strftime('%Y%m%d_%H%M%S')}.txt"
		path = os.path.join(desktop, fn)
		try:
			with open(path, "w", encoding="utf-8") as f:
				f.write(rep)
			showAccessibleMessageBox(
				_t("CONTACT_REPORT_SAVED_MSG", path=path),
				_t("CONTACT_DESKTOP_SAVED_TITLE"),
				parent=self
			)
		except Exception as e:
			showAccessibleMessageBox(str(e), _t("CONTACT_DLG_TITLE"), parent=self)


# Sound playback debounce state
_last_sound_time = 0
_last_sound_key = None
_pending_debounce_call = None
_pending_sound_calls = []
_activeSettingsPanelRef = None
_msg_attachment_cache = {}

def get_active_settings_panel():
	global _activeSettingsPanelRef
	if _activeSettingsPanelRef:
		panel = _activeSettingsPanelRef()
		if panel:
			try:
				if isinstance(panel, wx.Window) and panel:
					return panel
			except Exception:
				pass
	return None


def _get_scaled_sound_path(sound_name, volume=None):
	"""
	Returns the path to a .wav sound file scaled to the specified volume (0-100).
	If volume is 100, returns original file path.
	If volume <= 0, returns None (muted).
	Caches scaled audio files in temp directory for zero-latency subsequent playback.
	"""
	if volume is None:
		panel = get_active_settings_panel()
		if panel and hasattr(panel, "soundVolumeSlider"):
			try:
				volume = panel.soundVolumeSlider.GetValue()
			except Exception:
				volume = None
		if volume is None:
			try:
				conf = config.conf.get(CONFIG_SECTION, {})
				volume = conf.get("soundVolume", 80)
			except Exception:
				volume = 80

	try:
		volume = int(volume)
	except (ValueError, TypeError):
		volume = 100

	if volume <= 0:
		return None

	theme = "associative"
	panel = get_active_settings_panel()
	if panel and hasattr(panel, "soundThemeChoice") and hasattr(panel, "themeKeys"):
		try:
			selIdx = panel.soundThemeChoice.GetSelection()
			if 0 <= selIdx < len(panel.themeKeys):
				theme = panel.themeKeys[selIdx]
		except Exception:
			pass
	if not theme or theme == "associative":
		try:
			theme = config.conf.get(CONFIG_SECTION, {}).get("soundTheme", "associative")
		except Exception:
			theme = "associative"

	theme_dir = os.path.join(SOUNDS_DIR, theme)
	candidate = os.path.join(theme_dir, f"{sound_name}.wav")
	orig_path = candidate if os.path.isfile(candidate) else os.path.join(SOUNDS_DIR, f"{sound_name}.wav")
	if not os.path.isfile(orig_path):
		return None

	if volume >= 100:
		return orig_path

	cached_path = os.path.join(SOUNDS_CACHE_DIR, f"{theme}_{sound_name}_vol{volume}.wav")
	if os.path.isfile(cached_path) and os.path.getsize(cached_path) > 44:
		return cached_path

	try:
		os.makedirs(SOUNDS_CACHE_DIR, exist_ok=True)
		factor = volume / 100.0
		with wave.open(orig_path, "rb") as in_wf:
			nchannels = in_wf.getnchannels()
			sampwidth = in_wf.getsampwidth()
			framerate = in_wf.getframerate()
			nframes = in_wf.getnframes()
			raw_frames = in_wf.readframes(nframes)

		if sampwidth == 2:
			# 16-bit signed PCM
			samples = array.array("h", raw_frames)
			scaled = array.array("h", (max(-32768, min(32767, int(s * factor))) for s in samples))
			scaled_bytes = scaled.tobytes()
		elif sampwidth == 1:
			# 8-bit unsigned PCM
			samples = array.array("B", raw_frames)
			scaled = array.array("B", (max(0, min(255, int(128 + (b - 128) * factor))) for b in samples))
			scaled_bytes = scaled.tobytes()
		else:
			return orig_path

		tmp_path = cached_path + f".tmp_{os.getpid()}_{int(time.time() * 1000)}"
		with wave.open(tmp_path, "wb") as out_wf:
			out_wf.setnchannels(nchannels)
			out_wf.setsampwidth(sampwidth)
			out_wf.setframerate(framerate)
			out_wf.writeframes(scaled_bytes)

		os.replace(tmp_path, cached_path)
		return cached_path
	except Exception as e:
		log.debugWarning(f"Nie udało się przeskalować dźwięku {sound_name} do głośności {volume}: {e}")
		return orig_path


def _play_status_sound(sound_name, volume=None):
	"""Play a single status earcon sound file at the specified or configured volume."""
	filepath = _get_scaled_sound_path(sound_name, volume)
	if filepath and os.path.isfile(filepath):
		try:
			nvwave.playWaveFile(filepath, asynchronous=True)
		except Exception:
			pass


def _stop_pending_sounds():
	"""Stop all pending scheduled sound timers and immediately purge playing audio."""
	global _pending_sound_calls, _pending_debounce_call
	if _pending_debounce_call:
		try:
			_pending_debounce_call.Stop()
		except Exception:
			pass
		_pending_debounce_call = None
	for call in _pending_sound_calls:
		try:
			call.Stop()
		except Exception:
			pass
	_pending_sound_calls.clear()
	try:
		winsound.PlaySound(None, winsound.SND_PURGE)
	except Exception:
		pass


def _schedule_status_sounds_sequence(sound_names, volume=None, delay_ms=230):
	"""Play a sequence of sounds with spacing in ms, cancelling previous calls (for settings dialog tests)."""
	global _pending_sound_calls
	_stop_pending_sounds()
	if not sound_names:
		return
	_play_status_sound(sound_names[0], volume=volume)
	for i, name in enumerate(sound_names[1:], start=1):
		call = wx.CallLater(i * delay_ms, lambda s=name, v=volume: _play_status_sound(s, volume=v))
		_pending_sound_calls.append(call)


def _schedule_status_sounds(sound_names, volume=None, debounce_ms=120):
	"""
	Schedule status sounds for playback with debounce and spacing.
	Opóźnia odtwarzanie o debounce_ms. Jeśli użytkownik szybko nawiguje po liście wiadomości,
	poprzednie timery są anulowane i dźwięki wiadomości pośrednich NIE SĄ w ogóle odtwarzane.
	Zapobiega to kolejkowaniu dźwięków i zacinaniu nawigacji na wolniejszych komputerach.
	"""
	global _pending_debounce_call, _pending_sound_calls
	_stop_pending_sounds()
	if not sound_names:
		return

	def start_playback():
		global _pending_debounce_call, _pending_sound_calls
		_pending_debounce_call = None
		# Odtwórz pierwszy dźwięk od razu po upływie debounce
		_play_status_sound(sound_names[0], volume=volume)
		# Kolejne dźwięki (jeśli wiadomość ma kilka statusów) odtwarzaj w odstępach 170 ms
		for i, name in enumerate(sound_names[1:], start=1):
			call = wx.CallLater(i * 170, lambda s=name, v=volume: _play_status_sound(s, volume=v))
			_pending_sound_calls.append(call)

	_pending_debounce_call = wx.CallLater(debounce_ms, start_playback)


_original_UIAGridRow_get_name = None
_original_getReceivedMessageString = None
_patches_applied = False


ATTACHMENT_TOKENS = {
	"ma załączniki", "ma załącznik", "załącznik", "załączniki",
	"has attachment", "has attachments", "attachment", "attachments",
}
UNREAD_TOKENS = {"nieprzeczytana", "nieprzeczytane", "unread"}
READ_TOKENS = {"przeczytana", "przeczytane", "read"}
FORWARD_REPLY_TOKENS = {
	"odpowiedziano", "odpowiedziano wszystkim", "przesłano dalej", "przesłana dalej",
	"replied", "replied all", "forwarded",
}
EXTERNAL_TOKENS = {"zewnętrzna", "zewnętrzny", "zewnętrzne", "external", "[zewnętrzna]", "[zewnętrzny]", "[external]"}
EXTERNAL_PREFIXES = ("[zewnętrzna]", "[zewnętrzny]", "[external]", "zewnętrzna:", "zewnętrzny:", "external:")

RE_ATTACHMENT_TEXT = re.compile(
	r'\b(?:załącznik|załączniki|załącznika|załączników|attachment|attachments)\b',
	re.IGNORECASE,
)

RE_EXACT_ATTACHMENT_INDICATOR = re.compile(
	r'^(?:(?:\d+\s+)?(?:załącznik(?:i|a|ów)?|attachments?)|(?:has\s+attachments?|ma\s+załączniki?))(?:\s*:\s*\d+)?$',
	re.IGNORECASE
)

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".ico", ".tiff", ".tif", ".svg"}


def _is_attachment_inline_image(att):
	"""
	Precyzyjnie rozpoznaje, czy dany załącznik jest obrazem osadzonym w treści (inline).
	Chroni zwykłe pliki o nazwach np. 'image_budget.pdf' oraz osadzone maile .msg (Type == 5).
	"""
	try:
		fn = (getattr(att, "FileName", "") or "").lower()
		ext = os.path.splitext(fn)[1]
		# 1. Sprawdź właściwości MAPI (Content-ID oraz flagę MHTML_REF)
		cid = ""
		flags = 0
		try:
			pa = getattr(att, "PropertyAccessor", None)
			if pa:
				cid = str(pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F") or "")
				flags = int(pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37140003") or 0)
		except Exception:
			pass

		if (cid or (flags & 4)) and (ext in IMAGE_EXTENSIONS or not ext):
			return True

		attType = int(getattr(att, "Type", 1) or 1)
		# Typ 6 to olOLE (osadzony obiekt OLE)
		if attType == 6 and (fn.startswith("oledata") or not ext):
			return True

		# Zrzuty ekranu wklejane ze schowka automatycznie nazywane przez Outlook (np. image001.png)
		if ext in IMAGE_EXTENSIONS and re.match(r'^image\d{3,}\.(?:png|jpg|jpeg|gif|bmp)$', fn) and (cid or not fn):
			return True

		return False
	except Exception:
		return False

_last_focused_statuses = []


def _is_attachment_text(text):
	if not text:
		return False
	t = text.strip().lower()
	if t in ATTACHMENT_TOKENS or t.rstrip(":").strip() in ATTACHMENT_TOKENS:
		return True
	return bool(RE_EXACT_ATTACHMENT_INDICATOR.match(t))


def _extract_attachment_count(text):
	if not text:
		return 0
	m = re.search(r'\d+', text)
	if m:
		try:
			return int(m.group(0))
		except Exception:
			pass
	return 1 if _is_attachment_text(text) else 0


def _format_attachment_status(count):
	"""Returns human-readable text for attachments, always including the count."""
	if count <= 0:
		return ""
	if count == 1:
		return _t("STATUS_ATTACHMENT_SINGLE")
	return _t("STATUS_ATTACHMENTS_MULTIPLE", count=count)


def _get_real_attachments_count(selection, uiaObj=None):
	"""Returns the count of real (non-inline, non-hidden) file attachments."""
	if not selection:
		if uiaObj:
			try:
				uiaElem = getattr(uiaObj, "UIAElement", None)
				if uiaElem:
					cachedChildren = uiaElem.getCachedChildren()
					if cachedChildren:
						for idx in range(cachedChildren.length):
							el = cachedChildren.getElement(idx)
							n = (el.cachedName or "").strip().lower()
							if _is_attachment_text(n):
								cnt = _extract_attachment_count(n)
								return cnt if cnt > 0 else 1
			except Exception:
				pass
		return 0

	try:
		attColl = getattr(selection, "attachments", None)
		total = int(getattr(attColl, "count", 0)) if attColl else 0
		if total == 0:
			return 0

		realCount = 0
		for i in range(1, total + 1):
			try:
				att = attColl.item(i)
				if _is_attachment_inline_image(att):
					continue
				realCount += 1
			except Exception:
				realCount += 1

		return realCount
	except Exception:
		return 0


def _format_inline_images_status(count):
	"""Returns human-readable text for inline graphics/screenshots, including count if > 1."""
	if count <= 0:
		return ""
	if count == 1:
		return _t("STATUS_INLINE_IMAGE")
	return _t("STATUS_INLINE_IMAGES_MULTIPLE", count=count)


def _get_inline_images_count(selection):
	"""Returns the count of inline/embedded images and screenshots in the message body."""
	if not selection:
		return 0
	try:
		attColl = getattr(selection, "attachments", None)
		total = int(getattr(attColl, "count", 0)) if attColl else 0
		if total == 0:
			return 0

		inlineCount = 0
		for i in range(1, total + 1):
			try:
				att = attColl.item(i)
				if _is_attachment_inline_image(att):
					inlineCount += 1
			except Exception:
				pass

		return inlineCount
	except Exception:
		return 0


_in_external_check = False


def _is_external_message(selection, uiaObj=None):
	"""Accurately checks whether the given message is external without invoking NVDAObject descriptors."""
	global _in_external_check
	if _in_external_check:
		return False
	_in_external_check = True
	try:
		# 1. Sprawdź temat w selection
		if selection:
			try:
				subj = (selection.subject or "").lower()
				if any(t in subj for t in EXTERNAL_PREFIXES):
					return True
			except Exception:
				pass

		# 2. Sprawdź UIAElement bezpośrednio (bez odwoływania się do właściwości .name na NVDAObject)
		if uiaObj:
			targetElems = []
			elem = getattr(uiaObj, "UIAElement", None)
			if elem:
				targetElems.append(elem)
			parentObj = getattr(uiaObj, "parent", None)
			if parentObj:
				pElem = getattr(parentObj, "UIAElement", None)
				if pElem:
					targetElems.append(pElem)

			for el in targetElems:
				try:
					rawName = (getattr(el, "cachedName", "") or getattr(el, "currentName", "") or "").lower()
					if any(t in rawName for t in EXTERNAL_PREFIXES):
						return True
				except Exception:
					pass

				try:
					cachedChildren = el.getCachedChildren()
					if cachedChildren:
						for idx in range(cachedChildren.length):
							childEl = cachedChildren.getElement(idx)
							n = (childEl.cachedName or "").strip().lower()
							if n in EXTERNAL_TOKENS or n.rstrip(":").strip() in EXTERNAL_TOKENS:
								return True
							if any(n.startswith(t) for t in EXTERNAL_PREFIXES):
								return True
				except Exception:
					pass

		return False
	finally:
		_in_external_check = False


def _get_message_statuses_report(appModule, selection, windowThreadID=None, uiaObj=None):
	"""Returns a list of all human-readable status descriptions for a message in CANONICAL_STATUS_ORDER."""
	# 1. Unread
	isUnread = False
	if selection:
		try:
			if selection.unread:
				isUnread = True
		except (COMError, Exception):
			pass
	elif uiaObj:
		try:
			uiaElem = getattr(uiaObj, "UIAElement", None)
			if uiaElem:
				cachedChildren = uiaElem.getCachedChildren()
				if cachedChildren:
					for idx in range(cachedChildren.length):
						el = cachedChildren.getElement(idx)
						n = (el.cachedName or "").strip().lower()
						if n in UNREAD_TOKENS:
							isUnread = True
							break
		except Exception:
			pass

	# 2. Flag
	hasFlag = False
	if selection:
		try:
			if selection.flagIcon or selection.flagStatus:
				hasFlag = True
		except (COMError, Exception):
			pass
	if not hasFlag and uiaObj:
		try:
			for child in getattr(uiaObj, "children", []) or []:
				if getattr(child, "UIAClassName", "") == "FlagField":
					hasFlag = True
					break
		except Exception:
			pass

	# 3. Attachment
	realCount = _get_real_attachments_count(selection, uiaObj)
	inlineImgCount = _get_inline_images_count(selection)

	# 4. Reply / Forward (wykrywane bezpiecznie przez elementy UIA bez odwoływania się do mapiObject)
	verbText = None
	verbSound = None
	if uiaObj:
		try:
			for child in getattr(uiaObj, "children", []) or []:
				cn = (getattr(child, "name", "") or "").strip().lower()
				if cn in FORWARD_REPLY_TOKENS:
					verbText = getattr(child, "name", "")
					verbSound = "forwarded" if any(w in cn for w in ("przesł", "forward")) else "replied"
					break
		except Exception:
			pass

	# 5. Importance
	importanceText = None
	importanceSound = None
	if selection:
		try:
			importance = selection.importance
			if importance == 2:
				importanceText = _t("STATUS_IMPORTANCE_HIGH")
				importanceSound = "importance_high"
			elif importance == 0:
				importanceText = _t("STATUS_IMPORTANCE_LOW")
				importanceSound = "importance_low"
		except (COMError, Exception):
			pass

	# 6. Meeting request
	isMeeting = False
	if selection:
		try:
			if selection.messageClass == "IPM.Schedule.Meeting.Request":
				isMeeting = True
		except (COMError, Exception):
			pass

	# 7. Categories
	catText = None
	if selection:
		try:
			cat = (selection.categories or "").strip()
			if cat and cat.lower() not in ("bez kategorii", "brak kategorii", "no category", "no categories", "none"):
				catText = _t("STATUS_CATEGORY", cat=cat)
		except (COMError, Exception):
			pass

	# 8. External
	isExternal = _is_external_message(selection, uiaObj)

	# Build result strictly adhering to CANONICAL_STATUS_ORDER
	statuses = []
	for status_key in CANONICAL_STATUS_ORDER:
		if status_key == "unread" and isUnread:
			statuses.append(_t("STATUS_UNREAD"))
		elif status_key == "flag" and hasFlag:
			statuses.append(_t("STATUS_FLAG"))
		elif status_key == "attachment" and realCount > 0:
			statuses.append(_format_attachment_status(realCount))
		elif status_key == "inline_image" and inlineImgCount > 0:
			statuses.append(_format_inline_images_status(inlineImgCount))
		elif status_key == "replied" and verbText and verbSound == "replied":
			statuses.append(verbText)
		elif status_key == "forwarded" and verbText and verbSound == "forwarded":
			statuses.append(verbText)
		elif status_key == "importance_high" and importanceText and importanceSound == "importance_high":
			statuses.append(importanceText)
		elif status_key == "importance_low" and importanceText and importanceSound == "importance_low":
			statuses.append(importanceText)
		elif status_key == "meeting" and isMeeting:
			statuses.append(_t("STATUS_MEETING"))
		elif status_key == "category" and catText:
			statuses.append(catText)
		elif status_key == "external" and isExternal:
			statuses.append(_t("STATUS_EXTERNAL"))

	return statuses


KNOWN_COLORS = (
	"pomarańczowa", "pomarańczowy", "pomarańczowe",
	"czerwona", "czerwony", "czerwone",
	"żółta", "żółty", "żółte",
	"zielona", "zielony", "zielone",
	"niebieska", "niebieski", "niebieskie",
	"błękitna", "błękitny", "błękitne",
	"fioletowa", "fioletowy", "fioletowe",
	"purpurowa", "purpurowy", "purpurowe",
	"szara", "szary", "szare",
	"stalowa", "stalowy", "stalowe",
	"oliwkowa", "oliwkowy", "oliwkowe",
	"morska", "morski", "morskie",
	"brązowa", "brązowy", "brązowe",
	"biała", "biały", "białe",
	"czarna", "czarny", "czarne",
	"orange", "red", "yellow", "green", "blue", "purple", "gray", "grey", "brown"
)


def _clean_group_header_name(rawName):
	"""Cleans verbose Outlook group headers so that the category color is spoken first, then label, then count."""
	if not rawName:
		return rawName

	lower = rawName.lower()
	if not any(k in lower for k in ("grupuj według", "group by", "kategorie:", "kategoria:", "flaga:", "status flagi:", "liczba elementów", "items:")):
		return rawName

	parts = re.split(r'\s+grupa\s+|\s+group\s+', rawName, flags=re.IGNORECASE)
	mainPart = parts[0].strip() if parts else rawName.strip()

	countMatch = re.search(r'\((?:Liczba elementów|Items?|Elementy):\s*(\d+)\)', mainPart, re.IGNORECASE)
	countText = ""
	if countMatch:
		countText = f"{countMatch.group(1)} elementów"

	clean = re.sub(r'\((?:Liczba elementów|Items?|Elementy):\s*\d+\)', '', mainPart, flags=re.IGNORECASE)
	clean = re.sub(r'\b(?:grupuj według|group by)\s*:\s*', '', clean, flags=re.IGNORECASE)
	clean = re.sub(r'\b(?:zwinięte|zwinięty|rozwinięte|rozwinięty|collapsed|expanded)\s*:\s*', '', clean, flags=re.IGNORECASE)
	clean = re.sub(r'\b(?:kategorie|kategoria|categories|category|flagi|flaga|status flagi|flag status)\s*:\s*', '', clean, flags=re.IGNORECASE)
	clean = clean.strip(" ,:-")

	foundColor = ""
	remainingLabel = clean
	for color in KNOWN_COLORS:
		match = re.search(rf'\b({re.escape(color)})\b', clean, flags=re.IGNORECASE)
		if match:
			foundColor = match.group(1)
			remainingLabel = re.sub(rf'\b{re.escape(color)}\b', '', clean, flags=re.IGNORECASE).strip(" ,:-")
			break

	outParts = []
	if foundColor:
		outParts.append(foundColor)
	if remainingLabel:
		outParts.append(remainingLabel)
	elif not foundColor and clean:
		outParts.append(clean)
	if countText:
		outParts.append(countText)

	return ", ".join(outParts)


_in_custom_get_name = False


def _custom_UIAGridRow_get_name(self):
	global _in_custom_get_name
	if _in_custom_get_name:
		if _original_UIAGridRow_get_name:
			try:
				return _original_UIAGridRow_get_name(self)
			except Exception:
				pass
		return getattr(getattr(self, "UIAElement", None), "cachedName", "") or ""

	_in_custom_get_name = True
	try:
		return _custom_UIAGridRow_get_name_impl(self)
	finally:
		_in_custom_get_name = False


def _custom_UIAGridRow_get_name_impl(self):
	conf = config.conf.get(CONFIG_SECTION, {})
	enabled = _to_bool(conf.get("enabled", True), default=True)
	useSounds = _to_bool(conf.get("useSounds", True), default=True)
	speakStatuses = not enabled

	# 1. Obsługa nagłówków grup (np. grupowanie wg flag lub kategorii)
	cachedCls = getattr(getattr(self, "UIAElement", None), "cachedClassName", "") or ""
	rawName = getattr(getattr(self, "UIAElement", None), "cachedName", "") or getattr(getattr(self, "UIAElement", None), "currentName", "") or ""
	if not rawName:
		if _original_UIAGridRow_get_name:
			try:
				rawName = _original_UIAGridRow_get_name(self) or ""
			except Exception:
				rawName = ""

	lowerName = rawName.lower()
	isHeaderCls = cachedCls in ("ThreadHeader", "GroupHeader", "OutlookGridGroupHeader")
	isHeaderRole = getattr(self, "role", None) in (controlTypes.Role.GROUPING, controlTypes.Role.HEADER)
	isHeaderName = any(k in lowerName for k in ("grupuj według", "group by", "kategorie:", "kategoria:", "flaga:", "status flagi:"))

	if isHeaderCls or isHeaderRole or isHeaderName:
		cleanHeader = _clean_group_header_name(rawName)
		if cleanHeader:
			stateList = []
			if controlTypes.State.COLLAPSED in self.states:
				stateList.append(controlTypes.State.COLLAPSED.displayString)
			elif controlTypes.State.EXPANDED in self.states:
				stateList.append(controlTypes.State.EXPANDED.displayString)
			if stateList:
				return f"{cleanHeader}, {stateList[0]}"
			return cleanHeader

	textList = []
	if controlTypes.State.EXPANDED in self.states:
		textList.append(controlTypes.State.EXPANDED.displayString)
	elif controlTypes.State.COLLAPSED in self.states:
		textList.append(controlTypes.State.COLLAPSED.displayString)

	# Konfiguracja poszczególnych statusów (domyślne ustawienia) z jawną konwersją bool
	soundUnread = _to_bool(conf.get("soundUnread", conf.get("reportUnread", True)), default=True)
	soundFlags = _to_bool(conf.get("soundFlags", conf.get("reportFlags", False)), default=False)
	soundAttachments = _to_bool(conf.get("soundAttachments", conf.get("reportAttachments", True)), default=True)
	soundInlineImages = _to_bool(conf.get("soundInlineImages", conf.get("reportInlineImages", True)), default=True)
	soundReplied = _to_bool(conf.get("soundReplied", conf.get("soundForwardReply", conf.get("reportForwardReply", False))), default=False)
	soundForwarded = _to_bool(conf.get("soundForwarded", conf.get("soundForwardReply", conf.get("reportForwardReply", False))), default=False)
	soundImportanceHigh = _to_bool(conf.get("soundImportanceHigh", conf.get("soundImportance", conf.get("reportImportance", True))), default=True)
	soundImportanceLow = _to_bool(conf.get("soundImportanceLow", conf.get("soundImportance", conf.get("reportImportance", False))), default=False)
	soundMeetingRequest = _to_bool(conf.get("soundMeetingRequest", conf.get("reportMeetingRequest", True)), default=True)
	soundCategories = _to_bool(conf.get("soundCategories", conf.get("reportCategories", False)), default=False)
	soundExternal = _to_bool(conf.get("soundExternal", conf.get("reportExternal", False)), default=False)

	unread = False
	hasFlag = False
	realAttCount = 0
	inlineImgCount = 0
	verbText = None
	verbSound = None
	importanceText = None
	importanceSound = None
	isMeeting = False
	hasCats = False
	catText = None
	isExternalFound = False

	selection = None
	if self.appModule.nativeOm:
		try:
			explorer = self.appModule.nativeOm.activeExplorer()
			sel = getattr(explorer, "selection", None)
			# Używamy COM selection wyłącznie, gdy zaznaczony jest dokładnie 1 element
			if sel is not None and getattr(sel, "Count", 0) == 1:
				try:
					curFocus = api.getFocusObject()
					isRelated = (curFocus is self) or (getattr(curFocus, "parent", None) is self) or (getattr(self, "parent", None) is curFocus)
				except Exception:
					isRelated = True
				if isRelated:
					selection = sel.item(1)
		except (COMError, Exception):
			selection = None

	if selection:
		# 1. Stan nieprzeczytania (unread)
		try:
			unread = bool(selection.unread)
		except (COMError, Exception):
			unread = False

		# 2. Flagi
		try:
			hasFlag = bool(selection.flagIcon or selection.flagStatus)
		except (COMError, Exception):
			hasFlag = False

		# 3. Informacja o załącznikach (tylko realne pliki, bez inline grafik)
		entryID = None
		try:
			entryID = selection.EntryID
		except Exception:
			pass

		if entryID and entryID in _msg_attachment_cache:
			realAttCount, inlineImgCount = _msg_attachment_cache[entryID]
		else:
			realAttCount = _get_real_attachments_count(selection, self)
			needInline = _to_bool(conf.get("soundInlineImages", True)) or _to_bool(conf.get("reportInlineImages", True)) or speakStatuses
			inlineImgCount = _get_inline_images_count(selection) if needInline else 0
			if entryID:
				if len(_msg_attachment_cache) > 200:
					_msg_attachment_cache.clear()
				_msg_attachment_cache[entryID] = (realAttCount, inlineImgCount)

		# 4. Status odpowiedzi lub przesłania dalej (replied / forwarded)
		# Wykrywane w 100% bezpiecznie poniżej z komórek UIA bez odwoływania się do mapiObject

		# 5. Ważność / Priorytet (importance)
		try:
			importance = selection.importance
		except (COMError, Exception):
			importance = 1
		if importance == 2:
			importanceText = _t("STATUS_IMPORTANCE_HIGH")
			importanceSound = "importance_high"
		elif importance == 0:
			importanceText = _t("STATUS_IMPORTANCE_LOW")
			importanceSound = "importance_low"

		# 6. Żądanie spotkania (meeting request)
		try:
			messageClass = selection.messageClass
		except (COMError, Exception):
			messageClass = None
		if messageClass == "IPM.Schedule.Meeting.Request":
			isMeeting = True

		# 7. Kategorie
		try:
			cat = (selection.categories or "").strip()
			if cat and cat.lower() not in ("bez kategorii", "brak kategorii", "no category", "no categories", "none"):
				hasCats = True
				catText = _t("STATUS_CATEGORY", cat=cat)
		except (COMError, Exception):
			pass

	childrenCacheRequest = UIAHandler.handler.baseCacheRequest.clone()
	childrenCacheRequest.addProperty(UIAHandler.UIA_NamePropertyId)
	childrenCacheRequest.addProperty(UIAHandler.UIA_TableItemColumnHeaderItemsPropertyId)
	childrenCacheRequest.TreeScope = UIAHandler.TreeScope_Children
	childrenCacheRequest.treeFilter = builtinOutlook.createUIAMultiPropertyCondition(
		{
			UIAHandler.UIA_ControlTypePropertyId: [
				UIAHandler.UIA_TextControlTypeId,
				UIAHandler.UIA_ImageControlTypeId,
			],
		},
	)
	cachedChildren = self.UIAElement.buildUpdatedCache(childrenCacheRequest).getCachedChildren()
	if not cachedChildren:
		log.debugWarning("Unable to get relevant children for UIAGridRow", stack_info=True)
		if _original_UIAGridRow_get_name:
			try:
				return _original_UIAGridRow_get_name(self)
			except Exception:
				pass
		return getattr(getattr(self, "UIAElement", None), "cachedName", "") or ""

	contentCells = []

	for index in range(cachedChildren.length):
		e = cachedChildren.getElement(index)
		UIAControlType = e.cachedControlType
		UIAClassName = e.cachedClassName

		# Flagi
		if UIAClassName == "FlagField":
			fName = e.cachedName or ""
			if not hasFlag and fName:
				if any(w in fName.lower() for w in ("oflagow", "flaga", "flag")):
					hasFlag = True
			if not hasFlag and selection:
				try:
					hasFlag = bool(selection.flagIcon or selection.flagStatus)
				except Exception:
					pass
			continue

		# Kategorie
		elif UIAClassName == "CategoryField":
			cName = (e.cachedName or "").strip()
			if cName and cName.lower() not in ("bez kategorii", "brak kategorii", "no category", "no categories", "none"):
				hasCats = True
				if not catText:
					catText = _t("STATUS_CATEGORY", cat=cName)
			continue

		elif UIAControlType != UIAHandler.UIA_TextControlTypeId:
			continue

		name = e.cachedName
		if not name:
			continue

		cleanName = name.strip().lower()

		# Odpowiedź / przesłanie w komórce
		if cleanName in FORWARD_REPLY_TOKENS:
			if not verbText:
				verbText = name
				verbSound = "forwarded" if any(w in cleanName for w in ("przesł", "forward")) else "replied"
			continue

		# Wiadomość zewnętrzna jako osobny token/komórka
		if cleanName in EXTERNAL_TOKENS or cleanName.rstrip(":").strip() in EXTERNAL_TOKENS:
			isExternalFound = True
			continue

		# Sprawdzenie nagłówków kolumn tabeli
		columnHeaderTextList = []
		isAttachmentHeader = False
		try:
			columnHeaderItems = e.getCachedPropertyValueEx(
				UIAHandler.UIA_TableItemColumnHeaderItemsPropertyId,
				True,
			)
			if columnHeaderItems:
				columnHeaderItems = columnHeaderItems.QueryInterface(UIAHandler.IUIAutomationElementArray)
				for hIdx in range(columnHeaderItems.length):
					columnHeaderItem = columnHeaderItems.getElement(hIdx)
					hName = columnHeaderItem.currentName
					hClean = hName.strip().lower() if hName else ""
					if any(w in hClean for w in ("załącznik", "attachment")):
						isAttachmentHeader = True
						if realAttCount == 0:
							realAttCount = _extract_attachment_count(name) or 1
					if not speakStatuses:
						if _is_attachment_text(hClean) or hClean in UNREAD_TOKENS or hClean in READ_TOKENS or hClean in FORWARD_REPLY_TOKENS or hClean in EXTERNAL_TOKENS or hClean.rstrip(":").strip() in EXTERNAL_TOKENS:
							continue
					reportFlag = config.conf["documentFormatting"]["reportTableHeaders"]
					allowedFlags = (getattr(ReportTableHeaders, "ROWS_AND_COLUMNS", 2), getattr(ReportTableHeaders, "COLUMNS", 1)) if ReportTableHeaders else (2, 1)
					if reportFlag in allowedFlags:
						columnHeaderTextList.append(hName)
		except Exception:
			pass

		# Jeśli kolumna jest kolumną załączników, pomiń ją w treści, aby nie czytać surowego "1" lub "Załączniki: 1"
		if isAttachmentHeader:
			continue

		# Załączniki w komórce (wyłącznie gdy komórka jest faktycznym znacznikiem załącznika)
		if _is_attachment_text(cleanName):
			if realAttCount == 0:
				realAttCount = _extract_attachment_count(cleanName) or 1
			continue

		# Stan przeczytania / nieprzeczytania w komórce
		if cleanName in UNREAD_TOKENS:
			unread = True
			continue
		elif cleanName in READ_TOKENS:
			unread = False
			continue

		columnHeaderText = " ".join(columnHeaderTextList)
		if columnHeaderText:
			text = f"{columnHeaderText} {name}"
		else:
			text = name

		if text:
			# Wykrywanie i wycinanie prefiksów [Zewnętrzna]
			for tag in EXTERNAL_PREFIXES:
				if text.strip().lower().startswith(tag):
					isExternalFound = True
					text = text.strip()[len(tag):].lstrip(" -:,")
					break

		if text:
			contentCells.append(text + ",")

	# Weryfikacja wiadomości zewnętrznej
	isExternal = isExternalFound
	if not isExternal and selection:
		isExternal = _is_external_message(selection, self)

	# Zbuduj statusy wiersza oraz dźwięki dokładnie według CANONICAL_STATUS_ORDER
	rowStatuses = []
	sounds_to_play = []

	for status_key in CANONICAL_STATUS_ORDER:
		if status_key == "unread" and unread:
			rowStatuses.append(_t("STATUS_UNREAD"))
			if useSounds and soundUnread:
				sounds_to_play.append("unread")
		elif status_key == "flag" and hasFlag:
			rowStatuses.append(_t("STATUS_FLAG"))
			if useSounds and soundFlags:
				sounds_to_play.append("flag")
		elif status_key == "attachment" and realAttCount > 0:
			rowStatuses.append(_format_attachment_status(realAttCount))
			if useSounds and soundAttachments:
				sounds_to_play.append("attachment")
		elif status_key == "inline_image" and inlineImgCount > 0:
			rowStatuses.append(_format_inline_images_status(inlineImgCount))
			if useSounds and soundInlineImages:
				sounds_to_play.append("inline_image")
		elif status_key == "replied" and verbText and verbSound == "replied":
			rowStatuses.append(verbText)
			if useSounds and soundReplied:
				sounds_to_play.append("replied")
		elif status_key == "forwarded" and verbText and verbSound == "forwarded":
			rowStatuses.append(verbText)
			if useSounds and soundForwarded:
				sounds_to_play.append("forwarded")
		elif status_key == "importance_high" and importanceText and importanceSound == "importance_high":
			rowStatuses.append(importanceText)
			if useSounds and soundImportanceHigh:
				sounds_to_play.append("importance_high")
		elif status_key == "importance_low" and importanceText and importanceSound == "importance_low":
			rowStatuses.append(importanceText)
			if useSounds and soundImportanceLow:
				sounds_to_play.append("importance_low")
		elif status_key == "meeting" and isMeeting:
			rowStatuses.append(_t("STATUS_MEETING"))
			if useSounds and soundMeetingRequest:
				sounds_to_play.append("meeting")
		elif status_key == "category" and (catText or hasCats):
			cLabel = catText or _t("STATUS_CATEGORY", cat="")
			rowStatuses.append(cLabel)
			if useSounds and soundCategories:
				sounds_to_play.append("category")
		elif status_key == "external" and isExternal:
			rowStatuses.append(_t("STATUS_EXTERNAL"))
			if useSounds and soundExternal:
				sounds_to_play.append("external")

	self._osf_statuses = rowStatuses
	self._omc_statuses = rowStatuses
	global _last_focused_statuses
	_last_focused_statuses = rowStatuses

	# Odtwórz dźwięki statusów w kanonicznej kolejności
	if sounds_to_play:
		_schedule_status_sounds(sounds_to_play)

	# Gdy komunikaty mowy są włączone (speakStatuses = True), dodaj statusy na początku tekstu wiersza
	if speakStatuses and rowStatuses:
		textList.extend([s + "," for s in rowStatuses])

	textList.extend(contentCells)

	return " ".join(textList)


def _custom_getReceivedMessageString(obj):
	conf = config.conf.get(CONFIG_SECTION, {})
	enabled = _to_bool(conf.get("enabled", True), default=True)
	useSounds = _to_bool(conf.get("useSounds", True), default=True)
	speakStatuses = not enabled

	sounds_to_play = []
	nameList = []
	entryID = getattr(obj, "EntryID", None)
	if entryID and entryID in _msg_attachment_cache:
		realAttCount, inlineImgCount = _msg_attachment_cache[entryID]
	else:
		realAttCount = _get_real_attachments_count(obj)
		needInline = _to_bool(conf.get("soundInlineImages", conf.get("reportInlineImages", True)), default=True) or speakStatuses
		inlineImgCount = _get_inline_images_count(obj) if needInline else 0
		if entryID:
			if len(_msg_attachment_cache) > 200:
				_msg_attachment_cache.clear()
			_msg_attachment_cache[entryID] = (realAttCount, inlineImgCount)

	# Zbuduj statusy wiersza według CANONICAL_STATUS_ORDER
	soundUnread = _to_bool(conf.get("soundUnread", conf.get("reportUnread", True)), default=True)
	soundAttachments = _to_bool(conf.get("soundAttachments", conf.get("reportAttachments", True)), default=True)
	soundInlineImages = _to_bool(conf.get("soundInlineImages", conf.get("reportInlineImages", True)), default=True)

	if getattr(obj, "unread", False):
		if speakStatuses:
			nameList.append(_t("STATUS_UNREAD"))
		if useSounds and soundUnread:
			sounds_to_play.append("unread")

	if realAttCount > 0:
		if speakStatuses:
			nameList.append(_format_attachment_status(realAttCount))
		if useSounds and soundAttachments:
			sounds_to_play.append("attachment")

	if inlineImgCount > 0:
		if speakStatuses:
			nameList.append(_format_inline_images_status(inlineImgCount))
		if useSounds and soundInlineImages:
			sounds_to_play.append("inline_image")

	sender = ""
	try:
		sender = getattr(obj, "senderName", "") or ""
	except Exception:
		pass
	if sender:
		nameList.append(sender)

	try:
		subj = getattr(obj, "subject", "") or ""
		if subj:
			nameList.append(f"temat: {subj}")
	except Exception:
		pass

	try:
		rec = getattr(obj, "receivedTime", "") or ""
		if rec:
			nameList.append(f"odebrano: {rec}")
	except Exception:
		pass

	text = ", ".join(nameList)

	if sounds_to_play:
		_schedule_status_sounds(sounds_to_play)

	return text


_saved_class_methods = {}


def _safe_patch_method(target_obj, attr_name, new_method):
	"""Bezpiecznie podmienia metodę klasy, zachowując indywidualny oryginał dla danej klasy/obiektu."""
	key = (target_obj, attr_name)
	if key not in _saved_class_methods:
		orig = getattr(target_obj, attr_name, None)
		if orig is not None:
			_saved_class_methods[key] = orig
	setattr(target_obj, attr_name, new_method)


def patchOutlookClasses():
	"""Podmienia _get_name oraz getReceivedMessageString na klasach wierszy Outlooka, zachowując oryginalne referencje."""
	global _original_UIAGridRow_get_name, _original_getReceivedMessageString, _patches_applied
	if _patches_applied:
		return

	if builtinOutlook and hasattr(builtinOutlook, "UIAGridRow"):
		if _original_UIAGridRow_get_name is None:
			_original_UIAGridRow_get_name = builtinOutlook.UIAGridRow._get_name
		_safe_patch_method(builtinOutlook.UIAGridRow, "_get_name", _custom_UIAGridRow_get_name)

		if hasattr(builtinOutlook, "getReceivedMessageString"):
			if _original_getReceivedMessageString is None:
				_original_getReceivedMessageString = builtinOutlook.getReceivedMessageString
			_safe_patch_method(builtinOutlook, "getReceivedMessageString", _custom_getReceivedMessageString)

	# Bezpieczne sprawdzanie załadowanych modułów w sys.modules
	for modName in list(sys.modules.keys()):
		if "appModules.outlook" in modName:
			mod = sys.modules.get(modName)
			if mod:
				if hasattr(mod, "UIAGridRow"):
					_safe_patch_method(mod.UIAGridRow, "_get_name", _custom_UIAGridRow_get_name)
				if hasattr(mod, "UIAGridRowWithReadStatus"):
					_safe_patch_method(mod.UIAGridRowWithReadStatus, "_get_name", _custom_UIAGridRow_get_name)

	_patches_applied = True


class CleanOutlookGridRow(builtinOutlook.UIAGridRow if builtinOutlook else object):
	def _get_name(self):
		return _custom_UIAGridRow_get_name(self)

	def _get_description(self):
		rawName = getattr(getattr(self, "UIAElement", None), "cachedName", "") or getattr(getattr(self, "UIAElement", None), "currentName", "") or ""
		if not rawName and _original_UIAGridRow_get_name:
			try:
				rawName = _original_UIAGridRow_get_name(self) or ""
			except Exception:
				rawName = ""
		if rawName and any(k in rawName.lower() for k in ("grupuj według", "group by", "kategorie:", "kategoria:")):
			return ""
		try:
			rawDesc = super().description
			if rawDesc and any(k in rawDesc.lower() for k in ("grupuj według", "group by")):
				return ""
			return rawDesc
		except Exception:
			return ""


PARENT_SOUNDS_LABEL = _("Odtwarzaj dyskretne dźwięki zamiast wyciszonych komunikatów mowy")


class SettingsPanelTitleDescriptor:
	"""
	Deskryptor dynamicznego tytułu kategorii ustawień w NVDA oraz zdarzeń wejścia.
	Zwraca zlokalizowaną nazwę panelu zależnie od wybranego języka:
	- dla języka polskiego: 'Outlook - Filtr mowy (Outlook - Speech Filter)'
	- dla języka angielskiego: 'Outlook - Speech Filter'
	"""
	def __get__(self, obj, objtype=None):
		return _t("TITLE")


# ==============================================================================
# MODUŁ AKTUALIZACJI DODATKU (WCAG 2.1 / 2.2 AA)
# ==============================================================================

def _parse_version(v_str: str) -> tuple:
	parts = re.findall(r'\d+', str(v_str or ""))
	return tuple(int(p) for p in parts) if parts else (0,)

def _is_newer(remote_v: str, local_v: str) -> bool:
	return _parse_version(remote_v) > _parse_version(local_v)


class UpdateNotificationDialog(wx.Dialog):
	"""
	W pełni dostępne okno powiadomienia o aktualizacji (WCAG 2.1 / 2.2 AA).
	Zapewnia pełną obsługę klawiatury (Enter, Escape, Tab), wysoki kontrast
	oraz czytelny odczyt listy zmian.
	"""
	def __init__(self, parent, addon_title, local_ver, remote_ver, changelog, on_install_callback):
		super().__init__(
			parent,
			title=f"Aktualizacja: {addon_title}",
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER
		)
		self.on_install_callback = on_install_callback

		main_sizer = wx.BoxSizer(wx.VERTICAL)

		# Nagłówek
		lbl_header = wx.StaticText(self, label=f"Dostępna jest nowa wersja dodatku: {addon_title}!")
		font = lbl_header.GetFont()
		font.SetWeight(wx.FONTWEIGHT_BOLD)
		lbl_header.SetFont(font)
		main_sizer.Add(lbl_header, 0, wx.ALL, 12)

		# Informacja o wersjach
		lbl_ver = wx.StaticText(self, label=f"Zainstalowana wersja: {local_ver}   --->   Nowa wersja: {remote_ver}")
		main_sizer.Add(lbl_ver, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		# Etykieta zmian
		lbl_changes = wx.StaticText(self, label="&Wykaz nowości i ulepszeń w tej wersji:")
		main_sizer.Add(lbl_changes, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 4)

		# Pole tekstowe z opisem zmian (dostępne, wieloliniowe, tylko do odczytu)
		self.txt_changelog = wx.TextCtrl(
			self,
			value=changelog or "Brak dodatkowego opisu zmian.",
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL,
			size=(520, 160)
		)
		main_sizer.Add(self.txt_changelog, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 12)

		# Komunikat pytający
		lbl_prompt = wx.StaticText(self, label="Czy chcesz teraz pobrać i zainstalować tę aktualizację?")
		main_sizer.Add(lbl_prompt, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 12)

		# Przyciski akcji
		btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
		self.btn_install = wx.Button(self, wx.ID_YES, label="&Zaktualizuj teraz")
		self.btn_cancel = wx.Button(self, wx.ID_CANCEL, label="&Przypomnij później")

		self.btn_install.Bind(wx.EVT_BUTTON, self.on_install)
		self.btn_cancel.Bind(wx.EVT_BUTTON, self.on_cancel)

		btn_sizer.Add(self.btn_install, 0, wx.RIGHT, 10)
		btn_sizer.Add(self.btn_cancel, 0)
		main_sizer.Add(btn_sizer, 0, wx.ALIGN_RIGHT | wx.ALL, 12)

		self.SetSizer(main_sizer)
		main_sizer.Fit(self)
		self.CenterOnScreen()
		self.btn_install.SetDefault()
		self.btn_install.SetFocus()

	def on_install(self, event):
		self.EndModal(wx.ID_YES)
		if callable(self.on_install_callback):
			wx.CallAfter(self.on_install_callback)

	def on_cancel(self, event):
		self.EndModal(wx.ID_CANCEL)


class AddonUpdaterEngine:
	def __init__(self, addon_name: str, addon_title: str, current_version: str, addon_dir: str, config_section: str):
		self.addon_name = addon_name
		self.addon_title = addon_title
		self.current_version = current_version
		self.addon_dir = addon_dir
		self.config_section = config_section

	def get_configured_source(self) -> str:
		conf = config.conf.get(self.config_section, {})
		return conf.get("updateSourceUrl", "https://raw.githubusercontent.com/strugo10/outlookSpeechFilter/main/update_info.json")

	def fetch_manifest(self) -> tuple[dict | None, str | None]:
		configured = self.get_configured_source()
		candidates = []
		if configured:
			candidates.append(configured.strip())
		candidates.extend([
			"https://raw.githubusercontent.com/strugo10/outlookSpeechFilter/main/update_info.json",
			"https://raw.githubusercontent.com/strugo10/outlookSpeechFilter/main/updates.json",
			r"G:\Mój dysk\Antigravity\Nvda dodatki\updates.json",
			os.path.join(self.addon_dir, "update_info.json"),
			os.path.join(self.addon_dir, "updates.json"),
			os.path.join(os.path.dirname(self.addon_dir), "updates.json"),
			r"C:\Antigravity\Nvda dodatki\updates.json",
		])

		raw_content = None
		resolved_source = None
		last_err = None

		for s in candidates:
			if not s:
				continue
			try:
				if s.startswith("http://") or s.startswith("https://"):
					req = urllib.request.Request(s, headers={"User-Agent": "NVDA-Addon-Updater/1.0"})
					with urllib.request.urlopen(req, timeout=10) as resp:
						raw_content = resp.read().decode("utf-8-sig")
						resolved_source = s
						break
				elif os.path.isfile(s):
					with open(s, "r", encoding="utf-8-sig") as f:
						raw_content = f.read()
						resolved_source = s
						break
			except Exception as ex:
				last_err = ex
				continue

		if not raw_content:
			return None, str(last_err or "Nie znaleziono pliku ogłoszeniowego na dysku ani pod wskazanym adresem.")

		try:
			data = json.loads(raw_content)
			if self.addon_name in data and isinstance(data[self.addon_name], dict):
				item = data[self.addon_name]
			elif isinstance(data, dict) and "version" in data:
				item = data
			else:
				return None, f"Brak wpisu dla dodatku '{self.addon_name}' w pliku ogłoszeniowym."

			item["_source"] = resolved_source
			return item, None
		except Exception as ex:
			return None, f"Błąd przetwarzania formatu JSON: {ex}"

	def check_for_updates(self, manual: bool = False):
		t = threading.Thread(target=self._worker_check, args=(manual,), daemon=True)
		t.start()

	def _worker_check(self, manual: bool):
		try:
			info, err = self.fetch_manifest()
			if not info:
				if manual:
					msg = f"Nie udało się sprawdzić aktualizacji dla dodatku {self.addon_title}.\n\nSzczegóły: {err}"
					wx.CallAfter(gui.messageBox, msg, f"Aktualizacje: {self.addon_title}", wx.OK | wx.ICON_WARNING)
				return

			remote_ver = str(info.get("version", "")).strip()
			if _is_newer(remote_ver, self.current_version):
				try:
					config.conf[self.config_section]["lastUpdateCheckTime"] = time.time()
				except Exception:
					pass
				download_url = info.get("downloadUrl", "")
				changelog = info.get("changelog", "")

				def _show_dialog():
					def _on_confirm():
						self.perform_update(download_url)
					dlg = UpdateNotificationDialog(
						parent=gui.mainFrame,
						addon_title=self.addon_title,
						local_ver=self.current_version,
						remote_ver=remote_ver,
						changelog=changelog,
						on_install_callback=_on_confirm
					)
					gui.mainFrame.prePopup()
					dlg.ShowModal()
					dlg.Destroy()
					gui.mainFrame.postPopup()

				wx.CallAfter(_show_dialog)
			else:
				if manual:
					msg = f"Posiadasz najnowszą wersję dodatku {self.addon_title} ({self.current_version}).\nBrak nowszych aktualizacji na Twoim dysku."
					wx.CallAfter(gui.messageBox, msg, f"Aktualizacje: {self.addon_title}", wx.OK | wx.ICON_INFORMATION)
		except Exception as ex:
			if manual:
				wx.CallAfter(gui.messageBox, f"Wystąpił nieoczekiwany błąd podczas sprawdzania aktualizacji:\n{ex}", f"Błąd: {self.addon_title}", wx.OK | wx.ICON_ERROR)

	def perform_update(self, download_url: str):
		t = threading.Thread(target=self._worker_download_install, args=(download_url,), daemon=True)
		t.start()

	def _worker_download_install(self, download_url: str):
		import addonHandler

		ui.message(f"Pobieranie aktualizacji dla {self.addon_title}...")
		target_pkg = None

		try:
			if download_url.startswith("http://") or download_url.startswith("https://"):
				req = urllib.request.Request(download_url, headers={"User-Agent": "NVDA-Addon-Updater/1.0"})
				tmp_file = os.path.join(tempfile.gettempdir(), f"{self.addon_name}_update.nvda-addon")
				with urllib.request.urlopen(req, timeout=30) as resp, open(tmp_file, "wb") as out:
					shutil.copyfileobj(resp, out)
				target_pkg = tmp_file
			else:
				candidates = [download_url]
				candidates.extend([
					os.path.join(r"G:\Mój dysk\Antigravity\Nvda dodatki\Pakiety_Instalacyjne", f"{self.addon_name}.nvda-addon"),
					os.path.join(r"C:\Antigravity\Nvda dodatki\Pakiety_Instalacyjne", f"{self.addon_name}.nvda-addon"),
				])
				for cand in candidates:
					if cand and os.path.isfile(cand):
						target_pkg = cand
						break

			if not target_pkg or not os.path.isfile(target_pkg):
				raise FileNotFoundError(f"Nie znaleziono pliku instalacyjnego dodatku:\n{download_url}")

			def _do_install():
				try:
					if hasattr(addonHandler, "installAddonPackage"):
						try:
							bundle = addonHandler.AddonBundle(target_pkg)
							addonHandler.installAddonPackage(bundle)
							return
						except TypeError:
							addonHandler.installAddonPackage(target_pkg)
							return
					elif hasattr(addonHandler, "installAddon"):
						bundle = addonHandler.AddonBundle(target_pkg)
						addonHandler.installAddon(bundle)
						return
				except Exception as ex:
					log.debug(f"{self.addon_name} updater internal install error: {ex}")
				try:
					os.startfile(target_pkg)
				except Exception as ex2:
					gui.messageBox(f"Nie udało się uruchomić instalatora:\n{ex2}", "Błąd instalacji", wx.OK | wx.ICON_ERROR)

			wx.CallAfter(_do_install)
		except Exception as ex:
			wx.CallAfter(gui.messageBox, f"Nie udało się pobrać aktualizacji:\n{ex}", f"Błąd aktualizacji: {self.addon_title}", wx.OK | wx.ICON_ERROR)


class OutlookSpeechFilterSettingsPanel(SettingsPanel):
	title = SettingsPanelTitleDescriptor()

	def makeSettings(self, settingsSizer):
		global _activeSettingsPanelRef
		_activeSettingsPanelRef = weakref.ref(self)
		try:
			self._makeSettingsInternal(settingsSizer)
		except Exception as e:
			log.error(f"outlookSpeechFilter: Błąd podczas budowania panelu ustawień: {e}", exc_info=True)
			errLabel = wx.StaticText(self, label=f"Wystąpił błąd podczas ładowania panelu ustawień Outlook Speech Filter: {e}")
			settingsSizer.Add(errLabel, flag=wx.ALL, border=15)

	def _makeSettingsInternal(self, settingsSizer):
		sHelper = gui.guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
		try:
			conf = config.conf[CONFIG_SECTION]
		except Exception:
			conf = {}

		# 0. Wybór języka interfejsu dodatku (Automatyczny / Polski / English)
		langBox = wx.BoxSizer(wx.HORIZONTAL)
		self.langLabel = wx.StaticText(self, label=_t("LANG_LABEL"))
		langBox.Add(self.langLabel, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

		langChoices = [
			_t("LANG_AUTO"),
			_t("LANG_PL"),
			_t("LANG_EN"),
		]
		self.langKeys = ["auto", "pl", "en"]
		currentLang = str(conf.get("language", "auto"))
		if currentLang not in self.langKeys:
			currentLang = "auto"
		self.initialLanguage = currentLang
		selectedIdx = self.langKeys.index(currentLang)

		self.languageChoice = wx.Choice(self, choices=langChoices)
		self.languageChoice.SetName(_t("LANG_LABEL"))
		self.languageChoice.SetSelection(selectedIdx)
		self.languageChoice.Bind(wx.EVT_CHOICE, self.onLanguageChanged)
		langBox.Add(self.languageChoice, flag=wx.ALIGN_CENTER_VERTICAL)
		sHelper.addItem(langBox)

		self.enabledCB = sHelper.addItem(
			wx.CheckBox(self, label=_t("ENABLE_FILTER"))
		)
		self.enabledCB.SetValue(_to_bool(conf.get("enabled", True), default=True))

		self.threadNavCB = sHelper.addItem(
			wx.CheckBox(self, label=_t("ENABLE_THREAD_NAV"))
		)
		self.threadNavCB.SetValue(_to_bool(conf.get("threadNavigationEnabled", True), default=True))

		# Translators: Główna opcja nadrzędna odtwarzania dyskretnych dźwięków zamiast komunikatów mowy
		self.useSoundsCB = sHelper.addItem(
			wx.CheckBox(self, label=_t("PARENT_SOUNDS_LABEL"))
		)
		self.useSoundsCB.SetValue(_to_bool(conf.get("useSounds", True), default=True))
		self.useSoundsCB.Bind(wx.EVT_CHECKBOX, self.onUseSoundsChanged)

		# Sekcja dźwięków (widoczna tylko gdy włączono dźwięki)
		self.soundBoxSizer = wx.BoxSizer(wx.VERTICAL)
		self.soundControls = []
		self.childSoundCBs = []

		# Wybór motywu dźwięków (asocjacyjny, minimalistyczny, instrumentalny)
		themeBox = wx.BoxSizer(wx.HORIZONTAL)
		self.themeLabel = wx.StaticText(self, label=_t("SOUND_THEME_LABEL"))
		themeBox.Add(self.themeLabel, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

		currentTheme = conf.get("soundTheme", "associative")
		self.initialSoundTheme = currentTheme
		self.themeKeys = ["associative", "minimal", "instrumental"]
		themeLabels = [
			_t("THEME_ASSOCIATIVE"),
			_t("THEME_MINIMAL"),
			_t("THEME_INSTRUMENTAL"),
		]
		self.soundThemeChoice = wx.Choice(self, choices=themeLabels)
		self.soundThemeChoice.SetName(_t("SOUND_THEME_LABEL"))
		themeIdx = 0
		if currentTheme in self.themeKeys:
			themeIdx = self.themeKeys.index(currentTheme)
		self.soundThemeChoice.SetSelection(themeIdx)
		self.soundThemeChoice.Bind(wx.EVT_CHOICE, self.onSoundThemeChanged)
		themeBox.Add(self.soundThemeChoice, flag=wx.ALIGN_CENTER_VERTICAL)

		self.soundBoxSizer.Add(themeBox, flag=wx.LEFT | wx.TOP | wx.BOTTOM, border=5)
		self.soundControls.extend([self.themeLabel, self.soundThemeChoice])

		# 1. Treść wyjaśniająca pod przełącznikiem nadrzędnym
		self.soundExplText = wx.StaticText(
			self,
			label=_t("SOUND_EXPL_TEXT")
		)
		self.soundBoxSizer.Add(self.soundExplText, flag=wx.LEFT | wx.TOP | wx.BOTTOM, border=5)
		self.soundControls.append(self.soundExplText)

		# 2. Przyciski zbiorcze: Zaznacz wszystkie / Odznacz wszystkie (umieszczone bezpośrednio pod opisem)
		bulkBox = wx.BoxSizer(wx.HORIZONTAL)
		self.selectAllBtn = wx.Button(self, label=_t("SELECT_ALL_SOUNDS"))
		self.selectAllBtn.Bind(wx.EVT_BUTTON, self.onSelectAllSounds)
		bulkBox.Add(self.selectAllBtn, flag=wx.RIGHT, border=10)

		self.deselectAllBtn = wx.Button(self, label=_t("DESELECT_ALL_SOUNDS"))
		self.deselectAllBtn.Bind(wx.EVT_BUTTON, self.onDeselectAllSounds)
		bulkBox.Add(self.deselectAllBtn)

		self.soundBoxSizer.Add(bulkBox, flag=wx.LEFT | wx.BOTTOM, border=20)
		self.soundControls.extend([self.selectAllBtn, self.deselectAllBtn])

		# 3. Podrzędna lista poszczególnych dźwięków z wcięciem 20px, w kanonicznej kolejności
		soundListSizer = wx.BoxSizer(wx.VERTICAL)
		self.soundChildItems = []

		def addSoundOptionWithPlayBtn(cbKey, confKey, soundName, defaultVal=True, btnKey=None, fallbackKeys=None):
			box = wx.BoxSizer(wx.HORIZONTAL)
			cb = wx.CheckBox(self, label=_t(cbKey))
			val = conf.get(confKey, None)
			if val is None and fallbackKeys:
				for fk in fallbackKeys:
					if fk in conf:
						val = conf[fk]
						break
			if val is None:
				val = conf.get("report" + confKey[5:], defaultVal)
			cb.SetValue(_to_bool(val, default=defaultVal))
			cb.Bind(wx.EVT_CHECKBOX, self.onChildSoundChanged)
			box.Add(cb, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

			bKey = btnKey or "PLAY_GENERIC"
			btn = wx.Button(self, label=_t(bKey))
			# Niezależnie od suwaka, przy każdym pojedynczym dźwięku odtwarzana jest głośność maksymalna (100)
			btn.Bind(wx.EVT_BUTTON, lambda evt, s=soundName: self._playSingleSoundMaxVol(s))
			box.Add(btn, flag=wx.ALIGN_CENTER_VERTICAL)

			soundListSizer.Add(box, flag=wx.BOTTOM, border=5)
			self.childSoundCBs.append(cb)
			self.soundControls.extend([cb, btn])
			self.soundChildItems.append((cb, cbKey, btn, bKey))
			return cb

		# 1. Nieprzeczytana (domyślnie włączony)
		self.soundUnreadCB = addSoundOptionWithPlayBtn(
			"SOUND_UNREAD",
			"soundUnread",
			"unread",
			defaultVal=True,
			btnKey="PLAY_UNREAD",
		)

		# 2. Flagi (domyślnie wyłączony)
		self.soundFlagsCB = addSoundOptionWithPlayBtn(
			"SOUND_FLAG",
			"soundFlags",
			"flag",
			defaultVal=False,
			btnKey="PLAY_FLAG",
		)

		# 3. Załączniki (domyślnie włączony)
		self.soundAttachmentsCB = addSoundOptionWithPlayBtn(
			"SOUND_ATTACHMENT",
			"soundAttachments",
			"attachment",
			defaultVal=True,
			btnKey="PLAY_ATTACHMENT",
		)

		# 3b. Wklejona grafika / zrzut ekranu (domyślnie włączony)
		self.soundInlineImagesCB = addSoundOptionWithPlayBtn(
			"SOUND_INLINE_IMAGE",
			"soundInlineImages",
			"inline_image",
			defaultVal=True,
			btnKey="PLAY_INLINE_IMAGE",
		)

		# 4. Odpowiedź na wiadomość (domyślnie wyłączony)
		self.soundRepliedCB = addSoundOptionWithPlayBtn(
			"SOUND_REPLIED",
			"soundReplied",
			"replied",
			defaultVal=False,
			btnKey="PLAY_REPLIED",
			fallbackKeys=["soundForwardReply", "reportForwardReply"],
		)

		# 5. Przesłanie wiadomości dalej (domyślnie wyłączony)
		self.soundForwardedCB = addSoundOptionWithPlayBtn(
			"SOUND_FORWARDED",
			"soundForwarded",
			"forwarded",
			defaultVal=False,
			btnKey="PLAY_FORWARDED",
			fallbackKeys=["soundForwardReply", "reportForwardReply"],
		)

		# 6. Wysoki priorytet wiadomości (domyślnie włączony)
		self.soundImportanceHighCB = addSoundOptionWithPlayBtn(
			"SOUND_IMPORTANCE_HIGH",
			"soundImportanceHigh",
			"importance_high",
			defaultVal=True,
			btnKey="PLAY_IMPORTANCE_HIGH",
			fallbackKeys=["soundImportance", "reportImportance"],
		)

		# 7. Niski priorytet wiadomości (domyślnie wyłączony)
		self.soundImportanceLowCB = addSoundOptionWithPlayBtn(
			"SOUND_IMPORTANCE_LOW",
			"soundImportanceLow",
			"importance_low",
			defaultVal=False,
			btnKey="PLAY_IMPORTANCE_LOW",
			fallbackKeys=["soundImportance", "reportImportance"],
		)

		# 8. Żądanie spotkania (domyślnie włączony)
		self.soundMeetingRequestCB = addSoundOptionWithPlayBtn(
			"SOUND_MEETING",
			"soundMeetingRequest",
			"meeting",
			defaultVal=True,
			btnKey="PLAY_MEETING",
		)

		# 9. Kategorie (domyślnie wyłączony)
		self.soundCategoriesCB = addSoundOptionWithPlayBtn(
			"SOUND_CATEGORY",
			"soundCategories",
			"category",
			defaultVal=False,
			btnKey="PLAY_CATEGORY",
		)

		# 10. Wiadomość zewnętrzna (domyślnie wyłączony)
		self.soundExternalCB = addSoundOptionWithPlayBtn(
			"SOUND_EXTERNAL",
			"soundExternal",
			"external",
			defaultVal=False,
			btnKey="PLAY_EXTERNAL",
		)

		self.soundBoxSizer.Add(soundListSizer, flag=wx.LEFT | wx.BOTTOM, border=20)

		# 4. Sekcja głośności dźwięków (przeniesiona POD poszczególne dźwięki) oraz przycisk pełnego testu sekwencyjnego
		volBox = wx.BoxSizer(wx.HORIZONTAL)
		self.volLabel = wx.StaticText(self, label=_t("VOL_LABEL"))
		volBox.Add(self.volLabel, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

		currentVol = conf.get("soundVolume", 80)
		try:
			currentVol = int(currentVol)
		except (ValueError, TypeError):
			currentVol = 80
		self.initialSoundVolume = currentVol

		self.soundVolumeSlider = wx.Slider(
			self,
			value=currentVol,
			minValue=0,
			maxValue=100,
			style=wx.SL_HORIZONTAL,
		)
		self.soundVolumeSlider.SetName(_t("SLIDER_NAME"))
		self.soundVolumeSlider.SetPageSize(10)
		self.soundVolumeSlider.SetLineSize(1)
		self.soundVolumeSlider.Bind(wx.EVT_SLIDER, self.onSoundVolumeChanged)
		self.soundVolumeSlider.Bind(wx.EVT_KEY_DOWN, self.onSliderKeyDown)
		volBox.Add(self.soundVolumeSlider, proportion=1, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

		self.volValueText = wx.StaticText(self, label=f"{currentVol}%")
		volBox.Add(self.volValueText, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

		self.testVolBtn = wx.Button(self, label=_t("TEST_VOL_BTN"))
		self.testVolBtn.Bind(wx.EVT_BUTTON, self.onTestSoundVolume)
		volBox.Add(self.testVolBtn, flag=wx.ALIGN_CENTER_VERTICAL)

		self.soundBoxSizer.Add(volBox, flag=wx.EXPAND | wx.LEFT | wx.BOTTOM, border=20)
		self.soundControls.extend([self.volLabel, self.soundVolumeSlider, self.volValueText, self.testVolBtn])

		settingsSizer.Add(self.soundBoxSizer, flag=wx.EXPAND | wx.BOTTOM, border=10)

		# Sekcja aktualizacji dodatku (WCAG 2.1 / 2.2 AA)
		settingsSizer.AddSpacer(10)
		self.chkAutoUpdates = sHelper.addItem(
			wx.CheckBox(self, label="&Automatycznie sprawdzaj dostępność aktualizacji przy uruchomieniu NVDA")
		)
		self.chkAutoUpdates.SetValue(_to_bool(conf.get("autoCheckUpdates", True), default=True))

		self.txtUpdateSource = sHelper.addLabeledControl(
			"&Lokalizacja pliku ogłoszeniowego (ścieżka na dysku lub URL):",
			wx.TextCtrl,
			value=conf.get("updateSourceUrl", "https://raw.githubusercontent.com/strugo10/outlookSpeechFilter/main/update_info.json")
		)

		# Sekcja narzędziowa: Pomoc kontekstowa, Samouczek / Pomoc HTML, Kontakt z autorem, Skróty klawiszowe oraz Przywracanie ustawień domyślnych
		toolsBox = wx.BoxSizer(wx.HORIZONTAL)
		self.checkUpdatesBtn = wx.Button(self, label="&Sprawdź dostępność aktualizacji teraz")
		self.checkUpdatesBtn.Bind(wx.EVT_BUTTON, self.onCheckUpdatesNow)
		toolsBox.Add(self.checkUpdatesBtn, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=8)

		self.contextHelpBtn = wx.Button(self, label=_t("CONTEXT_HELP_BTN"))
		self.contextHelpBtn.Bind(wx.EVT_BUTTON, lambda evt: showContextHelp("settings", parent=self))
		toolsBox.Add(self.contextHelpBtn, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=8)

		self.helpBtn = wx.Button(self, label=_t("HELP_BTN"))
		self.helpBtn.Bind(wx.EVT_BUTTON, self.onOpenHelp)
		toolsBox.Add(self.helpBtn, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=8)

		self.contactBtn = wx.Button(self, label=_t("CONTACT_BTN"))
		self.contactBtn.Bind(wx.EVT_BUTTON, self.onContactAuthor)
		toolsBox.Add(self.contactBtn, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=8)

		self.gesturesBtn = wx.Button(self, label=_t("GESTURES_BTN"))
		self.gesturesBtn.Bind(wx.EVT_BUTTON, self.onOpenInputGestures)
		toolsBox.Add(self.gesturesBtn, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=8)

		self.restoreDefaultsBtn = wx.Button(self, label=_t("RESTORE_DEFAULTS_BTN"))
		self.restoreDefaultsBtn.Bind(wx.EVT_BUTTON, self.onRestoreDefaults)
		toolsBox.Add(self.restoreDefaultsBtn, flag=wx.ALIGN_CENTER_VERTICAL)
		sHelper.addItem(toolsBox)

		# Informacja o wersji, buildzie i autorach
		self.infoText = wx.StaticText(
			self,
			label=_t(
				"ADDON_INFO_LINE",
				name=ADDON_NAME,
				ver=ADDON_VERSION,
				build=ADDON_BUILD,
				date=ADDON_BUILD_DATE,
				author=ADDON_AUTHOR,
			),
		)
		sHelper.addItem(self.infoText)

		# Zapamiętanie stanu zaznaczenia dźwięków do ewentualnego przywrócenia
		self._lastActiveChildStates = [cb.GetValue() for cb in self.childSoundCBs]
		if not any(self._lastActiveChildStates):
			self._lastActiveChildStates = [True] * len(self.childSoundCBs)

		# Ustawienie początkowej widoczności i stanów kontrolek sekcji dźwiękowej
		showSounds = self.useSoundsCB.GetValue()
		self.showSoundsPanel(showSounds)
		self.updateSoundControlsState()
		self.updateParentSoundsLabel()
		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)

	def onCheckUpdatesNow(self, evt):
		updater = AddonUpdaterEngine(
			addon_name=CONFIG_SECTION,
			addon_title=ADDON_NAME,
			current_version=ADDON_VERSION,
			addon_dir=ADDON_DIR,
			config_section=CONFIG_SECTION
		)
		updater.check_for_updates(manual=True)

	def onCharHook(self, event):
		if event.GetKeyCode() == wx.WXK_F1:
			showContextHelp("settings", parent=self)
			return
		event.Skip()

	def _playSingleSoundMaxVol(self, soundName):
		"""Odtwarza pojedynczy dźwięk w ustawieniach z pełną głośnością 100%."""
		_stop_pending_sounds()
		_play_status_sound(soundName, volume=100)

	def setCleanerEnabledState(self, state):
		if hasattr(self, "enabledCB") and self.enabledCB:
			self.enabledCB.SetValue(_to_bool(state, default=True))

	def setThreadNavEnabledState(self, state):
		if hasattr(self, "threadNavCB") and self.threadNavCB:
			self.threadNavCB.SetValue(_to_bool(state, default=True))

	def setUseSoundsState(self, state):
		if hasattr(self, "useSoundsCB") and self.useSoundsCB:
			self.useSoundsCB.SetValue(_to_bool(state, default=True))
			self.onUseSoundsChanged(None)

	def showSoundsPanel(self, show):
		"""Płynnie ukrywa lub pokazuje cały panel dźwiękowy i odświeża układ okna."""
		sizer = self.GetSizer()
		if sizer and hasattr(self, "soundBoxSizer") and self.soundBoxSizer:
			sizer.Show(self.soundBoxSizer, show, recursive=True)
		if hasattr(self, "soundControls"):
			for ctrl in self.soundControls:
				try:
					ctrl.Show(show)
				except Exception:
					pass
		self.Layout()
		parent = self.GetParent()
		if parent:
			try:
				parent.Layout()
			except Exception:
				pass

	def updateSoundControlsState(self):
		"""Aktualizuje dostępność przycisków zbiorczych Zaznacz / Odznacz wszystkie dźwięki."""
		total = len(self.childSoundCBs)
		selected = sum(1 for cb in self.childSoundCBs if cb.GetValue())
		if hasattr(self, "selectAllBtn") and self.selectAllBtn:
			self.selectAllBtn.Enable(selected < total)
		if hasattr(self, "deselectAllBtn") and self.deselectAllBtn:
			self.deselectAllBtn.Enable(selected > 0)

	def updateParentSoundsLabel(self):
		total = len(self.childSoundCBs)
		selected = sum(1 for cb in self.childSoundCBs if cb.GetValue())
		baseLabel = _t("PARENT_SOUNDS_LABEL")
		if not self.useSoundsCB.GetValue():
			self.useSoundsCB.SetLabel(baseLabel)
		elif selected == total:
			self.useSoundsCB.SetLabel(baseLabel)
		elif selected == 0:
			self.useSoundsCB.SetLabel(baseLabel + _t("PARENT_ALL_DISABLED"))
		else:
			self.useSoundsCB.SetLabel(baseLabel + _t("PARENT_PARTIAL"))

	def onSoundThemeChanged(self, evt):
		idx = self.soundThemeChoice.GetSelection()
		if 0 <= idx < len(self.themeKeys):
			chosenTheme = self.themeKeys[idx]
			try:
				config.conf[CONFIG_SECTION]["soundTheme"] = chosenTheme
			except Exception:
				pass
			_stop_pending_sounds()
			_play_status_sound("attachment", volume=100)

	def updateUILabels(self):
		"""Aktualizuje etykiety wszystkich kontrolek w otwartym panelu na bieżąco po zmianie języka."""
		if hasattr(self, "langLabel") and self.langLabel:
			self.langLabel.SetLabel(_t("LANG_LABEL"))
		if hasattr(self, "languageChoice") and self.languageChoice:
			self.languageChoice.SetName(_t("LANG_LABEL"))
			try:
				self.languageChoice.SetString(0, _t("LANG_AUTO"))
			except Exception:
				pass
		if hasattr(self, "enabledCB") and self.enabledCB:
			self.enabledCB.SetLabel(_t("ENABLE_FILTER"))
		if hasattr(self, "threadNavCB") and self.threadNavCB:
			self.threadNavCB.SetLabel(_t("ENABLE_THREAD_NAV"))
		if hasattr(self, "themeLabel") and self.themeLabel:
			self.themeLabel.SetLabel(_t("SOUND_THEME_LABEL"))
		if hasattr(self, "soundThemeChoice") and self.soundThemeChoice:
			self.soundThemeChoice.SetName(_t("SOUND_THEME_LABEL"))
			themeLabels = [
				_t("THEME_ASSOCIATIVE"),
				_t("THEME_MINIMAL"),
				_t("THEME_INSTRUMENTAL"),
			]
			currSel = self.soundThemeChoice.GetSelection()
			for idx, lbl in enumerate(themeLabels):
				try:
					self.soundThemeChoice.SetString(idx, lbl)
				except Exception:
					pass
			self.soundThemeChoice.SetSelection(currSel)
		if hasattr(self, "soundExplText") and self.soundExplText:
			self.soundExplText.SetLabel(_t("SOUND_EXPL_TEXT"))
		if hasattr(self, "selectAllBtn") and self.selectAllBtn:
			self.selectAllBtn.SetLabel(_t("SELECT_ALL_SOUNDS"))
		if hasattr(self, "deselectAllBtn") and self.deselectAllBtn:
			self.deselectAllBtn.SetLabel(_t("DESELECT_ALL_SOUNDS"))

		if hasattr(self, "soundChildItems"):
			for cb, cbKey, btn, bKey in self.soundChildItems:
				if cb:
					cb.SetLabel(_t(cbKey))
				if btn:
					btn.SetLabel(_t(bKey))

		self.updateParentSoundsLabel()

		if hasattr(self, "volLabel") and self.volLabel:
			self.volLabel.SetLabel(_t("VOL_LABEL"))
		if hasattr(self, "soundVolumeSlider") and self.soundVolumeSlider:
			self.soundVolumeSlider.SetName(_t("SLIDER_NAME"))
		if hasattr(self, "testVolBtn") and self.testVolBtn:
			self.testVolBtn.SetLabel(_t("TEST_VOL_BTN"))

		if hasattr(self, "contextHelpBtn") and self.contextHelpBtn:
			self.contextHelpBtn.SetLabel(_t("CONTEXT_HELP_BTN"))
		if hasattr(self, "helpBtn") and self.helpBtn:
			self.helpBtn.SetLabel(_t("HELP_BTN"))
		if hasattr(self, "contactBtn") and self.contactBtn:
			self.contactBtn.SetLabel(_t("CONTACT_BTN"))
		if hasattr(self, "gesturesBtn") and self.gesturesBtn:
			self.gesturesBtn.SetLabel(_t("GESTURES_BTN"))
		if hasattr(self, "restoreDefaultsBtn") and self.restoreDefaultsBtn:
			self.restoreDefaultsBtn.SetLabel(_t("RESTORE_DEFAULTS_BTN"))
		if hasattr(self, "infoText") and self.infoText:
			self.infoText.SetLabel(
				_t(
					"ADDON_INFO_LINE",
					name=ADDON_NAME,
					ver=ADDON_VERSION,
					build=ADDON_BUILD,
					date=ADDON_BUILD_DATE,
					author=ADDON_AUTHOR,
				)
			)
		try:
			parentDlg = self.GetParent()
			while parentDlg and not isinstance(parentDlg, gui.settingsDialogs.NVDASettingsDialog):
				parentDlg = parentDlg.GetParent()
			if parentDlg and hasattr(parentDlg, "catListCtrl"):
				for i in range(parentDlg.catListCtrl.GetItemCount()):
					text = parentDlg.catListCtrl.GetItemText(i)
					if "Speech Filter" in text or "Filtr mowy" in text:
						parentDlg.catListCtrl.SetItemText(i, _t("TITLE"))
						break
			elif parentDlg and hasattr(parentDlg, "categoriesList"):
				catList = parentDlg.categoriesList
				for i in range(catList.GetCount()):
					data = catList.GetClientData(i) if hasattr(catList, "GetClientData") else None
					text = catList.GetString(i)
					if data is self.__class__ or "Speech Filter" in text or "Filtr mowy" in text:
						catList.SetString(i, _t("TITLE"))
						break
		except Exception:
			pass
		try:
			self.Layout()
			p = self.GetParent()
			if p:
				p.Layout()
		except Exception:
			pass

	def onLanguageChanged(self, evt):
		idx = self.languageChoice.GetSelection()
		if 0 <= idx < len(self.langKeys):
			newKey = self.langKeys[idx]
			config.conf[CONFIG_SECTION]["language"] = newKey
			try:
				config.conf.save()
			except Exception as e:
				log.warning(f"outlookSpeechFilter: Błąd zapisu konfiguracji przy zmianie języka: {e}")
			self.updateUILabels()
			chosenName = self.languageChoice.GetString(idx)
			msg = _t("LANG_CHANGED", lang=chosenName)
			ui.message(msg)

	def onContactAuthor(self, evt):
		dlg = ContactAuthorDialog(self)
		dlg.ShowModal()
		dlg.Destroy()

	def _openContactDialog(self, *args, **kwargs):
		self.onContactAuthor(None)

	def onChildSoundChanged(self, evt):
		selected = sum(1 for cb in self.childSoundCBs if cb.GetValue())
		if selected == 0:
			# Gdy wszystkie dźwięki zostaną odznaczone, automatycznie wyłącz całą opcję i zwiń panel
			self.useSoundsCB.SetValue(False)
			self.showSoundsPanel(False)
			self.updateSoundControlsState()
			self.updateParentSoundsLabel()
			wx.CallAfter(self.useSoundsCB.SetFocus)
			ui.message(_t("SOUNDS_ALL_DESELECTED"))
		else:
			self._lastActiveChildStates = [cb.GetValue() for cb in self.childSoundCBs]
			self.updateSoundControlsState()
			self.updateParentSoundsLabel()
			self.Layout()

	def onSelectAllSounds(self, evt):
		for cb in self.childSoundCBs:
			cb.SetValue(True)
		self._lastActiveChildStates = [True] * len(self.childSoundCBs)
		self.updateSoundControlsState()
		self.updateParentSoundsLabel()
		self.Layout()
		if hasattr(self, "deselectAllBtn") and self.deselectAllBtn and self.deselectAllBtn.IsEnabled():
			wx.CallAfter(self.deselectAllBtn.SetFocus)
		ui.message(_t("SOUNDS_ALL_SELECTED"))

	def onDeselectAllSounds(self, evt):
		# Zapisz stan przed odznaczeniem wszystkich
		currentStates = [cb.GetValue() for cb in self.childSoundCBs]
		if any(currentStates):
			self._lastActiveChildStates = currentStates
		for cb in self.childSoundCBs:
			cb.SetValue(False)
		# Gdy wszystkie dźwięki są odznaczane, następuje wyłączenie całej opcji i zwinięcie panelu
		self.useSoundsCB.SetValue(False)
		self.showSoundsPanel(False)
		self.updateSoundControlsState()
		self.updateParentSoundsLabel()
		wx.CallAfter(self.useSoundsCB.SetFocus)
		ui.message(_t("SOUNDS_ALL_DESELECTED"))

	def _updateSliderValue(self, newVal):
		"""Aktualizuje suwak, etykietę tekstową oraz natychmiast konfigurację z powiadomieniem dla NVDA."""
		newVal = max(0, min(100, newVal))
		self.soundVolumeSlider.SetValue(newVal)
		self.volValueText.SetLabel(f"{newVal}%")
		try:
			config.conf[CONFIG_SECTION]["soundVolume"] = int(newVal)
		except Exception:
			pass
		try:
			ctypes.windll.user32.NotifyWinEvent(
				0x800E,  # EVENT_OBJECT_VALUECHANGE
				self.soundVolumeSlider.GetHandle(),
				-4,      # OBJID_CLIENT
				0,       # CHILDID_SELF
			)
		except Exception:
			pass

	def onSliderKeyDown(self, evt):
		keyCode = evt.GetKeyCode()
		currentVal = self.soundVolumeSlider.GetValue()

		# Page Up = głośniej (+10%)
		if keyCode in (wx.WXK_PAGEUP, getattr(wx, "WXK_NUMPAD_PAGEUP", -1)):
			self._updateSliderValue(currentVal + 10)
			return

		# Page Down = ciszej (-10%)
		elif keyCode in (wx.WXK_PAGEDOWN, getattr(wx, "WXK_NUMPAD_PAGEDOWN", -1)):
			self._updateSliderValue(currentVal - 10)
			return

		# Strzałka w górę lub w prawo = głośniej (+1%)
		elif keyCode in (
			wx.WXK_UP,
			wx.WXK_RIGHT,
			getattr(wx, "WXK_NUMPAD_UP", -1),
			getattr(wx, "WXK_NUMPAD_RIGHT", -1),
		):
			self._updateSliderValue(currentVal + 1)
			return

		# Strzałka w dół lub w lewo = ciszej (-1%)
		elif keyCode in (
			wx.WXK_DOWN,
			wx.WXK_LEFT,
			getattr(wx, "WXK_NUMPAD_DOWN", -1),
			getattr(wx, "WXK_NUMPAD_LEFT", -1),
		):
			self._updateSliderValue(currentVal - 1)
			return

		# Home = maksymalna głośność (100%)
		elif keyCode in (wx.WXK_HOME, getattr(wx, "WXK_NUMPAD_HOME", -1)):
			self._updateSliderValue(100)
			return

		# End = wyciszenie (0%)
		elif keyCode in (wx.WXK_END, getattr(wx, "WXK_NUMPAD_END", -1)):
			self._updateSliderValue(0)
			return

		evt.Skip()

	def onSoundVolumeChanged(self, evt):
		val = self.soundVolumeSlider.GetValue()
		self.volValueText.SetLabel(f"{val}%")
		# Natychmiastowa aktualizacja konfiguracji bez konieczności klikania OK
		try:
			config.conf[CONFIG_SECTION]["soundVolume"] = int(val)
		except Exception:
			pass

	def onTestSoundVolume(self, evt):
		_stop_pending_sounds()
		val = self.soundVolumeSlider.GetValue()
		if val <= 0:
			ui.message(_t("SOUNDS_VOL_MUTED"))
		else:
			# Odtwarza sekwencyjnie wszystkie dźwięki przy wybranej głośności suwaka
			_schedule_status_sounds_sequence(TEST_SOUNDS_SEQUENCE, volume=val, delay_ms=230)

	def onUseSoundsChanged(self, evt):
		show = self.useSoundsCB.GetValue()
		if show:
			# Jeśli wszystkie dźwięki były odznaczone, przywróć stan sprzed wyłączenia
			if not any(cb.GetValue() for cb in self.childSoundCBs):
				if hasattr(self, "_lastActiveChildStates") and any(self._lastActiveChildStates):
					for cb, st in zip(self.childSoundCBs, self._lastActiveChildStates):
						cb.SetValue(_to_bool(st, default=True))
				else:
					for cb in self.childSoundCBs:
						cb.SetValue(True)
				ui.message(_t("SOUNDS_STATE_RESTORED"))
		else:
			currentStates = [cb.GetValue() for cb in self.childSoundCBs]
			if any(currentStates):
				self._lastActiveChildStates = currentStates

		self.showSoundsPanel(show)
		self.updateSoundControlsState()
		self.updateParentSoundsLabel()

	def onOpenHelp(self, evt):
		docFile = get_doc_path()
		if docFile and os.path.isfile(docFile):
			try:
				os.startfile(docFile)
			except Exception as e:
				log.debugWarning(f"Nie udało się otworzyć pomocy: {e}")
		else:
			ui.message(_t("HELP_NOT_FOUND"))

	def onOpenInputGestures(self, evt):
		try:
			wx.CallAfter(gui.mainFrame.onInputGesturesCommand, None)
		except Exception as e:
			log.debugWarning(f"Nie udało się otworzyć okna zdarzeń wejścia: {e}")

	def onRestoreDefaults(self, evt):
		if not showAccessibleConfirmBox(
			_t("RESTORE_CONFIRM_MSG"),
			_t("RESTORE_CONFIRM_TITLE"),
			parent=self,
			defaultYes=False
		):
			return

		# 0. Język
		if hasattr(self, "languageChoice"):
			self.languageChoice.SetSelection(0)
		try:
			config.conf[CONFIG_SECTION]["language"] = "auto"
		except Exception:
			pass

		# 0b. Motyw dźwięków
		if hasattr(self, "soundThemeChoice"):
			self.soundThemeChoice.SetSelection(0)
		try:
			config.conf[CONFIG_SECTION]["soundTheme"] = "associative"
		except Exception:
			pass

		# 1. Główne przełączniki
		self.enabledCB.SetValue(True)
		if hasattr(self, "threadNavCB") and self.threadNavCB:
			self.threadNavCB.SetValue(True)
		self.useSoundsCB.SetValue(True)

		# 2. Suwak głośności i etykieta
		self.soundVolumeSlider.SetValue(80)
		self.volValueText.SetLabel("80%")
		try:
			config.conf[CONFIG_SECTION]["soundVolume"] = 80
		except Exception:
			pass

		# 3. Poszczególne dźwięki według oficjalnych ustawień domyślnych
		self.soundUnreadCB.SetValue(True)
		self.soundFlagsCB.SetValue(False)
		self.soundAttachmentsCB.SetValue(True)
		if hasattr(self, "soundInlineImagesCB") and self.soundInlineImagesCB:
			self.soundInlineImagesCB.SetValue(True)
		self.soundRepliedCB.SetValue(False)
		self.soundForwardedCB.SetValue(False)
		self.soundImportanceHighCB.SetValue(True)
		self.soundImportanceLowCB.SetValue(False)
		self.soundMeetingRequestCB.SetValue(True)
		self.soundCategoriesCB.SetValue(False)
		self.soundExternalCB.SetValue(False)

		# 4. Zapamiętany stan do ewentualnego przywrócenia
		self._lastActiveChildStates = [cb.GetValue() for cb in self.childSoundCBs]

		# 5. Aktualizacja widoczności, stanów przycisków zbiorczych i etykiety
		self.showSoundsPanel(True)
		self.updateSoundControlsState()
		self.updateParentSoundsLabel()
		self.updateUILabels()
		self.Layout()
		try:
			config.conf.save()
		except Exception as e:
			log.warning(f"outlookSpeechFilter: Błąd zapisu konfiguracji w onRestoreDefaults: {e}")

		ui.message(_t("DEFAULTS_RESTORED"))

	def onDiscard(self):
		"""W razie anulowania okna dialogowego przywraca pierwotną głośność i język."""
		global _activeSettingsPanelRef
		_activeSettingsPanelRef = None
		try:
			if hasattr(self, "initialSoundVolume"):
				config.conf[CONFIG_SECTION]["soundVolume"] = _to_int(self.initialSoundVolume, 80)
			if hasattr(self, "initialLanguage"):
				config.conf[CONFIG_SECTION]["language"] = str(self.initialLanguage)
			if hasattr(self, "initialSoundTheme"):
				config.conf[CONFIG_SECTION]["soundTheme"] = str(self.initialSoundTheme)
		except Exception as e:
			log.warning(f"outlookSpeechFilter: Błąd w onDiscard: {e}")

	def onSave(self):
		conf = config.conf[CONFIG_SECTION]
		try:
			if hasattr(self, "languageChoice") and hasattr(self, "langKeys"):
				idx = self.languageChoice.GetSelection()
				if 0 <= idx < len(self.langKeys):
					chosenLang = self.langKeys[idx]
					conf["language"] = chosenLang
					self.initialLanguage = chosenLang
			if hasattr(self, "soundThemeChoice") and hasattr(self, "themeKeys"):
				idx = self.soundThemeChoice.GetSelection()
				if 0 <= idx < len(self.themeKeys):
					chosenTheme = self.themeKeys[idx]
					conf["soundTheme"] = chosenTheme
					self.initialSoundTheme = chosenTheme
			if hasattr(self, "enabledCB") and self.enabledCB:
				conf["enabled"] = bool(self.enabledCB.GetValue())
			if hasattr(self, "threadNavCB") and self.threadNavCB:
				conf["threadNavigationEnabled"] = bool(self.threadNavCB.GetValue())
			if hasattr(self, "useSoundsCB") and self.useSoundsCB:
				conf["useSounds"] = bool(self.useSoundsCB.GetValue())
			if hasattr(self, "soundVolumeSlider") and self.soundVolumeSlider:
				val = int(self.soundVolumeSlider.GetValue())
				conf["soundVolume"] = val
				self.initialSoundVolume = val
			if hasattr(self, "soundUnreadCB") and self.soundUnreadCB:
				conf["soundUnread"] = bool(self.soundUnreadCB.GetValue())
			if hasattr(self, "soundFlagsCB") and self.soundFlagsCB:
				conf["soundFlags"] = bool(self.soundFlagsCB.GetValue())
			if hasattr(self, "soundAttachmentsCB") and self.soundAttachmentsCB:
				conf["soundAttachments"] = bool(self.soundAttachmentsCB.GetValue())
			if hasattr(self, "soundInlineImagesCB") and self.soundInlineImagesCB:
				conf["soundInlineImages"] = bool(self.soundInlineImagesCB.GetValue())
			if hasattr(self, "soundRepliedCB") and self.soundRepliedCB:
				conf["soundReplied"] = bool(self.soundRepliedCB.GetValue())
			if hasattr(self, "soundForwardedCB") and self.soundForwardedCB:
				conf["soundForwarded"] = bool(self.soundForwardedCB.GetValue())
			if hasattr(self, "soundImportanceHighCB") and self.soundImportanceHighCB:
				conf["soundImportanceHigh"] = bool(self.soundImportanceHighCB.GetValue())
			if hasattr(self, "soundImportanceLowCB") and self.soundImportanceLowCB:
				conf["soundImportanceLow"] = bool(self.soundImportanceLowCB.GetValue())
			if hasattr(self, "soundMeetingRequestCB") and self.soundMeetingRequestCB:
				conf["soundMeetingRequest"] = bool(self.soundMeetingRequestCB.GetValue())
			if hasattr(self, "soundCategoriesCB") and self.soundCategoriesCB:
				conf["soundCategories"] = bool(self.soundCategoriesCB.GetValue())
			if hasattr(self, "soundExternalCB") and self.soundExternalCB:
				conf["soundExternal"] = bool(self.soundExternalCB.GetValue())
			if hasattr(self, "chkAutoUpdates") and self.chkAutoUpdates:
				conf["autoCheckUpdates"] = bool(self.chkAutoUpdates.GetValue())
			if hasattr(self, "txtUpdateSource") and self.txtUpdateSource:
				conf["updateSourceUrl"] = self.txtUpdateSource.GetValue().strip()
			# Synchronize legacy keys
			conf["soundForwardReply"] = bool(conf.get("soundReplied", False) or conf.get("soundForwarded", False))
			conf["soundImportance"] = bool(conf.get("soundImportanceHigh", True) or conf.get("soundImportanceLow", False))
			conf["reportUnread"] = bool(conf.get("soundUnread", True))
			conf["reportFlags"] = bool(conf.get("soundFlags", False))
			conf["reportAttachments"] = bool(conf.get("soundAttachments", True))
			conf["reportInlineImages"] = bool(conf.get("soundInlineImages", True))
			conf["reportReplied"] = bool(conf.get("soundReplied", False))
			conf["reportForwarded"] = bool(conf.get("soundForwarded", False))
			conf["reportImportanceHigh"] = bool(conf.get("soundImportanceHigh", True))
			conf["reportImportanceLow"] = bool(conf.get("soundImportanceLow", False))
			conf["reportForwardReply"] = bool(conf.get("soundForwardReply", False))
			conf["reportImportance"] = bool(conf.get("soundImportance", True))
			conf["reportMeetingRequest"] = bool(conf.get("soundMeetingRequest", True))
			conf["reportCategories"] = bool(conf.get("soundCategories", False))
			conf["reportExternal"] = bool(conf.get("soundExternal", False))
			self.updateUILabels()
			try:
				config.conf.save()
			except Exception as e:
				log.warning(f"outlookSpeechFilter: Błąd zapisu konfiguracji do pliku w onSave: {e}")
		except Exception as e:
			log.warning(f"outlookSpeechFilter: Błąd podczas onSave: {e}")


OutlookMessageCleanerSettingsPanel = OutlookSpeechFilterSettingsPanel


def _neutralize_legacy_addon():
	"""
	Bezpieczna i nieinwazyjna obsługa ewentualnej zdublowanej kategorii ustawień.
	Nie kasuje plików w tle ani nie modyfikuje modułów systemowych podczas startu.
	Usunięcie starego dodatku jest delegowane do standardowego mechanizmu addonHandler.
	"""
	try:
		categoryClasses = getattr(gui.settingsDialogs.NVDASettingsDialog, "categoryClasses", None)
		if categoryClasses:
			for catCls in list(categoryClasses):
				if catCls is OutlookSpeechFilterSettingsPanel:
					continue
				clsName = getattr(catCls, "__name__", "")
				modName = getattr(catCls, "__module__", "")
				title = getattr(catCls, "title", "")
				if (
					"OutlookMessageCleaner" in clsName
					or "outlookMessageCleaner" in modName
					or title in ("Outlook - Wyciszanie informacji", "Outlook Message Cleaner")
				):
					try:
						categoryClasses.remove(catCls)
						log.info(f"outlookSpeechFilter: Usunięto zdublowaną kategorię ustawień: {clsName}")
					except Exception:
						pass
	except Exception as e:
		log.debug(f"outlookSpeechFilter: Sprawdzanie kategorii dialogu: {e}")


_original_popupSettingsDialog = None


def _safe_popupSettingsDialog(dialog, *args, **kwargs):
	"""
	Bezpieczny wrapper dla gui.popupSettingsDialog kompatybilny z wieloma wersjami NVDA.
	Zapobiega błędowi RuntimeError: Cannot open new settings dialog while instance still exists.
	"""
	global _original_popupSettingsDialog
	try:
		inst = getattr(dialog, "_instance", None)
		if inst is not None:
			try:
				if isinstance(inst, wx.Window) and inst and inst.IsShown():
					inst.Raise()
					inst.SetFocus()
					return
				else:
					dialog._instance = None
					try:
						inst.Destroy()
					except Exception:
						pass
			except Exception:
				dialog._instance = None
	except Exception:
		pass

	try:
		instances = getattr(dialog, "_instances", None)
		if isinstance(instances, (dict, list, set)):
			to_remove = []
			for it in list(instances):
				w = instances[it] if isinstance(instances, dict) else it
				if not (isinstance(w, wx.Window) and w and w.IsShown()):
					to_remove.append(it)
			for it in to_remove:
				try:
					if isinstance(instances, dict):
						del instances[it]
					else:
						instances.remove(it)
				except Exception:
					pass
	except Exception:
		pass

	if _original_popupSettingsDialog:
		try:
			return _original_popupSettingsDialog(dialog, *args, **kwargs)
		except RuntimeError:
			try:
				dialog._instance = None
				return _original_popupSettingsDialog(dialog, *args, **kwargs)
			except Exception as e:
				log.warning(f"outlookSpeechFilter: Nie udało się otworzyć okna ustawień po resecie _instance: {e}")
				return None
	return None


# ==============================================================================
# NAWIGATOR WĄTKÓW WIADOMOŚCI OUTLOOK CLASSIC (WCAG 2.1 / 2.2 AA)
# ==============================================================================

def _getWordDocumentFromHwnd(hwnd: int):
	"""Pobiera natywny obiekt Document programu Word na podstawie uchwytu okna _WwG."""
	if not hwnd:
		return None
	try:
		import winUser
		if not winUser.isWindow(hwnd):
			return None
		import comtypes.client
		from comtypes.automation import IDispatch
		from ctypes import POINTER, byref, c_void_p, oledll

		pUnknown = POINTER(c_void_p)()
		hr = oledll.oleacc.AccessibleObjectFromWindow(
			hwnd,
			-16,  # OBJID_NATIVEOM
			byref(IDispatch._iid_),
			byref(pUnknown),
		)
		if hr == 0 and pUnknown:
			disp = comtypes.client.wrap(pUnknown)
			return getattr(disp, "Document", None)
	except Exception as e:
		log.debug(f"outlookSpeechFilter: _getWordDocumentFromHwnd error: {e}")
	return None


def _findWordGWindow(parentHwnd: int) -> int:
	"""Przeszukuje okno w poszukiwaniu potomka o klasie _WwG (edytor WordMail)."""
	if not parentHwnd:
		return 0
	found = [0]
	try:
		import winUser
		def enumChild(hwnd, lParam):
			try:
				cls = winUser.getClassName(hwnd)
				if cls == "_WwG" and winUser.isWindowVisible(hwnd):
					found[0] = hwnd
					return False
			except Exception:
				pass
			return True

		WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)
		ctypes.windll.user32.EnumChildWindows(parentHwnd, WNDENUMPROC(enumChild), 0)
	except Exception:
		pass
	return found[0]


def getActiveWordDocument(mailItem=None):
	"""
	Zwraca aktywny obiekt Word.Document powiązany z bieżącą wiadomością Outlooka.
	Obsługuje aktywne okno wiadomości (Inspector), okienko podglądu (Reading Pane),
	a także pobranie WordEditor z obiektu mailItem.GetInspector gdy użytkownik jest na liście wiadomości.
	"""
	focus = api.getFocusObject()
	if not focus:
		return None

	import winUser
	appMod = getattr(focus, "appModule", None)
	isOutlookApp = bool(appMod and getattr(appMod, "appName", "").lower() == "outlook")
	if not isOutlookApp:
		fgHwnd = winUser.getForegroundWindow()
		if fgHwnd and "renwnd32" in (winUser.getClassName(fgHwnd) or "").lower():
			isOutlookApp = True
	if not isOutlookApp:
		return None

	# 1. Bezpośredni obiekt WordDocument na fokusie lub treeInterceptorze
	doc = getattr(focus, "WinwordDocumentObject", None)
	if doc:
		return doc

	ti = getattr(focus, "treeInterceptor", None)
	if ti:
		root = getattr(ti, "rootNVDAObject", None)
		doc = getattr(root, "WinwordDocumentObject", None)
		if doc:
			return doc

	hwnd = getattr(focus, "windowHandle", 0)
	if hwnd and winUser.getClassName(hwnd) == "_WwG":
		doc = _getWordDocumentFromHwnd(hwnd)
		if doc:
			return doc

	if not appMod:
		return None

	nativeOm = getattr(appMod, "nativeOm", None)
	if nativeOm:
		try:
			activeWin = nativeOm.ActiveWindow()
			if activeWin:
				# Jeśli aktywne okno to Inspector (otwarta wiadomość)
				if hasattr(activeWin, "CurrentItem"):
					doc = getattr(activeWin, "WordEditor", None)
					if doc:
						return doc

				# Jeśli aktywne okno to Explorer:
				if hasattr(activeWin, "Selection"):
					fgHwnd = winUser.getForegroundWindow()
					if fgHwnd:
						wHwnd = _findWordGWindow(fgHwnd)
						if wHwnd and winUser.isWindowVisible(wHwnd):
							doc = _getWordDocumentFromHwnd(wHwnd)
							if doc:
								return doc
		except Exception as ex:
			log.debug(f"outlookSpeechFilter: getActiveWordDocument błąd: {ex}")

	# Jeśli nie znaleziono WordEditor w oknach UI, a przekazano mailItem (np. na liście wiadomości)
	if mailItem:
		try:
			insp = getattr(mailItem, "GetInspector", None)
			if insp:
				doc = getattr(insp, "WordEditor", None)
				if doc:
					return doc
		except Exception as ex:
			log.debug(f"outlookSpeechFilter: GetInspector.WordEditor błąd: {ex}")

	return None


def getWordCaretPosition(doc) -> int:
	"""Zwraca pozycję znaku kursora (caret) w aktywnym dokumencie Word."""
	if not doc:
		return 0
	try:
		app = getattr(doc, "Application", None)
		if app:
			sel = getattr(app, "Selection", None)
			if sel:
				return int(getattr(sel.Range, "Start", 0))
	except Exception:
		pass
	return 0


def selectWordRange(doc, char_start: int) -> bool:
	"""Przesuwa kursor i widok w Wordzie dokładnie do wskazanego znaku."""
	if not doc:
		return False
	try:
		target_rng = doc.Range(char_start, char_start)
		target_rng.Select()
		app = getattr(doc, "Application", None)
		if app:
			sel = getattr(app, "Selection", None)
			if sel:
				sel.Collapse(1)  # wdCollapseStart = 1
		try:
			active_win = getattr(doc, "ActiveWindow", None)
			if active_win:
				active_win.ScrollIntoView(target_rng)
		except Exception:
			pass
		return True
	except Exception as ex:
		log.debug(f"outlookSpeechFilter: selectWordRange error: {ex}")
		return False


def resolveMessageRangeInWord(doc, msg) -> int:
	"""
	Wyznacza bezbłędną, rzeczywistą pozycję wiadomości w modelu Worda,
	wykorzystując natywne przeszukiwanie Word.Range.Find.
	Eliminuje to w 100% błędy przesunięcia (drift) spowodowane tabelami HTML,
	obrazkami, podpisami czy kodami pól.
	"""
	if not doc:
		return getattr(msg, "range_start", 0)
	if getattr(msg, "is_top", False):
		return 0
	if hasattr(msg, "real_range_start") and msg.real_range_start is not None:
		return msg.real_range_start

	try:
		date_str = str(getattr(msg, "date", "") or "").strip()
		sender_str = str(getattr(msg, "sender", "") or "").strip()
		raw_h = str(getattr(msg, "raw_header", "") or "").strip()

		candidates = []
		# 1. Unikalna data i godzina wiadomości (najwyższa pewność)
		if date_str and len(date_str) >= 4:
			candidates.append(date_str)
			sub_date = re.sub(r'^[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]+,\s*', '', date_str).strip()
			if sub_date and sub_date != date_str and len(sub_date) >= 4:
				candidates.append(sub_date)

		# 2. Fragment nagłówka z nazwiskiem / From / Od
		if sender_str and len(sender_str) >= 3:
			candidates.append(f"From: {sender_str}")
			candidates.append(f"Od: {sender_str}")
			candidates.append(sender_str)

		# 3. Pierwsze czytelne linie nagłówka
		if raw_h:
			for line in raw_h.splitlines():
				l_s = line.strip()
				if len(l_s) >= 8 and not l_s.startswith("-----") and not l_s.startswith("___"):
					candidates.append(l_s)

		for query in candidates:
			q_clean = query.replace("\r", " ").replace("\n", " ").strip()
			if not q_clean or len(q_clean) < 3:
				continue

			matches = []
			find_rng = doc.Content
			find_rng.Find.ClearFormatting()
			while find_rng.Find.Execute(
				FindText=q_clean[:80],
				MatchCase=False,
				MatchWholeWord=False,
				Forward=True,
				Wrap=0  # wdFindStop
			):
				p = find_rng.Paragraphs(1)
				start_pos = int(p.Range.Start)

				try:
					cur_p = p
					for _ in range(3):
						prev_p = cur_p.Previous()
						if not prev_p:
							break
						prev_text = str(prev_p.Range.Text or "").strip().lower()
						if any(prev_text.startswith(kw) for kw in ("od:", "from:", "von:", "de:", "-----", "_____")):
							start_pos = int(prev_p.Range.Start)
							cur_p = prev_p
						else:
							break
				except Exception:
					pass

				matches.append(start_pos)
				try:
					if find_rng.End >= doc.Content.End:
						break
					find_rng = doc.Range(find_rng.End, doc.Content.End)
				except Exception:
					break

			if matches:
				expected_pos = getattr(msg, "range_start", 0)
				best_pos = min(matches, key=lambda m_pos: abs(m_pos - expected_pos))
				msg.real_range_start = best_pos
				return best_pos
	except Exception as ex:
		log.debug(f"outlookSpeechFilter: resolveMessageRangeInWord błąd: {ex}")

	return getattr(msg, "range_start", 0)


class ThreadMessage:
	"""Reprezentuje pojedynczą wiadomość w wyodrębnionym wątku korespondencji."""
	def __init__(self, index: int, sender: str, date: str, subject: str, to: str,
				 range_start: int, range_end: int, header_end: int, preview: str,
				 raw_header: str, is_top: bool = False):
		self.index = index
		self.sender = sender
		self.date = date
		self.subject = subject
		self.to = to
		self.range_start = range_start
		self.range_end = range_end
		self.header_end = header_end
		self.preview = preview
		self.raw_header = raw_header
		self.is_top = is_top

	@property
	def display_label(self) -> str:
		date_part = f" ({self.date})" if self.date else ""
		if self.preview:
			return f"[{self.index}] {self.sender}{date_part} — {self.preview}"
		return f"[{self.index}] {self.sender}{date_part}"


def _clean_preview(text: str, max_len: int = 140) -> str:
	"""Czyści tekst podglądu z symboli cytowania (>), linii poziomych i pustych wierszy."""
	cleaned = []
	for line in text.splitlines():
		line_clean = re.sub(r'^[>\s\-_=]+', '', line).strip()
		if line_clean and not any(line_clean.lower().startswith(prefix) for prefix in (
			"od:", "from:", "wysłano:", "sent:", "do:", "to:", "temat:", "subject:", "dw:", "cc:",
			"gesendet:", "an:", "betreff:", "de:", "envoyé:", "à:", "objet:"
		)):
			cleaned.append(line_clean)
			if len(" ".join(cleaned)) >= max_len:
				break
	joined = " ".join(cleaned)
	if len(joined) > max_len:
		return joined[:max_len].rstrip() + "..."
	return joined


def _clean_sender(sender_raw: str) -> str:
	"""Wyciąga czytelne imię i nazwisko nadawcy, oczyszczając z ostrych nawiasów i mailto."""
	s = sender_raw.strip()
	m = re.match(r'^["\']?([^"\'<>]+?)["\']?\s*(?:<[^>]+>|\[mailto:[^\]]+\])?$', s)
	if m and m.group(1).strip():
		name = m.group(1).strip()
		if "@" not in name or not re.search(r'<[^>]+>', s):
			return name
	return s


def parseEmailThread(full_text: str, current_sender: str = "", current_date: str = "", current_subject: str = "") -> list:
	"""
	Błyskawicznie analizuje treść maila i wyodrębnia strukturę wątku rozmowy.
	Rozpoznaje nagłówki Outlooka (polskie, angielskie, niemieckie, francuskie)
	oraz jednolinijkowe cytowania (Gmail, aplikacje mobilne, Thunderbird).
	"""
	if not full_text:
		return []

	lines = full_text.splitlines(keepends=True)
	re_block_start = re.compile(
		r'^\s*[-=_]{3,}\s*(?:Wiadomość oryginalna|Original Message|Ursprüngliche Nachricht|Message d\'origine|Wiadomość przesłana dalej|Forwarded Message)?\s*[-=_]{0,}\s*$',
		re.IGNORECASE
	)
	re_field_from = re.compile(
		r'^\s*(?:[-=_]{3,}\s*)?(?:Od|From|Von|De|Nadawca)\s*:\s*(.+)$',
		re.IGNORECASE
	)
	re_field_sent = re.compile(
		r'^\s*(?:Wysłano|Sent|Gesendet|Envoyé|Data|Date)\s*:\s*(.+)$',
		re.IGNORECASE
	)
	re_field_to = re.compile(
		r'^\s*(?:Do|To|An|À|Odbiorca)\s*:\s*(.+)$',
		re.IGNORECASE
	)
	re_field_subject = re.compile(
		r'^\s*(?:Temat|Subject|Betreff|Objet)\s*:\s*(.+)$',
		re.IGNORECASE
	)
	re_inline_header = re.compile(
		r'^\s*(?:>*\s*)*(?:W\s+dniu|Dnia|On\s+[A-Za-z]+|\d{1,2}\s+[a-ząćęłńóśźżA-Z]+\s+\d{4}).+?(?:napisał\(a\)|napisał|napisała|wrote|schrieb|a écrit)\s*:\s*$',
		re.IGNORECASE
	)

	headers = []
	char_offset = 0
	line_offsets = []
	for line in lines:
		line_offsets.append(char_offset)
		char_offset += len(line)

	i = 0
	num_lines = len(lines)
	while i < num_lines:
		line = lines[i]
		line_strip = line.strip()

		# 1. Nagłówek jednolinijkowy (Gmail / smartfony)
		if re_inline_header.match(line_strip):
			raw_h = line_strip
			sender = ""
			date_str = ""
			m_author = re.search(r'(?:o\s+\d{1,2}:\d{2},?|at\s+\d{1,2}:\d{2}\s*(?:AM|PM)?,?)\s+(.+?)\s+(?:napisał|wrote|schrieb|a écrit)', line_strip, re.IGNORECASE)
			if m_author:
				sender = m_author.group(1).strip()
			else:
				m_email = re.search(r'([^\s<>:]+@[^\s<>:]+)', line_strip)
				if m_email:
					sender = m_email.group(1)
				else:
					m_pre = re.search(r'([A-ZĄĆĘŁŃÓŚŹŻ][a-ząćęłńóśźż]+(?:\s+[A-ZĄĆĘŁŃÓŚŹŻ][a-ząćęłńóśźż]+)?)\s+(?:napisał|wrote)', line_strip)
					if m_pre:
						sender = m_pre.group(1)

			m_date = re.search(r'(?:W dniu|Dnia|On)\s+(.+?)(?:o\s+\d|\bat\b|,?\s+[A-ZĄĆĘŁŃÓŚŹŻa-z]+|\s*<)', line_strip, re.IGNORECASE)
			if m_date:
				date_str = m_date.group(1).strip()

			headers.append({
				"start_char": line_offsets[i],
				"end_header_char": line_offsets[i] + len(line),
				"sender": _clean_sender(sender) if sender else _t("THREAD_UNKNOWN_SENDER"),
				"date": date_str,
				"subject": "",
				"to": "",
				"raw_header": raw_h
			})
			i += 1
			continue

		# 2. Nagłówek blokowy (klasyczny Outlook / Exchange)
		is_block_start = bool(re_block_start.match(line_strip))
		m_from = re_field_from.match(line_strip)

		if is_block_start or m_from:
			h_sender = m_from.group(1).strip() if m_from else ""
			h_date = ""
			h_subject = ""
			h_to = ""
			has_other_field = False
			last_h_idx = i

			look_limit = min(num_lines, i + 8)
			curr_j = i + (0 if is_block_start and not m_from else 1)

			while curr_j < look_limit:
				sub_line = lines[curr_j].strip()
				if not sub_line:
					if has_other_field or h_sender:
						last_h_idx = curr_j
						break
					curr_j += 1
					continue

				mf = re_field_from.match(sub_line)
				ms = re_field_sent.match(sub_line)
				mt = re_field_to.match(sub_line)
				msb = re_field_subject.match(sub_line)

				if mf and not h_sender:
					h_sender = mf.group(1).strip()
					last_h_idx = curr_j
				elif ms:
					h_date = ms.group(1).strip()
					has_other_field = True
					last_h_idx = curr_j
				elif mt:
					h_to = mt.group(1).strip()
					has_other_field = True
					last_h_idx = curr_j
				elif msb:
					h_subject = msb.group(1).strip()
					has_other_field = True
					last_h_idx = curr_j
				else:
					break
				curr_j += 1

			if (h_sender and (has_other_field or "@" in h_sender)) or (is_block_start and (h_sender or has_other_field)):
				end_c = line_offsets[last_h_idx] + len(lines[last_h_idx])
				headers.append({
					"start_char": line_offsets[i],
					"end_header_char": end_c,
					"sender": _clean_sender(h_sender) if h_sender else _t("THREAD_UNKNOWN_SENDER"),
					"date": h_date,
					"subject": h_subject,
					"to": h_to,
					"raw_header": "".join(lines[i:last_h_idx+1]).strip()
				})
				i = last_h_idx + 1
				continue

		i += 1

	messages = []
	total_len = len(full_text)

	# 1. Wiadomość najnowsza (na górze dokumentu)
	if headers:
		top_end = headers[0]["start_char"]
		top_body = full_text[:top_end].strip()
		top_preview = _clean_preview(top_body)
		messages.append(ThreadMessage(
			index=1,
			sender=current_sender or _t("THREAD_TOP_MESSAGE"),
			date=current_date or "",
			subject=current_subject,
			to="",
			range_start=0,
			range_end=top_end,
			header_end=0,
			preview=top_preview,
			raw_header="",
			is_top=True
		))

		# 2. Kolejne wiadomości w wątku
		for idx, h in enumerate(headers):
			msg_num = idx + 2
			next_start = headers[idx + 1]["start_char"] if idx + 1 < len(headers) else total_len
			body_text = full_text[h["end_header_char"]:next_start].strip()
			preview = _clean_preview(body_text)

			messages.append(ThreadMessage(
				index=msg_num,
				sender=h["sender"],
				date=h["date"],
				subject=h["subject"],
				to=h["to"],
				range_start=h["start_char"],
				range_end=next_start,
				header_end=h["end_header_char"],
				preview=preview,
				raw_header=h["raw_header"],
				is_top=False
			))
	else:
		# Tylko 1 wiadomość bez cytowanego wątku
		messages.append(ThreadMessage(
			index=1,
			sender=current_sender or _t("THREAD_TOP_MESSAGE"),
			date=current_date,
			subject=current_subject,
			to="",
			range_start=0,
			range_end=total_len,
			header_end=0,
			preview=_clean_preview(full_text),
			raw_header="",
			is_top=True
		))

	return messages


class EmailThreadDialog(wx.Dialog):
	"""
	W pełni dostępne okno listy wiadomości w wątku korespondencji (WCAG 2.1 / 2.2 AA).
	Wzorowane na oknie listy elementów NVDA (NVDA+F7).
	"""
	def __init__(self, parent, title: str, messages: list, current_index: int, on_select_callback, doc=None):
		super().__init__(parent, title=title, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
		self.messages = messages
		self.current_index = current_index
		self.on_select_callback = on_select_callback
		self.doc = doc
		self.filtered_messages = list(messages)

		self.SetSize((760, 560))
		self.SetMinSize((520, 380))
		self.CentreOnScreen()

		panel = wx.Panel(self)
		mainSizer = wx.BoxSizer(wx.VERTICAL)

		# 1. Pole filtru wyszukiwania z czytelną etykietą i podpowiedzią (WCAG AA)
		lblFilter = wx.StaticText(panel, label=_t("THREAD_FILTER_LABEL"))
		mainSizer.Add(lblFilter, 0, wx.ALL, 8)

		self.txtFilter = wx.TextCtrl(panel, style=wx.TE_PROCESS_ENTER)
		self.txtFilter.SetHint(_t("THREAD_FILTER_HINT"))
		self.txtFilter.Bind(wx.EVT_TEXT, self.onFilterChange)
		self.txtFilter.Bind(wx.EVT_TEXT_ENTER, self.onEnter)
		self.txtFilter.Bind(wx.EVT_KEY_DOWN, self.onFilterKeyDown)
		mainSizer.Add(self.txtFilter, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 8)

		# 2. Lista wiadomości (ListBox)
		lblList = wx.StaticText(panel, label=_t("THREAD_LIST_LABEL"))
		mainSizer.Add(lblList, 0, wx.ALL, 8)

		self.listCtrl = wx.ListBox(panel, style=wx.LB_SINGLE | wx.LB_HSCROLL)
		self.listCtrl.Bind(wx.EVT_KEY_DOWN, self.onListKeyDown)
		self.listCtrl.Bind(wx.EVT_LISTBOX, self.onSelectionChange)
		self.listCtrl.Bind(wx.EVT_LISTBOX_DCLICK, self.onOk)
		mainSizer.Add(self.listCtrl, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

		# 3. Pasek przycisków
		btnSizer = wx.BoxSizer(wx.HORIZONTAL)

		self.btnHelp = wx.Button(panel, label=_t("BTN_HELP_CONTEXT"))
		self.btnHelp.Bind(wx.EVT_BUTTON, self.onHelp)
		btnSizer.Add(self.btnHelp, 0, wx.RIGHT, 8)

		self.btnCopy = wx.Button(panel, label=_t("THREAD_BTN_COPY"))
		self.btnCopy.Bind(wx.EVT_BUTTON, self.onCopyBody)
		btnSizer.Add(self.btnCopy, 0, wx.RIGHT, 8)

		btnSizer.AddStretchSpacer()

		self.btnOk = wx.Button(panel, wx.ID_OK, label=_t("THREAD_BTN_GO"))
		self.btnOk.Bind(wx.EVT_BUTTON, self.onOk)
		btnSizer.Add(self.btnOk, 0, wx.RIGHT, 8)

		self.btnCancel = wx.Button(panel, wx.ID_CANCEL, label=_t("BTN_CANCEL"))
		btnSizer.Add(self.btnCancel, 0)

		mainSizer.Add(btnSizer, 0, wx.EXPAND | wx.ALL, 8)
		panel.SetSizer(mainSizer)

		self.populateList()
		self.listCtrl.SetFocus()

		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)

	def populateList(self, query: str = ""):
		self.listCtrl.Clear()
		q = query.strip().lower()
		self.filtered_messages = []

		for m in self.messages:
			if not q or (
				q in m.sender.lower()
				or q in m.date.lower()
				or q in m.subject.lower()
				or q in m.preview.lower()
			):
				self.filtered_messages.append(m)
				self.listCtrl.Append(m.display_label)

		# Zaznacz pozycję odpowiadającą bieżącej
		if self.filtered_messages:
			target_idx = 0
			for idx, fit in enumerate(self.filtered_messages):
				if fit.index == self.current_index:
					target_idx = idx
					break
			self.listCtrl.SetSelection(target_idx)
			wx.CallLater(150, self._speak_current_selection)

	def onFilterChange(self, event):
		self.populateList(self.txtFilter.GetValue())

	def onFilterKeyDown(self, event):
		if event.GetKeyCode() == wx.WXK_DOWN:
			self.listCtrl.SetFocus()
			return
		event.Skip()

	def onSelectionChange(self, event):
		self._speak_current_selection()
		event.Skip()

	def _speak_current_selection(self):
		try:
			sel = self.listCtrl.GetSelection()
			if sel != wx.NOT_FOUND and sel < len(self.filtered_messages):
				msg = self.filtered_messages[sel]
				ui.message(msg.display_label)
		except Exception:
			pass

	def onEnter(self, event):
		self.onOk(event)

	def onListKeyDown(self, event):
		key = event.GetKeyCode()
		if key in (wx.WXK_RETURN, wx.WXK_NUMPAD_ENTER):
			self.onOk(event)
		elif key == wx.WXK_F1:
			self.onHelp(event)
		elif key in (ord('K'), ord('k')) and event.AltDown():
			self.onCopyBody(event)
		else:
			event.Skip()

	def onCharHook(self, event):
		key = event.GetKeyCode()
		if key == wx.WXK_ESCAPE:
			self.EndModal(wx.ID_CANCEL)
		elif key == wx.WXK_F1:
			self.onHelp(None)
		elif key in (ord('K'), ord('k')) and event.AltDown():
			self.onCopyBody(None)
		else:
			event.Skip()

	def onCopyBody(self, event):
		sel = self.listCtrl.GetSelection()
		if sel == wx.NOT_FOUND or sel >= len(self.filtered_messages):
			return
		target_m = self.filtered_messages[sel]
		copied_text = ""
		if self.doc:
			try:
				start_c = target_m.header_end if target_m.header_end > 0 else target_m.range_start
				end_c = target_m.range_end
				rng = self.doc.Range(start_c, end_c)
				copied_text = str(getattr(rng, "Text", "") or "").strip()
			except Exception as ex:
				log.debug(f"EmailThreadDialog: błąd pobierania tekstu z Range: {ex}")

		if not copied_text:
			copied_text = target_m.preview

		if copied_text:
			try:
				api.copyToClip(copied_text)
				ui.message(_t("THREAD_COPIED_TO_CLIPBOARD", sender=target_m.sender))
			except Exception as ex:
				log.debug(f"EmailThreadDialog: copyToClip error: {ex}")

	def onHelp(self, event):
		showContextHelp("thread")

	def onOk(self, event):
		sel = self.listCtrl.GetSelection()
		if sel != wx.NOT_FOUND and sel < len(self.filtered_messages):
			target_msg = self.filtered_messages[sel]
			if self.on_select_callback:
				self.on_select_callback(target_msg)
		self.EndModal(wx.ID_OK)


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	scriptCategory = SettingsPanelTitleDescriptor()

	__gestures = {
		# Przełącznik wyciszania mowy (NVDA+Ctrl+Shift+O oraz Insert+Ctrl+Shift+O we wszystkich kolejnościach klawiszy)
		"kb:NVDA+control+shift+o": "toggleOutlookCleaner",
		"kb:control+shift+NVDA+o": "toggleOutlookCleaner",
		"kb:insert+control+shift+o": "toggleOutlookCleaner",
		"kb:control+shift+insert+o": "toggleOutlookCleaner",
		"kb:NVDA+shift+control+o": "toggleOutlookCleaner",
		"kb:shift+control+NVDA+o": "toggleOutlookCleaner",
		"kb:insert+shift+control+o": "toggleOutlookCleaner",
		"kb:shift+control+insert+o": "toggleOutlookCleaner",

		# Przełącznik dźwięków (NVDA+Ctrl+Shift+S oraz Insert+Ctrl+Shift+S)
		"kb:NVDA+control+shift+s": "toggleSounds",
		"kb:control+shift+NVDA+s": "toggleSounds",
		"kb:insert+control+shift+s": "toggleSounds",
		"kb:control+shift+insert+s": "toggleSounds",
		"kb:NVDA+shift+control+s": "toggleSounds",
		"kb:shift+control+NVDA+s": "toggleSounds",
		"kb:insert+shift+control+s": "toggleSounds",
		"kb:shift+control+insert+s": "toggleSounds",

		# Odczyt wszystkich statusów na żądanie (NVDA+Ctrl+Shift+Z oraz Insert+Ctrl+Shift+Z)
		"kb:NVDA+control+shift+z": "reportAllMessageStatuses",
		"kb:control+shift+NVDA+z": "reportAllMessageStatuses",
		"kb:insert+control+shift+z": "reportAllMessageStatuses",
		"kb:control+shift+insert+z": "reportAllMessageStatuses",
		"kb:NVDA+shift+control+z": "reportAllMessageStatuses",
		"kb:shift+control+NVDA+z": "reportAllMessageStatuses",
		"kb:insert+shift+control+z": "reportAllMessageStatuses",
		"kb:shift+control+insert+z": "reportAllMessageStatuses",

		# Dodatkowy, prosty i wygodny skrót NVDA+Shift+Z / Insert+Shift+Z bez konieczności wciskania 4 klawiszy naraz
		"kb:NVDA+shift+z": "reportAllMessageStatuses",
		"kb:shift+NVDA+z": "reportAllMessageStatuses",
		"kb:insert+shift+z": "reportAllMessageStatuses",
		"kb:shift+insert+z": "reportAllMessageStatuses",

		# Pomoc kontekstowa i spis skrótów klawiszowych (NVDA+Ctrl+Shift+H oraz Insert+Ctrl+Shift+H)
		"kb:NVDA+control+shift+h": "showContextHelp",
		"kb:control+shift+NVDA+h": "showContextHelp",
		"kb:insert+control+shift+h": "showContextHelp",
		"kb:control+shift+insert+h": "showContextHelp",
		"kb:NVDA+shift+control+h": "showContextHelp",
		"kb:shift+control+NVDA+h": "showContextHelp",
		"kb:insert+shift+control+h": "showContextHelp",
		"kb:shift+control+insert+h": "showContextHelp",

		# Przełącznik nawigatora wątków (NVDA+Ctrl+Shift+T oraz Insert+Ctrl+Shift+T)
		"kb:NVDA+control+shift+t": "toggleThreadNavigation",
		"kb:control+shift+NVDA+t": "toggleThreadNavigation",
		"kb:insert+control+shift+t": "toggleThreadNavigation",
		"kb:control+shift+insert+t": "toggleThreadNavigation",
		"kb:NVDA+shift+control+t": "toggleThreadNavigation",
		"kb:shift+control+NVDA+t": "toggleThreadNavigation",
		"kb:insert+shift+control+t": "toggleThreadNavigation",
		"kb:shift+control+insert+t": "toggleThreadNavigation",

		# Nawigator wątków wiadomości (NVDA+F7, NVDA+Alt+F7, Insert+F7, Insert+Alt+F7)
		"kb:NVDA+f7": "showEmailThreadStructure",
		"kb:insert+f7": "showEmailThreadStructure",
		"kb:NVDA+alt+f7": "showEmailThreadStructure",
		"kb:alt+NVDA+f7": "showEmailThreadStructure",
		"kb:insert+alt+f7": "showEmailThreadStructure",
		"kb:alt+insert+f7": "showEmailThreadStructure",

		# Szybkie skoki po wiadomościach w wątku (NVDA+Alt+Dół/Góra/Home/End)
		"kb:NVDA+alt+downArrow": "nextThreadMessage",
		"kb:alt+NVDA+downArrow": "nextThreadMessage",
		"kb:insert+alt+downArrow": "nextThreadMessage",
		"kb:alt+insert+downArrow": "nextThreadMessage",

		"kb:NVDA+alt+upArrow": "prevThreadMessage",
		"kb:alt+NVDA+upArrow": "prevThreadMessage",
		"kb:insert+alt+upArrow": "prevThreadMessage",
		"kb:alt+insert+upArrow": "prevThreadMessage",

		"kb:NVDA+alt+home": "firstThreadMessage",
		"kb:alt+NVDA+home": "firstThreadMessage",
		"kb:insert+alt+home": "firstThreadMessage",
		"kb:alt+insert+home": "firstThreadMessage",

		"kb:NVDA+alt+end": "lastThreadMessage",
		"kb:alt+NVDA+end": "lastThreadMessage",
		"kb:insert+alt+end": "lastThreadMessage",
		"kb:alt+insert+end": "lastThreadMessage",
	}

	def getScript(self, gesture):
		script = super().getScript(gesture)
		if not script:
			return None
		script_name = getattr(script, "__name__", "")
		outlook_scripts = {
			"script_showEmailThreadStructure",
			"script_nextThreadMessage",
			"script_prevThreadMessage",
			"script_firstThreadMessage",
			"script_lastThreadMessage",
			"script_reportAllMessageStatuses",
		}
		if script_name in outlook_scripts:
			# Jeśli Outlook nie jest aktywny, zwracamy None, aby NVDA
			# przekazało gest do właściwego programu lub innego dodatku (np. listy elementów F7 w przeglądarce).
			if not self._is_outlook_active():
				return None
			# Jeśli to nawigacja po wątku:
			if script_name == "script_showEmailThreadStructure":
				conf = config.conf.get(CONFIG_SECTION, {})
				if not _to_bool(conf.get("threadNavigationEnabled", True), default=True):
					return None
				g_name = gesture.displayName.lower() if hasattr(gesture, "displayName") else ""
				if "alt" not in g_name and not self._is_outlook_message_active():
					return None
		return script

	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)

		# Zabezpieczenie popupSettingsDialog przed zablokowaniem skrótu NVDA+Ctrl+G
		global _original_popupSettingsDialog
		if _original_popupSettingsDialog is None and hasattr(gui, "popupSettingsDialog"):
			_original_popupSettingsDialog = gui.popupSettingsDialog
			gui.popupSettingsDialog = _safe_popupSettingsDialog

		# Natychmiastowe odblokowanie ewentualnej zawieszonej instancji NVDASettingsDialog
		try:
			curInst = getattr(gui.settingsDialogs.NVDASettingsDialog, "_instance", None)
			if curInst is not None:
				try:
					if not (isinstance(curInst, wx.Window) and curInst and curInst.IsShown()):
						gui.settingsDialogs.NVDASettingsDialog._instance = None
				except Exception:
					gui.settingsDialogs.NVDASettingsDialog._instance = None
		except Exception:
			pass

		config.conf.spec[CONFIG_SECTION] = confspec
		if CONFIG_SECTION not in config.conf:
			config.conf[CONFIG_SECTION] = {}
		targetSec = config.conf[CONFIG_SECTION]

		# Bezpieczna migracja ustawień ze starej konfiguracji (jeśli istnieje) z jawną konwersją typów
		if OLD_CONFIG_SECTION in config.conf:
			try:
				oldSec = config.conf[OLD_CONFIG_SECTION]
				for k, v in oldSec.items():
					if k in confspec and k not in targetSec:
						targetSec[k] = v
				log.info("outlookSpeechFilter: Przeprowadzono bezpieczną migrację konfiguracji ze starej sekcji outlookMessageCleaner")
			except Exception as e:
				log.warning(f"outlookSpeechFilter: Błąd migracji ze starej konfiguracji: {e}")

		# Ścisła sanitizacja i konwersja typów wszystkich kluczy w sekcji konfiguracji
		# Zapobiega błędom TypeError: CheckBox.SetValue(): argument 1 has unexpected type 'str'
		for k, specStr in confspec.items():
			if k in targetSec:
				rawVal = targetSec[k]
				if specStr.startswith("boolean"):
					defVal = "default=True" in specStr
					targetSec[k] = _to_bool(rawVal, default=defVal)
				elif specStr.startswith("integer"):
					targetSec[k] = _to_int(rawVal, default=80, min_val=0, max_val=100)
				elif specStr.startswith("string"):
					targetSec[k] = str(rawVal)

		# Walidacja i zapis poprawionej konfiguracji
		try:
			config.conf.validate(config.conf.validator, copy=True)
		except Exception:
			pass
		try:
			config.conf.save()
		except Exception:
			pass

		patchOutlookClasses()

		# Bezpieczna rejestracja panelu ustawień bez dublowania
		if OutlookSpeechFilterSettingsPanel not in gui.settingsDialogs.NVDASettingsDialog.categoryClasses:
			gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(OutlookSpeechFilterSettingsPanel)

		# Jawne powiązanie gestów klawiszowych dla natychmiastowego działania z klawiszem Insert i CapsLock
		for gestureStr, scriptName in self.__gestures.items():
			try:
				self.bindGesture(gestureStr, scriptName)
			except Exception:
				pass

		# Inicjalizacja modułu aktualizacji i opóźnione sprawdzenie w tle
		self._updater = AddonUpdaterEngine(
			addon_name=CONFIG_SECTION,
			addon_title=ADDON_NAME,
			current_version=ADDON_VERSION,
			addon_dir=ADDON_DIR,
			config_section=CONFIG_SECTION
		)
		wx.CallLater(15000, self._check_updates_auto)

	def _check_updates_auto(self):
		conf = config.conf.get(CONFIG_SECTION, {})
		if not conf.get("autoCheckUpdates", True):
			return
		last_time = conf.get("lastUpdateCheckTime", 0.0)
		if time.time() - last_time < 86400:
			return
		self._updater.check_for_updates(manual=False)

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		appMod = getattr(obj, "appModule", None)
		if appMod and getattr(appMod, "appName", "").lower() == "outlook":
			patchOutlookClasses()
			cachedCls = getattr(getattr(obj, "UIAElement", None), "cachedClassName", "") or ""
			rawName = getattr(getattr(obj, "UIAElement", None), "cachedName", "") or getattr(getattr(obj, "UIAElement", None), "currentName", "") or ""
			role = getattr(obj, "role", None)

			isRow = cachedCls in ("LeafRow", "ThreadItem", "ThreadHeader", "GroupHeader", "OutlookGridGroupHeader", "TableItem")
			isGroup = (
				role in (controlTypes.Role.GROUPING, controlTypes.Role.HEADER)
				or (rawName and any(k in rawName.lower() for k in ("grupuj według", "group by", "kategorie:", "kategoria:")))
			)
			if isRow or isGroup:
				if CleanOutlookGridRow not in clsList:
					clsList.insert(0, CleanOutlookGridRow)

	def terminate(self, *args, **kwargs):
		super().terminate(*args, **kwargs)
		_stop_pending_sounds()
		global _original_UIAGridRow_get_name, _original_getReceivedMessageString, _patches_applied, _original_popupSettingsDialog

		# Przywrócenie dokładnych oryginalnych metod dla każdej klasy
		for (target_obj, attr_name), orig in list(_saved_class_methods.items()):
			try:
				curr = getattr(target_obj, attr_name, None)
				if curr in (_custom_UIAGridRow_get_name, _custom_getReceivedMessageString):
					setattr(target_obj, attr_name, orig)
			except Exception:
				pass
		_saved_class_methods.clear()

		_original_UIAGridRow_get_name = None
		_original_getReceivedMessageString = None
		_patches_applied = False

		# Przywrócenie oryginalnego popupSettingsDialog
		if _original_popupSettingsDialog is not None:
			try:
				gui.popupSettingsDialog = _original_popupSettingsDialog
			except Exception:
				pass
			_original_popupSettingsDialog = None

		# Czyste usunięcie wszystkich instancji panelu z categoryClasses
		while OutlookSpeechFilterSettingsPanel in gui.settingsDialogs.NVDASettingsDialog.categoryClasses:
			try:
				gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(OutlookSpeechFilterSettingsPanel)
			except (ValueError, KeyError):
				break

	@script(
		description=_("Przełącza filtrowanie mowy o wiadomościach w skrzynce programu Outlook"),
		gestures=[
			"kb:NVDA+control+shift+o",
			"kb:control+shift+NVDA+o",
			"kb:insert+control+shift+o",
			"kb:control+shift+insert+o",
		],
	)
	def script_toggleOutlookCleaner(self, gesture):
		conf = config.conf[CONFIG_SECTION]
		newState = not _to_bool(conf.get("enabled", True), default=True)
		conf["enabled"] = newState
		try:
			config.conf.save()
		except Exception:
			pass
		panel = get_active_settings_panel()
		if panel:
			def _update_ui():
				p = get_active_settings_panel()
				if p:
					p.setCleanerEnabledState(newState)
			wx.CallAfter(_update_ui)
		if newState:
			ui.message(_t("TOGGLE_FILTER_ON"))
		else:
			ui.message(_t("TOGGLE_FILTER_OFF"))

	# Alias dla nowej nazwy skryptu
	script_toggleOutlookSpeechFilter = script_toggleOutlookCleaner

	@script(
		description=_("Przełącza odtwarzanie dźwięków statusów w skrzynce programu Outlook"),
		gestures=[
			"kb:NVDA+control+shift+s",
			"kb:control+shift+NVDA+s",
			"kb:insert+control+shift+s",
			"kb:control+shift+insert+s",
		],
	)
	def script_toggleSounds(self, gesture):
		conf = config.conf[CONFIG_SECTION]
		newState = not _to_bool(conf.get("useSounds", True), default=True)
		conf["useSounds"] = newState
		try:
			config.conf.save()
		except Exception:
			pass
		panel = get_active_settings_panel()
		if panel:
			def _update_ui():
				p = get_active_settings_panel()
				if p:
					p.setUseSoundsState(newState)
			wx.CallAfter(_update_ui)
		if newState:
			ui.message(_t("TOGGLE_SOUNDS_ON"))
		else:
			ui.message(_t("TOGGLE_SOUNDS_OFF"))

	@script(
		description=_("Odczytuje na żądanie wszystkie statusy aktualnej wiadomości w Outlooku"),
		gestures=[
			"kb:NVDA+control+shift+z",
			"kb:control+shift+NVDA+z",
			"kb:insert+control+shift+z",
			"kb:control+shift+insert+z",
			"kb:NVDA+shift+z",
			"kb:shift+NVDA+z",
			"kb:insert+shift+z",
			"kb:shift+insert+z",
		],
	)
	def script_reportAllMessageStatuses(self, gesture):
		focus = api.getFocusObject()
		if not focus:
			ui.message(_t("STATUS_NO_ACTIVE_ITEM"))
			return

		appMod = getattr(focus, "appModule", None)
		if not appMod or getattr(appMod, "appName", "").lower() != "outlook":
			ui.message(_t("STATUS_ONLY_OUTLOOK"))
			return

		# Sprawdź, czy użytkownik ma zaznaczonych wiele wiadomości na liście w Explorerze
		selObj = None
		if getattr(appMod, "nativeOm", None):
			try:
				selObj = appMod.nativeOm.activeExplorer().selection
			except Exception:
				pass
		if not selObj and getattr(focus, "parent", None):
			p = focus.parent
			if getattr(p, "appModule", None) and getattr(p.appModule, "nativeOm", None):
				try:
					selObj = p.appModule.nativeOm.activeExplorer().selection
				except Exception:
					pass

		selCount = int(getattr(selObj, "Count", 0)) if selObj else 0
		if selCount > 1:
			unreadCount = 0
			attCount = 0
			flaggedCount = 0
			highCount = 0
			limit = min(selCount, 100)
			for i in range(1, limit + 1):
				try:
					item = selObj.Item(i)
					if getattr(item, "UnRead", False):
						unreadCount += 1
					if int(getattr(getattr(item, "Attachments", None), "Count", 0)) > 0:
						attCount += 1
					if bool(getattr(item, "IsMarkedAsTask", False)) or int(getattr(item, "FlagStatus", 0)) == 2:
						flaggedCount += 1
					if int(getattr(item, "Importance", 1)) == 2:
						highCount += 1
				except Exception:
					pass

			lang = get_effective_language()
			isPl = (lang == "pl")
			details = []
			if isPl:
				if unreadCount == selCount:
					details.append("wszystkie nieprzeczytane")
				elif unreadCount > 0:
					details.append(f"{unreadCount} nieprzeczytanych")
				else:
					details.append("wszystkie przeczytane")

				if attCount > 0:
					details.append(f"{attCount} z załącznikami")
				if flaggedCount > 0:
					details.append(f"{flaggedCount} oflagowanych")
				if highCount > 0:
					details.append(f"{highCount} o wysokim priorytecie")

				msg = f"Zaznaczono {selCount} wiadomości: " + ", ".join(details)
			else:
				if unreadCount == selCount:
					details.append("all unread")
				elif unreadCount > 0:
					details.append(f"{unreadCount} unread")
				else:
					details.append("all read")

				if attCount > 0:
					details.append(f"{attCount} with attachments")
				if flaggedCount > 0:
					details.append(f"{flaggedCount} flagged")
				if highCount > 0:
					details.append(f"{highCount} high importance")

				msg = f"Selected {selCount} messages: " + ", ".join(details)

			ui.message(msg)
			return

		# 1. Sprawdź zarejestrowane statusy na wierszu podczas ostatniego przetwarzania
		rowStatuses = getattr(focus, "_osf_statuses", None) or getattr(focus, "_omc_statuses", None)
		if not rowStatuses and getattr(focus, "parent", None):
			rowStatuses = getattr(focus.parent, "_osf_statuses", None) or getattr(focus.parent, "_omc_statuses", None)
		if not rowStatuses and _last_focused_statuses:
			rowStatuses = _last_focused_statuses

		# 2. Pobierz obiekt wiadomości (selection)
		selection = None
		if getattr(appMod, "nativeOm", None):
			try:
				selection = appMod.nativeOm.activeExplorer().selection.item(1)
			except (COMError, Exception):
				selection = None

		if not selection and getattr(focus, "parent", None):
			p = focus.parent
			if getattr(p, "appModule", None) and getattr(p.appModule, "nativeOm", None):
				try:
					selection = p.appModule.nativeOm.activeExplorer().selection.item(1)
				except (COMError, Exception):
					selection = None

		statuses = _get_message_statuses_report(
			appMod,
			selection,
			windowThreadID=getattr(focus, "windowThreadID", None),
			uiaObj=focus,
		)

		# 3. Jeśli rowStatuses zawiera informacje uzupełniające, połącz je (odfiltrowując fałszywe kategorie)
		if rowStatuses:
			for s in rowStatuses:
				sLower = s.strip().lower()
				if sLower in ("bez kategorii", "brak kategorii", "no category", "no categories", "none"):
					continue
				if s not in statuses:
					statuses.append(s)

		# 4. Posortuj ściśle według CANONICAL_STATUS_ORDER
		def _status_order_key(item):
			itemLower = item.lower()
			if "nieprzeczytana" in itemLower or "unread" in itemLower:
				return 0
			if "oflagow" in itemLower or "flaga" in itemLower or "flag" in itemLower:
				return 1
			if "załącznik" in itemLower or "attachment" in itemLower:
				return 2
			if "odpowiedz" in itemLower or "replied" in itemLower:
				return 3
			if "przesłan" in itemLower or "forward" in itemLower:
				return 4
			if "wysoki priorytet" in itemLower or "high" in itemLower:
				return 5
			if "niski priorytet" in itemLower or "low" in itemLower:
				return 6
			if any(k in itemLower for k in ("priorytet", "ważność", "importance")):
				return 6
			if any(k in itemLower for k in ("spotkani", "meeting")):
				return 7
			if any(k in itemLower for k in ("kategori", "category")):
				return 8
			if any(k in itemLower for k in ("zewnętrzn", "external")):
				return 9
			return 10

		statuses.sort(key=_status_order_key)

		if statuses:
			ui.message(", ".join(statuses))
		elif selection or rowStatuses is not None:
			ui.message(_t("STATUS_NO_ADDITIONAL"))
		else:
			ui.message(_t("STATUS_NO_SELECTION"))

	@script(
		description=_("Wyświetla pomoc kontekstową i spis skrótów klawiszowych dodatku Outlook Speech Filter"),
		gestures=[
			"kb:NVDA+control+shift+h",
			"kb:control+shift+NVDA+h",
			"kb:insert+control+shift+h",
			"kb:control+shift+insert+h",
		],
	)
	def script_showContextHelp(self, gesture):
		showContextHelp("outlook")

	script_contextHelp = script_showContextHelp

	# --------------------------------------------------------------------------
	# NAWIGATOR WĄTKÓW WIADOMOŚCI OUTLOOK (SKRYPTY I LOGIKA)
	# --------------------------------------------------------------------------

	def _is_outlook_active(self) -> bool:
		"""Sprawdza, czy aktywną aplikacją na pierwszym planie jest Microsoft Outlook."""
		focus = api.getFocusObject()
		if not focus:
			return False
		appMod = getattr(focus, "appModule", None)
		if appMod and getattr(appMod, "appName", "").lower() == "outlook":
			return True
		try:
			import winUser
			fg = winUser.getForegroundWindow()
			if fg:
				cls = (winUser.getClassName(fg) or "").lower()
				if "renwnd32" in cls or "outlook" in cls:
					return True
		except Exception:
			pass
		return False

	def _is_outlook_message_active(self) -> bool:
		"""Sprawdza, czy fokus znajduje się bezpośrednio w oknie lub podglądzie wiadomości."""
		focus = api.getFocusObject()
		if not focus or not self._is_outlook_active():
			return False
		try:
			import winUser
			wCls = (getattr(focus, "windowClassName", "") or "").lower()
			if "_wwg" in wCls:
				return True
			fg = winUser.getForegroundWindow()
			if fg:
				cls = (winUser.getClassName(fg) or "").lower()
				if "rctrl_renwnd32" in cls:
					return True
		except Exception:
			pass
		return False

	def _get_thread_messages_and_cur_idx(self) -> tuple[list, int, any, any]:
		"""
		Pobiera strukturę wątku, aktualny indeks wiadomości, obiekt doc i mailItem.
		Zwraca: (messages, cur_idx, doc, mailItem)
		"""
		doc = getActiveWordDocument()
		cur_idx = 0
		mailItem = None

		focus = api.getFocusObject()
		appMod = getattr(focus, "appModule", None) if focus else None

		if appMod and getattr(appMod, "appName", "").lower() == "outlook":
			nativeOm = getattr(appMod, "nativeOm", None)
			if nativeOm:
				try:
					actWin = nativeOm.ActiveWindow()
					if actWin and hasattr(actWin, "CurrentItem"):
						mailItem = actWin.CurrentItem
					elif actWin and hasattr(actWin, "Selection") and actWin.Selection.Count > 0:
						mailItem = actWin.Selection.Item(1)
				except Exception:
					pass

		if not doc and mailItem:
			try:
				insp = getattr(mailItem, "GetInspector", None)
				if insp:
					doc = getattr(insp, "WordEditor", None)
			except Exception:
				pass

		if not doc:
			return [], 0, None, mailItem

		try:
			full_text = str(getattr(doc.Content, "Text", "") or "")
		except Exception as ex:
			log.debug(f"outlookSpeechFilter: błąd odczytu doc.Content.Text: {ex}")
			return [], 0, None, mailItem

		sender_name = ""
		date_str = ""
		subject_str = ""
		if mailItem:
			try:
				sender_name = str(getattr(mailItem, "SenderName", "") or "")
			except Exception:
				pass
			try:
				sent_on = getattr(mailItem, "SentOn", None)
				if sent_on:
					date_str = str(sent_on)
			except Exception:
				pass
			try:
				subject_str = str(getattr(mailItem, "Subject", "") or "")
			except Exception:
				pass

		msgs = parseEmailThread(
			full_text,
			current_sender=sender_name,
			current_date=date_str,
			current_subject=subject_str
		)

		if not msgs:
			return [], 0, doc, mailItem

		cur_pos = getWordCaretPosition(doc)
		for idx, m in enumerate(msgs):
			if m.range_start <= cur_pos < m.range_end:
				cur_idx = idx
				break
			elif cur_pos >= m.range_end:
				cur_idx = idx

		return msgs, cur_idx, doc, mailItem

	def _jump_to_thread_message(self, target_msg, msgs: list, doc=None, mailItem=None):
		"""Przechodzi dokładnie do wskazanej wiadomości w wątku, grając dźwięk i anonsując szczegóły."""
		if mailItem and not doc:
			try:
				mailItem.Display()
				time.sleep(0.15)
				doc = getActiveWordDocument()
			except Exception:
				pass

		if not doc:
			doc = getActiveWordDocument()

		if doc:
			exact_pos = resolveMessageRangeInWord(doc, target_msg)
			selectWordRange(doc, exact_pos)

		# Dźwięk potwierdzenia przejścia
		_play_status_sound("replied")

		# Ogłoszenie mowy
		idx = target_msg.index
		total = len(msgs)
		sender = target_msg.sender or _t("THREAD_UNKNOWN_SENDER")
		date = target_msg.date or ""
		date_info = f", {date}" if date else ""
		msg_text = _t("THREAD_MSG_ANNOUNCE", index=idx, total=total, sender=sender, date=date_info)
		ui.message(msg_text)

	@script(
		description=_("Przełącza nawigator wątków wiadomości w programie Outlook"),
		gestures=[
			"kb:NVDA+control+shift+t",
			"kb:control+shift+NVDA+t",
			"kb:insert+control+shift+t",
			"kb:control+shift+insert+t",
		],
	)
	def script_toggleThreadNavigation(self, gesture):
		conf = config.conf[CONFIG_SECTION]
		newState = not _to_bool(conf.get("threadNavigationEnabled", True), default=True)
		conf["threadNavigationEnabled"] = newState
		try:
			config.conf.save()
		except Exception:
			pass
		panel = get_active_settings_panel()
		if panel:
			def _update_ui():
				p = get_active_settings_panel()
				if p:
					p.setThreadNavEnabledState(newState)
			wx.CallAfter(_update_ui)
		if newState:
			ui.message(_t("TOGGLE_THREAD_NAV_ON"))
		else:
			ui.message(_t("TOGGLE_THREAD_NAV_OFF"))

	@script(
		description=_("Wyświetla listę wiadomości w wątku bieżącej korespondencji w Outlooku (okno struktury wątku)"),
		gestures=[
			"kb:NVDA+f7",
			"kb:insert+f7",
			"kb:NVDA+alt+f7",
			"kb:insert+alt+f7",
		]
	)
	def script_showEmailThreadStructure(self, gesture):
		conf = config.conf.get(CONFIG_SECTION, {})
		if not _to_bool(conf.get("threadNavigationEnabled", True), default=True):
			if "f7" in gesture.displayName.lower() and "alt" not in gesture.displayName.lower():
				gesture.send()
			else:
				ui.message(_t("THREAD_NAV_DISABLED"))
			return

		if not self._is_outlook_active():
			gesture.send()
			return

		msgs, cur_idx, doc, mailItem = self._get_thread_messages_and_cur_idx()
		if not msgs:
			if "alt" not in gesture.displayName.lower() and not self._is_outlook_message_active():
				gesture.send()
			else:
				ui.message(_t("THREAD_NO_MESSAGES_OR_DOC"))
			return

		if len(msgs) <= 1:
			ui.message(_t("THREAD_ONLY_ONE_MESSAGE"))
			return

		def _show_dialog():
			def _on_select(target_m):
				self._jump_to_thread_message(target_m, msgs, doc, mailItem)

			dlg = EmailThreadDialog(
				parent=gui.mainFrame,
				title=_t("THREAD_DIALOG_TITLE"),
				messages=msgs,
				current_index=msgs[cur_idx].index if msgs else 1,
				on_select_callback=_on_select,
				doc=doc
			)
			gui.mainFrame.prePopup()
			try:
				dlg.ShowModal()
			finally:
				dlg.Destroy()
				gui.mainFrame.postPopup()

		wx.CallAfter(_show_dialog)

	@script(
		description=_("Przechodzi do następnej (starszej) wiadomości w wątku korespondencji w Outlooku"),
		gestures=[
			"kb:NVDA+alt+downArrow",
			"kb:insert+alt+downArrow",
		]
	)
	def script_nextThreadMessage(self, gesture):
		conf = config.conf.get(CONFIG_SECTION, {})
		if not _to_bool(conf.get("threadNavigationEnabled", True), default=True):
			ui.message(_t("THREAD_NAV_DISABLED"))
			return

		if not self._is_outlook_active():
			gesture.send()
			return

		msgs, cur_idx, doc, mailItem = self._get_thread_messages_and_cur_idx()
		if not msgs or len(msgs) <= 1:
			ui.message(_t("THREAD_ONLY_ONE_MESSAGE"))
			return

		target_idx = cur_idx + 1
		if target_idx >= len(msgs):
			ui.message(_t("THREAD_LAST_MESSAGE"))
			try:
				import tones
				tones.beep(220, 60)
			except Exception:
				pass
			return

		self._jump_to_thread_message(msgs[target_idx], msgs, doc, mailItem)

	@script(
		description=_("Przechodzi do poprzedniej (nowszej) wiadomości w wątku korespondencji w Outlooku"),
		gestures=[
			"kb:NVDA+alt+upArrow",
			"kb:insert+alt+upArrow",
		]
	)
	def script_prevThreadMessage(self, gesture):
		conf = config.conf.get(CONFIG_SECTION, {})
		if not _to_bool(conf.get("threadNavigationEnabled", True), default=True):
			ui.message(_t("THREAD_NAV_DISABLED"))
			return

		if not self._is_outlook_active():
			gesture.send()
			return

		msgs, cur_idx, doc, mailItem = self._get_thread_messages_and_cur_idx()
		if not msgs or len(msgs) <= 1:
			ui.message(_t("THREAD_ONLY_ONE_MESSAGE"))
			return

		target_idx = cur_idx - 1
		if target_idx < 0:
			ui.message(_t("THREAD_FIRST_MESSAGE"))
			try:
				import tones
				tones.beep(220, 60)
			except Exception:
				pass
			return

		self._jump_to_thread_message(msgs[target_idx], msgs, doc, mailItem)

	@script(
		description=_("Przechodzi do najnowszej (pierwszej) wiadomości w wątku korespondencji w Outlooku"),
		gestures=[
			"kb:NVDA+alt+home",
			"kb:insert+alt+home",
		]
	)
	def script_firstThreadMessage(self, gesture):
		conf = config.conf.get(CONFIG_SECTION, {})
		if not _to_bool(conf.get("threadNavigationEnabled", True), default=True):
			ui.message(_t("THREAD_NAV_DISABLED"))
			return

		if not self._is_outlook_active():
			gesture.send()
			return
		msgs, cur_idx, doc, mailItem = self._get_thread_messages_and_cur_idx()
		if not msgs or len(msgs) <= 1:
			ui.message(_t("THREAD_ONLY_ONE_MESSAGE"))
			return
		self._jump_to_thread_message(msgs[0], msgs, doc, mailItem)

	@script(
		description=_("Przechodzi do najstarszej (ostatniej) wiadomości w wątku korespondencji w Outlooku"),
		gestures=[
			"kb:NVDA+alt+end",
			"kb:insert+alt+end",
		]
	)
	def script_lastThreadMessage(self, gesture):
		conf = config.conf.get(CONFIG_SECTION, {})
		if not _to_bool(conf.get("threadNavigationEnabled", True), default=True):
			ui.message(_t("THREAD_NAV_DISABLED"))
			return

		if not self._is_outlook_active():
			gesture.send()
			return
		msgs, cur_idx, doc, mailItem = self._get_thread_messages_and_cur_idx()
		if not msgs or len(msgs) <= 1:
			ui.message(_t("THREAD_ONLY_ONE_MESSAGE"))
			return
		self._jump_to_thread_message(msgs[-1], msgs, doc, mailItem)

