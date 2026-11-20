# -*- coding: utf-8 -*-
# NVDA Add-on: Outlook Speech Filter
# Author: Sławomir Strugarek z pomocą Antigravity (Google DeepMind)
# Description: Controls and mutes unwanted message statuses (attachments, read/unread, replied/forwarded, etc.)
# in Microsoft Outlook's message list.

from __future__ import annotations
import array
import collections
import ctypes
from ctypes import wintypes
import hashlib
import json
import locale
import os
import platform
import re
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import wave
import weakref
from comtypes import COMError
import comtypes.automation

import addonHandler
import api
import config
try:
	from config.configFlags import ReportTableHeaders
except (ImportError, AttributeError):
	ReportTableHeaders = None
import controlTypes
import core
import globalPluginHandler
import globalVars
import gui
import gui.guiHelper
import gui.settingsDialogs
from gui.settingsDialogs import SettingsPanel
from logHandler import log
import nvwave
from scriptHandler import script
import tones
import ui
import UIAHandler
import wx

try:
	import nvdaBuiltin.appModules.outlook as builtinOutlook
except ImportError:
	builtinOutlook = None

try:
	from NVDAObjects.UIA import UIA as _UIAObjectClass
except ImportError:
	_UIAObjectClass = None

try:
	from UIAHandler.utils import createUIAMultiPropertyCondition
except ImportError:
	createUIAMultiPropertyCondition = getattr(builtinOutlook, "createUIAMultiPropertyCondition", None)

# initTranslation() udostępnia gettext jako _ (bez katalogów .po dodatek korzysta z własnego mechanizmu _t)
addonHandler.initTranslation()

CONFIG_SECTION = "outlookSpeechFilter"
OLD_CONFIG_SECTION = "outlookMessageCleaner"
ADDON_NAME = "Outlook Speech Filter"
ADDON_VERSION = "1.21.0"
ADDON_BUILD = "20261120.1"
ADDON_BUILD_DATE = "2026-11-20"
ADDON_AUTHOR = "Sławomir Strugarek (z pomocą Antigravity)"
ADDON_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SOUNDS_DIR = os.path.join(ADDON_DIR, "sounds")
SOUNDS_CACHE_DIR = os.path.join(tempfile.gettempdir(), "nvda_osf_sounds")
DOC_PATH = os.path.join(ADDON_DIR, "doc.html")
DOC_PATH_PL = os.path.join(ADDON_DIR, "doc_pl.html")
DOC_PATH_EN = os.path.join(ADDON_DIR, "doc_en.html")
GITHUB_REPO_URL = "https://github.com/strugo10/outlookSpeechFilter"
GITHUB_ISSUES_URL = "https://github.com/strugo10/outlookSpeechFilter/issues"
GITHUB_RELEASES_URL = "https://github.com/strugo10/outlookSpeechFilter/releases"
GITHUB_LATEST_DOWNLOAD = "https://github.com/strugo10/outlookSpeechFilter/releases/latest/download/outlookSpeechFilter.nvda-addon"
DEFAULT_UPDATE_SOURCE_URL = "https://raw.githubusercontent.com/strugo10/outlookSpeechFilter/main/update_info.json"
SECONDARY_UPDATE_SOURCE_URL = "https://raw.githubusercontent.com/strugo10/outlookSpeechFilter/main/updates.json"
# Aktualizacje wyłącznie przez HTTPS z serwerów GitHub (również po przekierowaniach).
# release-assets.githubusercontent.com to obecny host, na który GitHub przekierowuje pobieranie zasobów wydań.
UPDATE_ALLOWED_HOSTS = (
	"github.com",
	"raw.githubusercontent.com",
	"objects.githubusercontent.com",
	"release-assets.githubusercontent.com",
)
UPDATE_MAX_PACKAGE_SIZE = 50 * 1024 * 1024

confspec = {
	"enabled": "boolean(default=True)",
	"threadNavigationEnabled": "boolean(default=True)",
	"language": "string(default='auto')",
	"soundTheme": "string(default='associative')",
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
	Wynik (poza podglądem wyboru w otwartym panelu ustawień) jest buforowany do zmiany konfiguracji.
	"""
	global _effective_language_cache
	# Podgląd języka wybranego w otwartym (jeszcze niezapisanym) panelu ustawień — tylko w wątku głównym wx
	try:
		if _activeSettingsPanelRef and wx.IsMainThread():
			panel = _activeSettingsPanelRef()
			if panel and hasattr(panel, "languageChoice") and hasattr(panel, "langKeys"):
				idx = panel.languageChoice.GetSelection()
				if 0 <= idx < len(panel.langKeys):
					chosen = panel.langKeys[idx]
					if chosen in ("pl", "en"):
						return chosen
					if chosen == "auto":
						return _detect_system_language()
	except Exception:
		pass
	cached = _effective_language_cache
	if cached:
		return cached
	try:
		confLang = config.conf.get(CONFIG_SECTION, {}).get("language", "auto")
	except Exception:
		confLang = "auto"
	if confLang in ("pl", "en"):
		result = confLang
	else:
		result = _detect_system_language()
	_effective_language_cache = result
	return result


def _detect_system_language():
	"""Wykrywa język NVDA (a w razie potrzeby systemu Windows) bez przestarzałego locale.getdefaultlocale()."""
	try:
		import languageHandler
		nvdaLang = (languageHandler.getLanguage() or "").lower()
		if nvdaLang.startswith("pl"):
			return "pl"
		if nvdaLang and nvdaLang != "windows" and "default" not in nvdaLang:
			return "en"
		getWinLang = getattr(languageHandler, "getWindowsLanguage", None)
		if getWinLang:
			if (getWinLang() or "").lower().startswith("pl"):
				return "pl"
			return "en"
	except Exception:
		pass
	try:
		loc = (locale.getlocale()[0] or "").lower()
		if loc.startswith("pl") or loc.startswith("polish"):
			return "pl"
	except Exception:
		pass
	return "en"


_effective_language_cache = None
_activeSettingsPanelRef = None


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
				"- Open GitHub Issues in browser:\n"
				"  Opens the official GitHub Issues page for bug reports and feature requests.\n"
				"- Copy report to clipboard:\n"
				"  Copies a complete diagnostic report to your clipboard containing NVDA version,\n"
				"  Windows version, add-on settings, and Outlook integration state.\n"
				"  You can paste it (Ctrl+V) directly into your GitHub issue.\n"
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
		elif context == "sound_customization":
			title = "Help: Customize Event Sounds"
			text = (
				"OUTLOOK SPEECH FILTER — SOUND CUSTOMIZATION HELP\n"
				+ "=" * 60 + "\n\n"
				"OPTIONS IN THIS DIALOG:\n"
				"- Status sounds theme:\n"
				"  Choose between Associative (real-world icons), Minimalist (pure tones), and Instrumental (acoustic instruments).\n"
				"- Sound volume slider & Test volume button:\n"
				"  Adjust playback volume (0-100%) and preview sounds immediately.\n"
				"- Select all / Deselect all buttons:\n"
				"  Quickly toggle all status sounds at once.\n"
				"- Status sounds list:\n"
				"  Enable or disable individual sounds and use the Preview button next to each sound.\n\n"
				"KEYBOARD SHORTCUTS:\n"
				"  F1              - Show this contextual help\n"
				"  Escape          - Cancel and close without saving changes\n"
				"  Enter (on OK)   - Apply sound settings and close\n"
			)
		else:  # settings
			title = "Help: Outlook Speech Filter Settings"
			text = (
				"OUTLOOK SPEECH FILTER — SETTINGS HELP\n"
				+ "=" * 60 + "\n\n"
				"AVAILABLE OPTIONS & SETTINGS:\n"
				"- Mute unwanted message status announcements:\n"
				"  Mutes verbose prefixes on the message list, allowing NVDA to read\n"
				"  sender, subject, and date/time first.\n"
				"- Enable conversation thread navigator (NVDA+F7):\n"
				"  Provides structured conversation tree navigation and quick email reading.\n"
				"- Play subtle earcons instead of muted speech announcements:\n"
				"  Master switch for audio cues (supports 3 states: checked, unchecked, partially checked).\n"
				"- Customize Event Sounds button (Alt+D):\n"
				"  Opens dedicated window with sound theme choice, volume slider, volume preview,\n"
				"  and independent selection of all 11 status sounds.\n"
				"- Automatic update checks:\n"
				"  Option to check GitHub for updates upon NVDA startup and on-demand check button.\n"
				"- Add-on Interface Language (at bottom of dialog):\n"
				"  Switch between Automatic, Polish, and English.\n\n"
				"UTILITY & HELP BUTTONS:\n"
				"- Contextual Help (F1): Displays this help buffer.\n"
				"- Open Help: Opens full HTML documentation in default web browser.\n"
				"- Configure shortcut gestures: Opens NVDA's Input Gestures dialog.\n"
				"- Share add-on: Share add-on link with others.\n"
				"- Contact add-on author: Opens GitHub issues report dialog.\n"
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
			title = "Pomoc: Kontakt z autorem dodatku (GitHub)"
			text = (
				"OUTLOOK FILTR MOWY — POMOC OKNA KONTAKTU\n"
				+ "=" * 60 + "\n\n"
				"OPCJE DOSTĘPNE W OKNIE KONTAKTU:\n"
				"- Otwórz zgłoszenia na GitHubie (GitHub Issues):\n"
				"  Otwiera oficjalną stronę zgłoszeń projektu na GitHubie, gdzie możesz opisać\n"
				"  problem, zadać pytanie lub zaproponować nową funkcję.\n"
				"- Kopiuj raport do schowka:\n"
				"  Kopiuje do schowka kompletny raport diagnostyczny dodatku zawierający wersję NVDA,\n"
				"  wersję systemu Windows, konfigurację dodatku i stan integracji z Outlookiem.\n"
				"  Możesz wkleić ten raport (Ctrl+V) w treści nowego zgłoszenia na GitHubie.\n"
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
		elif context == "sound_customization":
			title = "Pomoc: Dostosuj sygnały dźwiękowe zdarzeń"
			text = (
				"OUTLOOK FILTR MOWY — DOSTOSOWANIE SYGNAŁÓW DŹWIĘKOWYCH\n"
				+ "=" * 60 + "\n\n"
				"OPCJE W TYM OKNIE:\n"
				"- Motyw dźwięków statusów:\n"
				"  Wybór pomiędzy motywem Skojarzeniowym (dźwięki przedmiotów: spinacz, aparat, dzwonek),\n"
				"  Minimalistycznym (elektroniczne tony pojedyncze i podwójne) oraz Instrumentalnym (marimba, wibrafon).\n"
				"- Suwak głośności dźwięków oraz przycisk 'Przetestuj głośność':\n"
				"  Płynna regulacja głośności (0-100%) z możliwością natychmiastowego odsłuchania.\n"
				"- Przyciski zbiorcze: 'Zaznacz wszystkie dźwięki' oraz 'Odznacz wszystkie dźwięki'.\n"
				"- Lista 11 sygnałów dźwiękowych:\n"
				"  Indywidualne włączanie/wyłączanie poszczególnych dźwięków wraz z przyciskiem 'Odsłuchaj'.\n\n"
				"SKRÓTY KLAWISZOWE:\n"
				"  F1              - Wyświetl tę pomoc kontekstową\n"
				"  Escape          - Anuluj i zamknij bez zapisywania zmian\n"
				"  Enter (na OK)   - Zastosuj i zamknij okno dźwięków\n"
			)
		else:  # settings
			title = "Pomoc: Ustawienia dodatku Outlook Speech Filter"
			text = (
				"OUTLOOK FILTR MOWY — POMOC USTAWIEŃ DODATKU\n"
				+ "=" * 60 + "\n\n"
				"GŁÓWNE OPCJE I PRZEŁĄCZNIKI:\n"
				"- Włącz filtrowanie mowy o wiadomościach:\n"
				"  Wycisza uciążliwe prefiksy mowy na liście maili w programie Outlook,\n"
				"  pozwalając NVDA czytać od razu nadawcę, temat oraz datę/godzinę.\n"
				"- Włącz nawigator wątków konwersacji (NVDA+F7):\n"
				"  Umożliwia przeglądanie struktury korespondencji w czytelnym drzewie wiadomości.\n"
				"- Odtwarzaj dyskretne dźwięki zamiast wyciszonych komunikatów mowy:\n"
				"  Główny 3-stanowy przełącznik sygnałów dźwiękowych (earconów).\n"
				"- Przycisk 'Dostosuj sygnały dźwiękowe...' (Alt+D):\n"
				"  Otwiera dedykowane okno konfiguracji dźwięków, w którym możesz wybrać motyw,\n"
				"  ustawić i przetestować głośność oraz włączyć lub wyłączyć poszczególne sygnały.\n"
				"- Automatyczne sprawdzanie aktualizacji:\n"
				"  Opcja sprawdzania wydań na GitHubie oraz przycisk sprawdzenia na żądanie.\n"
				"- Język interfejsu dodatku (na dole okna):\n"
				"  Wybór języka (Automatyczny, Polski, English).\n\n"
				"PRZYCISKI NARZĘDZIOWE I POMOC:\n"
				"- Pomoc kontekstowa (F1): Wyświetla niniejsze okno pomocy.\n"
				"- Otwórz pomoc: Otwiera pełną dokumentację HTML w przeglądarce internetowej.\n"
				"- Skonfiguruj skróty klawiszowe: Przechodzi do konfiguracji gestów w Zdarzeniach wejścia NVDA.\n"
				"- Podziel się dodatkiem ze znajomymi: Generuje linki do pobrania z oficjalnego profilu GitHub.\n"
				"- Kontakt z autorem: Otwiera oficjalne zgłoszenia GitHub Issues i kreator raportu.\n"
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
		"pl": "Otwórz pomoc...",
		"en": "Open Help...",
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
		"pl": "Kontakt z autorem (GitHub)...",
		"en": "Contact Author (GitHub)...",
	},
	"CUSTOMIZE_SOUNDS_BTN": {
		"pl": "&Dostosuj sygnały dźwiękowe...",
		"en": "&Customize Event Sounds...",
	},
	"CUSTOMIZE_SOUNDS_DLG_TITLE": {
		"pl": "Dostosuj sygnały dźwiękowe zdarzeń",
		"en": "Customize Event Sounds",
	},
	"CUSTOMIZE_SOUNDS_INTRO": {
		"pl": "Zaznacz dźwięki, które mają być odtwarzane w tle zamiast czytanych komunikatów mowy. Przy każdym dźwięku możesz nacisnąć przycisk Odsłuchaj, aby usłyszeć jego brzmienie:",
		"en": "Select the sounds you want to play in the background instead of speech announcements. You can press the Preview button next to each sound to hear how it sounds:",
	},
	"SHARE_BTN": {
		"pl": "&Podziel się dodatkiem ze znajomymi...",
		"en": "&Share Add-on with Others...",
	},
	"SHARE_DLG_TITLE": {
		"pl": "Podziel się dodatkiem ze znajomymi",
		"en": "Share Add-on with Others",
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
		"pl": "Kontakt z autorem i zgłoszenia na GitHubie: Sławomir Strugarek",
		"en": "Contact Author & GitHub Issues: Sławomir Strugarek",
	},
	"CONTACT_DLG_INFO": {
		"pl": "Zgłoszenia problemów, pomysły na nowe funkcje oraz kontakt z autorem realizowane są bezpośrednio w oficjalnym serwisie GitHub (GitHub Issues).\n\nMożesz skopiować poniższy raport diagnostyczny jednym kliknięciem i wkleić go w nowym zgłoszeniu na GitHubie.",
		"en": "Bug reports, feature requests, and feedback are handled via official GitHub Issues.\n\nYou can copy the diagnostic report below with one click and paste it into a new issue on GitHub.",
	},
	"CONTACT_TOPIC_LABEL": {
		"pl": "Kategoria zgłoszenia:",
		"en": "Report category:",
	},
	"CONTACT_BTN_OPEN_FORM": {
		"pl": "Otwórz zgłoszenia na GitHubie (GitHub Issues)...",
		"en": "Open GitHub Issues in web browser...",
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
	"GROUP_ITEMS_COUNT": {
		"pl": "{count} elementów",
		"en": "{count} items",
	},
	"MSG_INFO_TITLE": {
		"pl": "Informacja",
		"en": "Information",
	},
	"MSG_CONFIRM_TITLE": {
		"pl": "Potwierdzenie",
		"en": "Confirmation",
	},
	"BTN_CLOSE_ACCEL": {
		"pl": "&Zamknij",
		"en": "&Close",
	},
	"SETTINGS_LOAD_ERROR": {
		"pl": "Wystąpił błąd podczas ładowania panelu ustawień Outlook Speech Filter: {err}",
		"en": "An error occurred while loading the Outlook Speech Filter settings panel: {err}",
	},
	"AUTO_UPDATES_CB": {
		"pl": "&Automatycznie sprawdzaj dostępność aktualizacji przy uruchomieniu NVDA",
		"en": "&Automatically check for updates when NVDA starts",
	},
	"CHECK_UPDATES_BTN": {
		"pl": "&Sprawdź dostępność aktualizacji teraz",
		"en": "&Check for updates now",
	},
	"CONTACT_REPORT_COPY_FAILED": {
		"pl": "Nie udało się skopiować raportu do schowka.",
		"en": "Could not copy the report to the clipboard.",
	},
	"REPORT_PRIVACY_WARNING": {
		"pl": "Uwaga: raport może zawierać dane osobowe (np. fragmenty dziennika NVDA, nazwy okien lub plików). Adresy e-mail i nazwa użytkownika w ścieżkach są maskowane automatycznie, ale przed publicznym udostępnieniem raportu (np. na GitHubie) przejrzyj jego treść.",
		"en": "Warning: the report may contain personal data (e.g. NVDA log excerpts, window or file names). E-mail addresses and the user name in paths are masked automatically, but please review the report before sharing it publicly (e.g. on GitHub).",
	},
	"THREAD_FILTER_RESULTS": {
		"pl": "Wyników: {count}",
		"en": "Results: {count}",
	},
	"UPD_DLG_TITLE": {
		"pl": "Aktualizacja: {title}",
		"en": "Update: {title}",
	},
	"UPD_HEADER": {
		"pl": "Dostępna jest nowa wersja dodatku: {title}!",
		"en": "A new version of the add-on is available: {title}!",
	},
	"UPD_VERSIONS": {
		"pl": "Zainstalowana wersja: {local}   --->   Nowa wersja: {remote}",
		"en": "Installed version: {local}   --->   New version: {remote}",
	},
	"UPD_CHANGES_LABEL": {
		"pl": "&Wykaz nowości i ulepszeń w tej wersji:",
		"en": "&What's new in this version:",
	},
	"UPD_NO_CHANGELOG": {
		"pl": "Brak dodatkowego opisu zmian.",
		"en": "No description of changes provided.",
	},
	"UPD_PROMPT": {
		"pl": "Czy chcesz teraz pobrać i zainstalować tę aktualizację?",
		"en": "Do you want to download and install this update now?",
	},
	"UPD_BTN_INSTALL": {
		"pl": "&Zaktualizuj teraz",
		"en": "&Update now",
	},
	"UPD_BTN_LATER": {
		"pl": "&Przypomnij później",
		"en": "&Remind me later",
	},
	"UPD_TITLE": {
		"pl": "Aktualizacje: {title}",
		"en": "Updates: {title}",
	},
	"UPD_CHECK_FAILED": {
		"pl": "Nie udało się sprawdzić aktualizacji dla dodatku {title}.\n\nSzczegóły: {err}",
		"en": "Could not check for updates of {title}.\n\nDetails: {err}",
	},
	"UPD_UP_TO_DATE": {
		"pl": "Posiadasz najnowszą wersję dodatku {title} ({ver}).\nBrak nowszych aktualizacji w serwisie GitHub.",
		"en": "You have the latest version of {title} ({ver}).\nNo newer updates are available on GitHub.",
	},
	"UPD_UNEXPECTED_ERROR": {
		"pl": "Wystąpił nieoczekiwany błąd podczas sprawdzania aktualizacji:\n{err}",
		"en": "An unexpected error occurred while checking for updates:\n{err}",
	},
	"UPD_ERROR_TITLE": {
		"pl": "Błąd: {title}",
		"en": "Error: {title}",
	},
	"UPD_DOWNLOADING": {
		"pl": "Pobieranie aktualizacji dla {title} z serwisu GitHub...",
		"en": "Downloading update for {title} from GitHub...",
	},
	"UPD_ERR_NETWORK": {
		"pl": "Nie udało się połączyć z serwerem aktualizacji GitHub (błąd sieci): {err}",
		"en": "Could not connect to the GitHub update server (network error): {err}",
	},
	"UPD_ERR_NO_ENTRY": {
		"pl": "Brak wpisu dla dodatku '{name}' w pliku informacji o aktualizacji.",
		"en": "No entry for add-on '{name}' in the update information file.",
	},
	"UPD_ERR_JSON": {
		"pl": "Błąd przetwarzania formatu JSON: {err}",
		"en": "JSON parsing error: {err}",
	},
	"UPD_ERR_URL": {
		"pl": "Niedozwolony adres aktualizacji (dozwolone jest wyłącznie HTTPS z serwerów GitHub): {url}",
		"en": "Update address not allowed (only HTTPS from GitHub servers is permitted): {url}",
	},
	"UPD_ERR_NO_SHA": {
		"pl": "Informacja o aktualizacji nie zawiera sumy kontrolnej SHA-256. Instalacja została zablokowana ze względów bezpieczeństwa.",
		"en": "The update information does not contain a SHA-256 checksum. Installation was blocked for security reasons.",
	},
	"UPD_ERR_SHA_MISMATCH": {
		"pl": "Suma kontrolna SHA-256 pobranego pakietu nie zgadza się z opublikowaną. Pakiet został odrzucony.",
		"en": "The SHA-256 checksum of the downloaded package does not match the published one. The package was rejected.",
	},
	"UPD_ERR_TOO_BIG": {
		"pl": "Pobrany pakiet przekracza dopuszczalny rozmiar.",
		"en": "The downloaded package exceeds the allowed size.",
	},
	"UPD_ERR_MANIFEST": {
		"pl": "Pobrany pakiet nie zawiera prawidłowego pliku manifest.ini.",
		"en": "The downloaded package does not contain a valid manifest.ini file.",
	},
	"UPD_ERR_NAME": {
		"pl": "Niezgodna nazwa dodatku w pakiecie: oczekiwano '{expected}', otrzymano '{got}'.",
		"en": "Add-on name mismatch in package: expected '{expected}', got '{got}'.",
	},
	"UPD_ERR_NOT_NEWER": {
		"pl": "Wersja w pobranym pakiecie ({ver}) nie jest nowsza od zainstalowanej ({cur}).",
		"en": "The version in the downloaded package ({ver}) is not newer than the installed one ({cur}).",
	},
	"UPD_INCOMPATIBLE": {
		"pl": "Pobrana wersja dodatku nie jest kompatybilna z bieżącą wersją NVDA.",
		"en": "The downloaded add-on version is not compatible with the current NVDA version.",
	},
	"UPD_OVERRIDE_CONFIRM": {
		"pl": "Pobrana wersja dodatku {title} ({ver}) nie jest oznaczona jako zgodna z bieżącą wersją NVDA.\n\nCzy mimo to chcesz ją zainstalować? Dodatek może działać nieprawidłowo.",
		"en": "The downloaded version of {title} ({ver}) is not marked as compatible with the current NVDA version.\n\nDo you want to install it anyway? The add-on may not work correctly.",
	},
	"UPD_OVERRIDE_TITLE": {
		"pl": "Niezgodna wersja dodatku",
		"en": "Incompatible add-on version",
	},
	"UPD_OVERRIDE_DECLINED": {
		"pl": "Instalacja niezgodnej wersji została anulowana.",
		"en": "Installation of the incompatible version was cancelled.",
	},
	"UPD_SUCCESS": {
		"pl": "Pomyślnie zaktualizowano dodatek {title} do wersji {ver}.\n\nCzy chcesz teraz zrestartować NVDA, aby zastosować zmiany?",
		"en": "{title} was successfully updated to version {ver}.\n\nDo you want to restart NVDA now to apply the changes?",
	},
	"UPD_SUCCESS_TITLE": {
		"pl": "Aktualizacja zakończona sukcesem",
		"en": "Update completed successfully",
	},
	"UPD_INSTALL_FAILED": {
		"pl": "Nie udało się zainstalować aktualizacji dodatku {title}.\n\nSzczegóły błędu:\n{err}",
		"en": "Could not install the update of {title}.\n\nError details:\n{err}",
	},
	"UPD_INSTALL_FAILED_TITLE": {
		"pl": "Błąd instalacji: {title}",
		"en": "Installation error: {title}",
	},
	"UPD_UNKNOWN_INSTALL_ERROR": {
		"pl": "Nieznany błąd instalacji pakietu.",
		"en": "Unknown package installation error.",
	},
	"UPD_INSTALL_RETURNED_NONE": {
		"pl": "Funkcja instalacji dodatku nie zwróciła obiektu dodatku.",
		"en": "The add-on installation function did not return an add-on object.",
	},
	"UPD_DOWNLOAD_FAILED": {
		"pl": "Nie udało się pobrać lub zweryfikować aktualizacji:\n{err}",
		"en": "Could not download or verify the update:\n{err}",
	},
	"UPD_DOWNLOAD_FAILED_TITLE": {
		"pl": "Błąd aktualizacji: {title}",
		"en": "Update error: {title}",
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
	"Otwórz pomoc...": "HELP_BTN",
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


class AccessibleMessageDialog(wx.Dialog):
	"""
	W pełni dostępne okno komunikatów i instrukcji dla czytnika ekranu NVDA.
	Fokus jest automatycznie ustawiany w wielowierszowym polu tekstowym tylko do odczytu (wx.TextCtrl).
	Dzięki temu użytkownik programu NVDA od razu słyszy treść komunikatu,
	może swobodnie czytać go linijka po linijce strzałkami w górę i w dół,
	skopiować tekst (Ctrl+C), a także zamknąć okno klawiszem Enter lub Escape.
	"""

	def __init__(self, parent, message: str, title: str = None, buttonLabel: str = None):
		title = title or _t("MSG_INFO_TITLE")
		buttonLabel = buttonLabel or _t("BTN_CLOSE_ACCEL")
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

	def __init__(self, parent, message: str, title: str = None, defaultYes: bool = False):
		title = title or _t("MSG_CONFIRM_TITLE")
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


def showAccessibleMessageBox(message: str, title: str = None, parent=None, buttonLabel: str = None):
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


def showAccessibleConfirmBox(message: str, title: str = None, parent=None, defaultYes: bool = False) -> bool:
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
			filtered = [
				l.strip() for l in all_lines
				if any(k in l.lower() for k in ("outlook", "speechfilter", "messagecleaner", "error", "exception", "traceback"))
				and not _is_private_log_line(l)
			]
			for fl in filtered[-50:]:
				lines.append(f"  {fl}")
		else:
			lines.append("  Brak dostępu do pliku logu NVDA")
	except Exception as e:
		lines.append(f"  Błąd odczytu logu NVDA: {e}")
	lines.append("=" * 60)
	lines.append("KONIEC RAPORTU / END OF REPORT")
	lines.append("=" * 60)
	return _sanitize_report_text("\n".join(lines))


# Linie dziennika NVDA, które mogą zawierać wypowiadany lub wyświetlany tekst (treść maili, nazwiska itp.)
_PRIVATE_LOG_MARKERS = (
	"speaking",
	"braille",
	"speech.speech",
	"speechsequence",
	"speech sequence",
	"speakwithoutpauses",
	"input: kb",
	"inputcore.inputmanager.executegesture",
)
_RE_EMAIL = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
_RE_USER_PATH = re.compile(r"(?i)([A-Z]:[\\/]+(?:Users|Documents and Settings)[\\/]+)[^\\/\r\n'\"<>|:*?]+")


def _is_private_log_line(line: str) -> bool:
	low = (line or "").lower()
	return any(m in low for m in _PRIVATE_LOG_MARKERS)


def _sanitize_report_text(text: str) -> str:
	"""Maskuje adresy e-mail oraz nazwę użytkownika w ścieżkach profilu (C:\\Users\\<nazwa>)."""
	try:
		text = _RE_EMAIL.sub("<e-mail>", text)
		text = _RE_USER_PATH.sub(lambda m: m.group(1) + "<user>", text)
		try:
			userName = os.environ.get("USERNAME") or ""
			if len(userName) >= 3:
				text = re.sub(re.escape(userName), "<user>", text, flags=re.IGNORECASE)
		except Exception:
			pass
	except Exception:
		pass
	return text


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
	Przekierowuje bezpośrednio do zgłoszeń w serwisie GitHub (GitHub Issues).
	"""
	def __init__(self, parent, topicKey="", topicTitle=""):
		self.topicKey = topicKey
		self.topicTitle = topicTitle
		dlgTitle = _t("CONTACT_DLG_TITLE")
		super().__init__(parent, title=dlgTitle, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, size=(580, 390))
		self.SetMinSize((540, 350))
		self.reportContent = None

		mainSizer = wx.BoxSizer(wx.VERTICAL)

		# 1. Informacja o zgłoszeniach na GitHubie
		infoText = wx.StaticText(self, label=_t("CONTACT_DLG_INFO"))
		infoText.Wrap(540)
		mainSizer.Add(infoText, flag=wx.ALL, border=10)

		# 2. Główny przycisk: Otwórz GitHub Issues w przeglądarce
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
		url = GITHUB_ISSUES_URL
		try:
			os.startfile(url)
		except Exception:
			try:
				import webbrowser
				webbrowser.open(url)
			except Exception as e:
				log.warning(f"outlookSpeechFilter: nie można otworzyć strony GitHub: {e}")

	def onCopyReport(self, evt):
		rep = self.getReport()
		ok = copy_text_to_clipboard(rep)
		if ok:
			ui.message(_t("CONTACT_REPORT_COPIED_MSG"))
		else:
			ui.message(_t("CONTACT_REPORT_COPY_FAILED"))

	def onViewReport(self, evt):
		rep = self.getReport()
		dlg = wx.Dialog(self, title=_t("REPORT_VIEW_TITLE"), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, size=(600, 480))
		s = wx.BoxSizer(wx.VERTICAL)
		# Widoczne ostrzeżenie o danych osobowych (NVDA odczytuje je jako tekst okna dialogowego)
		warnText = wx.StaticText(dlg, label=_t("REPORT_PRIVACY_WARNING"))
		warnText.Wrap(560)
		s.Add(warnText, flag=wx.LEFT | wx.RIGHT | wx.TOP, border=10)
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
_pending_debounce_call = None
_pending_sound_calls = []
_last_own_sound_time = 0.0
# Pamięć podręczna liczników załączników (EntryID -> (realne, wklejone grafiki)), LRU o ograniczonym rozmiarze
_MSG_ATTACHMENT_CACHE_MAX = 200
_msg_attachment_cache = collections.OrderedDict()


def _attachment_cache_get(entryID):
	if not entryID:
		return None
	val = _msg_attachment_cache.get(entryID)
	if val is not None:
		try:
			_msg_attachment_cache.move_to_end(entryID)
		except Exception:
			pass
	return val


def _attachment_cache_put(entryID, value):
	if not entryID:
		return
	_msg_attachment_cache[entryID] = value
	try:
		_msg_attachment_cache.move_to_end(entryID)
		while len(_msg_attachment_cache) > _MSG_ATTACHMENT_CACHE_MAX:
			_msg_attachment_cache.popitem(last=False)
	except Exception:
		_msg_attachment_cache.clear()


# ==============================================================================
# MIGAWKA KONFIGURACJI (odświeżana po zmianie profilu i po zapisie ustawień)
# ==============================================================================

# Klucz statusu -> (klucz konfiguracji dźwięku, łańcuch zapasowych kluczy zgodności wstecznej, domyślna wartość)
_SOUND_CONF_SPEC = (
	("unread", "soundUnread", ("reportUnread",), True),
	("flag", "soundFlags", ("reportFlags",), False),
	("attachment", "soundAttachments", ("reportAttachments",), True),
	("inline_image", "soundInlineImages", ("reportInlineImages",), True),
	("replied", "soundReplied", ("soundForwardReply", "reportForwardReply"), False),
	("forwarded", "soundForwarded", ("soundForwardReply", "reportForwardReply"), False),
	("importance_high", "soundImportanceHigh", ("soundImportance", "reportImportance"), True),
	("importance_low", "soundImportanceLow", ("soundImportance", "reportImportance"), False),
	("meeting", "soundMeetingRequest", ("reportMeetingRequest",), True),
	("category", "soundCategories", ("reportCategories",), False),
	("external", "soundExternal", ("reportExternal",), False),
)

_conf_snapshot = None
_conf_snapshot_gen = 0


def _conf_get_chain(conf, key, fallbacks, default):
	"""Odczyt klucza z łańcuchem zapasowych kluczy (identycznie jak wcześniej zagnieżdżone conf.get)."""
	val = default
	for fb in reversed(fallbacks):
		try:
			val = conf.get(fb, val)
		except Exception:
			pass
	try:
		val = conf.get(key, val)
	except Exception:
		pass
	return _to_bool(val, default=default)


def _build_conf_snapshot():
	try:
		conf = config.conf.get(CONFIG_SECTION, {})
	except Exception:
		conf = {}
	snap = {
		"enabled": _to_bool(conf.get("enabled", True), default=True),
		"useSounds": _to_bool(conf.get("useSounds", True), default=True),
		"threadNavigationEnabled": _to_bool(conf.get("threadNavigationEnabled", True), default=True),
		"soundVolume": _to_int(conf.get("soundVolume", 80), 80),
		"soundTheme": str(conf.get("soundTheme", "associative") or "associative"),
		"sounds": {},
	}
	for statusKey, confKey, fallbacks, default in _SOUND_CONF_SPEC:
		snap["sounds"][statusKey] = _conf_get_chain(conf, confKey, fallbacks, default)
	return snap


def _get_conf_snapshot():
	global _conf_snapshot
	snap = _conf_snapshot
	if snap is None:
		snap = _build_conf_snapshot()
		_conf_snapshot = snap
	return snap


def _invalidate_conf_snapshot(*args, **kwargs):
	"""Unieważnia migawkę konfiguracji, bufor języka i pamięć podręczną wierszy (np. po zmianie profilu)."""
	global _conf_snapshot, _conf_snapshot_gen, _effective_language_cache
	_conf_snapshot = None
	_effective_language_cache = None
	_conf_snapshot_gen += 1

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


def _get_scaled_sound_path(sound_name, volume=None, theme=None):
	"""
	Returns the path to a .wav sound file scaled to the specified volume (0-100).
	If volume is 100, returns original file path.
	If volume <= 0, returns None (muted).
	Caches scaled audio files in temp directory for zero-latency subsequent playback.
	"""
	if volume is None:
		panel = get_active_settings_panel()
		if panel and hasattr(panel, "soundVolume"):
			volume = panel.soundVolume
		elif panel and hasattr(panel, "soundVolumeSlider"):
			try:
				volume = panel.soundVolumeSlider.GetValue()
			except Exception:
				volume = None
		if volume is None:
			volume = _get_conf_snapshot().get("soundVolume", 80)

	try:
		volume = int(volume)
	except (ValueError, TypeError):
		volume = 100

	if volume <= 0:
		return None

	if theme is None:
		panel = get_active_settings_panel()
		if panel and hasattr(panel, "soundTheme"):
			theme = panel.soundTheme
		elif panel and hasattr(panel, "soundThemeChoice") and hasattr(panel, "themeKeys"):
			try:
				selIdx = panel.soundThemeChoice.GetSelection()
				if 0 <= selIdx < len(panel.themeKeys):
					theme = panel.themeKeys[selIdx]
			except Exception:
				pass
	if not theme:
		theme = _get_conf_snapshot().get("soundTheme", "associative") or "associative"
	if theme not in ("associative", "minimal", "instrumental"):
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


def _play_status_sound(sound_name, volume=None, theme=None):
	"""Play a single status earcon sound file at the specified or configured volume and theme."""
	global _last_own_sound_time
	filepath = _get_scaled_sound_path(sound_name, volume=volume, theme=theme)
	if filepath and os.path.isfile(filepath):
		try:
			nvwave.playWaveFile(filepath, asynchronous=True)
			_last_own_sound_time = time.time()
		except Exception:
			pass


def _stop_own_wave_playback():
	"""
	Zatrzymuje odtwarzanie pliku WAV przez nvwave (nvwave.fileWavePlayer), ale tylko gdy
	ostatni dźwięk pochodził z tego dodatku — aby nie przerywać innych dźwięków NVDA.
	"""
	if time.time() - _last_own_sound_time > 3.0:
		return
	try:
		player = getattr(nvwave, "fileWavePlayer", None)
		if player:
			player.stop()
	except Exception:
		pass


def _stop_pending_sounds():
	"""Stop all pending scheduled sound timers and stop this add-on's wave playback."""
	global _pending_sound_calls, _pending_debounce_call
	if not wx.IsMainThread():
		wx.CallAfter(_stop_pending_sounds)
		return
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
	_stop_own_wave_playback()


def _schedule_status_sounds_sequence(sound_names, volume=None, delay_ms=230, theme=None):
	"""Play a sequence of sounds with spacing in ms, cancelling previous calls (for settings dialog tests)."""
	global _pending_sound_calls
	if not wx.IsMainThread():
		wx.CallAfter(_schedule_status_sounds_sequence, sound_names, volume, delay_ms, theme)
		return
	_stop_pending_sounds()
	if not sound_names:
		return
	_play_status_sound(sound_names[0], volume=volume, theme=theme)
	for i, name in enumerate(sound_names[1:], start=1):
		call = wx.CallLater(i * delay_ms, lambda s=name, v=volume, t=theme: _play_status_sound(s, volume=v, theme=t))
		_pending_sound_calls.append(call)


def _schedule_status_sounds(sound_names, volume=None, debounce_ms=120):
	"""
	Schedule status sounds for playback with debounce and spacing.
	Opóźnia odtwarzanie o debounce_ms. Jeśli użytkownik szybko nawiguje po liście wiadomości,
	poprzednie timery są anulowane i dźwięki wiadomości pośrednich NIE SĄ w ogóle odtwarzane.
	Zapobiega to kolejkowaniu dźwięków i zacinaniu nawigacji na wolniejszych komputerach.
	"""
	global _pending_debounce_call, _pending_sound_calls
	if not wx.IsMainThread():
		wx.CallAfter(_schedule_status_sounds, sound_names, volume, debounce_ms)
		return
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


# Słowa kluczowe nagłówków kolumn statusowych (ikony/stany), w których tokeny typu „Przeczytana”,
# „Załączniki”, „Przesłano dalej” oznaczają status, a nie treść wiadomości.
STATUS_COLUMN_HEADER_KEYWORDS = (
	"załącznik", "attachment",
	"przeczyt", "read",
	"ikona", "icon",
	"ważność", "waznosc", "priorytet", "importance", "priority",
	"flag",
	"zewnętrzn", "external",
	"stan", "status",
	"odpowiedz", "reply", "replied", "przesłan", "forward",
)
# Kolumny z treścią wiadomości — nigdy nie usuwamy z nich tekstu jako „statusu”.
CONTENT_COLUMN_HEADERS = {
	"temat", "subject", "od", "from", "do", "to", "nadawca", "sender", "adresat", "recipient",
	"dw", "cc", "odebrano", "received", "wysłano", "sent", "rozmiar", "size",
	"w folderze", "in folder", "folder", "kategorie", "categories", "podgląd", "preview",
}
# Tokeny jednoznaczne — dopuszczalne do odfiltrowania nawet bez informacji o nagłówku kolumny.
UNAMBIGUOUS_ATTACHMENT_RE = re.compile(r'^(?:has\s+attachments?|ma\s+załączniki?)(?:\s*:\s*\d+)?$', re.IGNORECASE)


def _classify_column(headerNames):
	"""
	Zwraca 'content' (temat, od, ...), 'status' (kolumna ikon/stanów) lub '' (nieznana / brak nagłówka)
	na podstawie nazw nagłówków kolumn UIA (TableItem.ColumnHeaderItems).
	"""
	kind = ""
	for h in headerNames:
		hl = (h or "").strip().lower().rstrip(":").strip()
		if not hl:
			continue
		if hl in CONTENT_COLUMN_HEADERS:
			return "content"
		if any(k in hl for k in STATUS_COLUMN_HEADER_KEYWORDS):
			kind = "status"
	return kind


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

	return _build_status_items({
		"unread": isUnread,
		"flag": hasFlag,
		"attachments": realCount,
		"inline_images": inlineImgCount,
		"verb_key": verbSound,
		"verb_text": verbText,
		"importance": 2 if importanceSound == "importance_high" else (0 if importanceSound == "importance_low" else 1),
		"meeting": isMeeting,
		"category": bool(catText),
		"category_text": catText,
		"external": isExternal,
	})


def _build_status_items(st):
	"""
	Jedyne miejsce budowania listy statusów wiadomości w kolejności CANONICAL_STATUS_ORDER.
	Zwraca listę krotek (klucz_statusu, tekst). Klucz statusu jest jednocześnie nazwą dźwięku.
	"""
	items = []
	for key in CANONICAL_STATUS_ORDER:
		if key == "unread":
			if st.get("unread"):
				items.append((key, _t("STATUS_UNREAD")))
		elif key == "flag":
			if st.get("flag"):
				items.append((key, _t("STATUS_FLAG")))
		elif key == "attachment":
			cnt = st.get("attachments", 0) or 0
			if cnt > 0:
				items.append((key, _format_attachment_status(cnt)))
		elif key == "inline_image":
			cnt = st.get("inline_images", 0) or 0
			if cnt > 0:
				items.append((key, _format_inline_images_status(cnt)))
		elif key in ("replied", "forwarded"):
			if st.get("verb_key") == key and st.get("verb_text"):
				items.append((key, st["verb_text"]))
		elif key == "importance_high":
			if st.get("importance") == 2:
				items.append((key, _t("STATUS_IMPORTANCE_HIGH")))
		elif key == "importance_low":
			if st.get("importance") == 0:
				items.append((key, _t("STATUS_IMPORTANCE_LOW")))
		elif key == "meeting":
			if st.get("meeting"):
				items.append((key, _t("STATUS_MEETING")))
		elif key == "category":
			if st.get("category") or st.get("category_text"):
				items.append((key, st.get("category_text") or _t("STATUS_CATEGORY", cat="")))
		elif key == "external":
			if st.get("external"):
				items.append((key, _t("STATUS_EXTERNAL")))
	return items


def _sounds_for_status_items(items, snap=None):
	"""Zwraca listę dźwięków do odtworzenia dla podanych statusów zgodnie z konfiguracją."""
	snap = snap or _get_conf_snapshot()
	if not snap.get("useSounds"):
		return []
	sounds = []
	for key, _text in items:
		if not snap["sounds"].get(key):
			continue
		if key == "inline_image" and _is_image_helper_handling_inline_sounds():
			continue
		sounds.append(key)
	return sounds


def _status_order_key(key):
	"""Dokładna pozycja statusu w CANONICAL_STATUS_ORDER (bez dopasowywania podciągów)."""
	try:
		return CANONICAL_STATUS_ORDER.index(key)
	except ValueError:
		return len(CANONICAL_STATUS_ORDER)


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
		countText = _t("GROUP_ITEMS_COUNT", count=countMatch.group(1))

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

# Klasy UIA wierszy wiadomości w Outlooku — zgodnie z nvdaBuiltin.appModules.outlook
# (AppModule.chooseNVDAObjectOverlayClasses: LeafRow, ThreadItem, ThreadHeader -> UIAGridRow).
# ThreadHeader to nagłówek konwersacji traktowany przez NVDA jako wiersz wiadomości (poziom 1), a nie nagłówek grupy.
MESSAGE_ROW_CLASSES = ("LeafRow", "ThreadItem", "ThreadHeader")
# Nagłówki grup (np. „Grupuj według: Kategorie”) rozpoznawane wyłącznie po klasie UIA
GROUP_HEADER_CLASSES = ("GroupHeader", "OutlookGridGroupHeader")
# Statusy, których ustalenie wymaga modelu obiektowego Outlooka (COM)
COM_STATUS_KEYS = ("unread", "flag", "attachment", "inline_image", "importance_high", "importance_low", "meeting", "category", "external")
_RE_SUBJECT_PREFIX = re.compile(r'^\s*(?:(?:re|fw|fwd|odp|pd|aw|wg|tr|sv|vs)\s*:\s*)+', re.IGNORECASE)
_NO_CATEGORY_NAMES = ("bez kategorii", "brak kategorii", "no category", "no categories", "none")


def _custom_UIAGridRow_get_name(self):
	"""
	Getter nazwy wiersza listy wiadomości Outlooka.
	Nie ma efektów ubocznych (dźwięki są odtwarzane w event_gainFocus wiersza),
	a każdy błąd (np. COMError) kończy się powrotem do oryginalnej nazwy zamiast ciszy.
	"""
	global _in_custom_get_name
	if _in_custom_get_name:
		return _fallback_row_name(self)
	_in_custom_get_name = True
	try:
		return _get_row_info(self)[0]
	except Exception:
		log.debugWarning("outlookSpeechFilter: błąd budowania nazwy wiersza Outlooka, używam nazwy oryginalnej", exc_info=True)
		return _fallback_row_name(self)
	finally:
		_in_custom_get_name = False


def _fallback_row_name(obj):
	if _original_UIAGridRow_get_name:
		try:
			return _original_UIAGridRow_get_name(obj)
		except Exception:
			log.debugWarning("outlookSpeechFilter: oryginalny UIAGridRow._get_name zgłosił wyjątek", exc_info=True)
	return _raw_uia_name(obj)


def _raw_uia_name(obj):
	el = getattr(obj, "UIAElement", None)
	if el is None:
		return ""
	name = ""
	try:
		name = el.cachedName
	except Exception:
		name = ""
	if not name:
		try:
			name = el.currentName
		except Exception:
			name = ""
	return name or ""


def _uia_class_name(obj):
	try:
		return getattr(getattr(obj, "UIAElement", None), "cachedClassName", "") or ""
	except Exception:
		return ""


def _is_row_related_to_focus(obj):
	try:
		curFocus = api.getFocusObject()
		if curFocus is obj:
			return True
		return (getattr(curFocus, "parent", None) is obj) or (getattr(obj, "parent", None) is curFocus)
	except Exception:
		return True


def _get_row_info(obj):
	"""
	Zwraca (tekst_do_odczytu, statusy[(klucz, tekst)], dźwięki[]) dla wiersza.
	Wynik jest zapamiętywany na obiekcie (do zmiany surowej nazwy UIA, stanu rozwinięcia,
	konfiguracji lub fokusu), aby event_gainFocus i odczyt nazwy nie odpytywały COM dwukrotnie.
	"""
	snap = _get_conf_snapshot()
	rawName = _raw_uia_name(obj)
	related = _is_row_related_to_focus(obj)
	try:
		states = obj.states
	except Exception:
		states = set()
	expandKey = (controlTypes.State.EXPANDED in states, controlTypes.State.COLLAPSED in states)
	key = (rawName, expandKey, related, _conf_snapshot_gen, get_effective_language())
	cached = getattr(obj, "_osf_rowCache", None)
	if cached and cached[0] == key:
		return cached[1]
	result = _compute_row_info(obj, snap, rawName, related, states)
	try:
		obj._osf_rowCache = (key, result)
	except Exception:
		pass
	return result


def _get_single_selection(obj):
	"""Zwraca element zaznaczony w Outlooku (COM), gdy zaznaczony jest dokładnie jeden."""
	try:
		nativeOm = obj.appModule.nativeOm
		if not nativeOm:
			return None
		explorer = nativeOm.activeExplorer()
		sel = getattr(explorer, "selection", None)
		if sel is not None and getattr(sel, "Count", 0) == 1:
			return sel.item(1)
	except (COMError, Exception):
		pass
	return None


def _normalize_subject(subject):
	s = (subject or "").strip().lower()
	s = _RE_SUBJECT_PREFIX.sub("", s)
	for tag in EXTERNAL_PREFIXES:
		if s.startswith(tag):
			s = s[len(tag):].strip(" -:,")
			break
	s = _RE_SUBJECT_PREFIX.sub("", s)
	return re.sub(r"\s+", " ", s).strip()


def _selection_matches_row(selection, rawName, cellTexts):
	"""
	Chroni przed wyścigiem „nieaktualnego zaznaczenia”: dane COM są używane tylko wtedy,
	gdy temat zaznaczonego elementu występuje w tekście wiersza, którego nazwę budujemy.
	"""
	try:
		subject = selection.Subject
	except Exception:
		return True
	ns = _normalize_subject(subject)
	if not ns:
		return True
	hay = re.sub(r"\s+", " ", (rawName + " " + " ".join(cellTexts)).lower())
	if not hay.strip():
		return True
	return ns[:60] in hay


def _get_attachment_counts(selection, uiaObj, needInline):
	"""Liczniki (realne załączniki, wklejone grafiki) z pamięcią podręczną LRU wg EntryID."""
	entryID = None
	try:
		entryID = selection.EntryID
	except Exception:
		entryID = None
	cachedVal = _attachment_cache_get(entryID)
	if cachedVal is not None:
		real, inline = cachedVal
		if inline is not None or not needInline:
			return real, (inline or 0)
	else:
		real = _get_real_attachments_count(selection, uiaObj)
	inline = _get_inline_images_count(selection) if needInline else None
	_attachment_cache_put(entryID, (real, inline))
	return real, (inline or 0)


def _compute_row_info(obj, snap, rawName, related, states):
	enabled = snap["enabled"]
	useSounds = snap["useSounds"]
	speakStatuses = not enabled
	cachedCls = _uia_class_name(obj)

	# 1. Nagłówki grup (grupowanie wg flag lub kategorii) — rozpoznawane wyłącznie po klasie UIA
	if cachedCls in GROUP_HEADER_CLASSES:
		headerName = rawName or _fallback_row_name(obj)
		cleanHeader = _clean_group_header_name(headerName)
		if cleanHeader:
			if controlTypes.State.COLLAPSED in states:
				return (f"{cleanHeader}, {controlTypes.State.COLLAPSED.displayString}", [], [])
			if controlTypes.State.EXPANDED in states:
				return (f"{cleanHeader}, {controlTypes.State.EXPANDED.displayString}", [], [])
			return (cleanHeader, [], [])

	textList = []
	if controlTypes.State.EXPANDED in states:
		textList.append(controlTypes.State.EXPANDED.displayString)
	elif controlTypes.State.COLLAPSED in states:
		textList.append(controlTypes.State.COLLAPSED.displayString)

	st = {
		"unread": False,
		"flag": False,
		"attachments": 0,
		"inline_images": 0,
		"verb_key": None,
		"verb_text": None,
		"importance": 1,
		"meeting": False,
		"category": False,
		"category_text": None,
		"external": False,
	}
	uiaUnread = None
	uiaAttachmentCount = 0

	# 2. Komórki UIA wiersza (preferowane źródło danych — bez odwołań do COM)
	conditionFactory = createUIAMultiPropertyCondition or getattr(builtinOutlook, "createUIAMultiPropertyCondition", None)
	childrenCacheRequest = UIAHandler.handler.baseCacheRequest.clone()
	childrenCacheRequest.addProperty(UIAHandler.UIA_NamePropertyId)
	childrenCacheRequest.addProperty(UIAHandler.UIA_TableItemColumnHeaderItemsPropertyId)
	childrenCacheRequest.TreeScope = UIAHandler.TreeScope_Children
	childrenCacheRequest.treeFilter = conditionFactory(
		{
			UIAHandler.UIA_ControlTypePropertyId: [
				UIAHandler.UIA_TextControlTypeId,
				UIAHandler.UIA_ImageControlTypeId,
			],
		},
	)
	cachedChildren = obj.UIAElement.buildUpdatedCache(childrenCacheRequest).getCachedChildren()
	if not cachedChildren:
		log.debugWarning("Unable to get relevant children for UIAGridRow", stack_info=True)
		return (_fallback_row_name(obj), [], [])

	try:
		reportFlag = config.conf["documentFormatting"]["reportTableHeaders"]
	except Exception:
		reportFlag = None
	allowedFlags = (getattr(ReportTableHeaders, "ROWS_AND_COLUMNS", 2), getattr(ReportTableHeaders, "COLUMNS", 1)) if ReportTableHeaders else (2, 1)
	reportHeaders = reportFlag in allowedFlags

	contentCells = []
	cellTexts = []

	for index in range(cachedChildren.length):
		e = cachedChildren.getElement(index)
		UIAControlType = e.cachedControlType
		UIAClassName = e.cachedClassName

		if UIAClassName == "FlagField":
			fName = (e.cachedName or "").lower()
			if fName and any(w in fName for w in ("oflagow", "flaga", "flag")):
				st["flag"] = True
			continue
		elif UIAClassName == "CategoryField":
			cName = (e.cachedName or "").strip()
			if cName and cName.lower() not in _NO_CATEGORY_NAMES:
				st["category"] = True
				if not st["category_text"]:
					st["category_text"] = _t("STATUS_CATEGORY", cat=cName)
			continue
		elif UIAControlType != UIAHandler.UIA_TextControlTypeId:
			continue

		name = e.cachedName
		if not name:
			continue
		cellTexts.append(name)
		cleanName = name.strip().lower()
		cleanToken = cleanName.rstrip(":").strip()

		# Nagłówki kolumn tabeli (do klasyfikacji kolumny oraz ewentualnego odczytu nagłówków)
		headerNames = []
		try:
			columnHeaderItems = e.getCachedPropertyValueEx(
				UIAHandler.UIA_TableItemColumnHeaderItemsPropertyId,
				True,
			)
			if columnHeaderItems:
				columnHeaderItems = columnHeaderItems.QueryInterface(UIAHandler.IUIAutomationElementArray)
				for hIdx in range(columnHeaderItems.length):
					hName = columnHeaderItems.getElement(hIdx).currentName
					if hName:
						headerNames.append(hName)
		except Exception:
			pass

		columnKind = _classify_column(headerNames)
		isAttachmentHeader = columnKind != "content" and any(
			("załącznik" in h.lower() or "attachment" in h.lower()) for h in headerNames
		)

		# Kolumna załączników — nie czytamy surowego „1” lub „Załączniki: 1”
		if isAttachmentHeader:
			uiaAttachmentCount = max(uiaAttachmentCount, _extract_attachment_count(name) or 1)
			continue

		# Tokeny statusów są interpretowane tylko w kolumnach statusowych; kolumny z treścią
		# (Temat, Od, Do...) nigdy nie są usuwane. Bez informacji o kolumnie akceptujemy tylko tokeny jednoznaczne.
		if columnKind == "status":
			if cleanName in FORWARD_REPLY_TOKENS:
				if not st["verb_text"]:
					st["verb_text"] = name
					st["verb_key"] = "forwarded" if any(w in cleanName for w in ("przesł", "forward")) else "replied"
				continue
			if cleanName in EXTERNAL_TOKENS or cleanToken in EXTERNAL_TOKENS:
				st["external"] = True
				continue
			if _is_attachment_text(cleanName):
				uiaAttachmentCount = max(uiaAttachmentCount, _extract_attachment_count(cleanName) or 1)
				continue
			if cleanName in UNREAD_TOKENS:
				uiaUnread = True
				continue
			if cleanName in READ_TOKENS:
				uiaUnread = False
				continue
		elif columnKind == "" and not headerNames:
			if cleanName in FORWARD_REPLY_TOKENS:
				if not st["verb_text"]:
					st["verb_text"] = name
					st["verb_key"] = "forwarded" if any(w in cleanName for w in ("przesł", "forward")) else "replied"
				continue
			if cleanName.startswith("[") and (cleanName in EXTERNAL_TOKENS or cleanToken in EXTERNAL_TOKENS):
				st["external"] = True
				continue
			if UNAMBIGUOUS_ATTACHMENT_RE.match(cleanName):
				uiaAttachmentCount = max(uiaAttachmentCount, _extract_attachment_count(cleanName) or 1)
				continue
			if cleanName in UNREAD_TOKENS:
				uiaUnread = True
				continue

		columnHeaderTextList = []
		if reportHeaders:
			for hName in headerNames:
				hClean = hName.strip().lower()
				if not speakStatuses:
					if _is_attachment_text(hClean) or hClean in UNREAD_TOKENS or hClean in READ_TOKENS or hClean in FORWARD_REPLY_TOKENS or hClean in EXTERNAL_TOKENS or hClean.rstrip(":").strip() in EXTERNAL_TOKENS:
						continue
				columnHeaderTextList.append(hName)

		columnHeaderText = " ".join(columnHeaderTextList)
		text = f"{columnHeaderText} {name}" if columnHeaderText else name

		# Wykrywanie i wycinanie prefiksów [Zewnętrzna] (doklejanych do tematu przez serwer)
		if text:
			for tag in EXTERNAL_PREFIXES:
				if text.strip().lower().startswith(tag):
					st["external"] = True
					text = text.strip()[len(tag):].lstrip(" -:,")
					break

		if text:
			contentCells.append(text + ",")

	# 3. Dane z modelu obiektowego Outlooka (COM) — tylko gdy są potrzebne, gdy wiersz jest
	# powiązany z fokusem i gdy zaznaczony element odpowiada temu wierszowi.
	needCom = speakStatuses or (useSounds and any(snap["sounds"].get(k) for k in COM_STATUS_KEYS))
	selection = None
	if needCom and related:
		selection = _get_single_selection(obj)
		if selection is not None and not _selection_matches_row(selection, rawName, cellTexts):
			log.debug("outlookSpeechFilter: zaznaczenie COM nie odpowiada wierszowi — pomijam dane COM")
			selection = None

	if selection is not None:
		try:
			st["unread"] = bool(selection.unread)
		except (COMError, Exception):
			pass
		if not st["flag"]:
			try:
				st["flag"] = bool(selection.flagIcon or selection.flagStatus)
			except (COMError, Exception):
				pass
		needInline = speakStatuses or snap["sounds"].get("inline_image", True)
		realCount, inlineCount = _get_attachment_counts(selection, obj, needInline)
		st["attachments"] = realCount if realCount > 0 else uiaAttachmentCount
		st["inline_images"] = inlineCount
		try:
			st["importance"] = selection.importance
		except (COMError, Exception):
			st["importance"] = 1
		try:
			st["meeting"] = (selection.messageClass == "IPM.Schedule.Meeting.Request")
		except (COMError, Exception):
			pass
		try:
			cat = (selection.categories or "").strip()
			if cat and cat.lower() not in _NO_CATEGORY_NAMES:
				st["category"] = True
				st["category_text"] = _t("STATUS_CATEGORY", cat=cat)
		except (COMError, Exception):
			pass
		if not st["external"]:
			st["external"] = _is_external_message(selection, obj)
	else:
		st["attachments"] = uiaAttachmentCount

	if uiaUnread is not None:
		st["unread"] = uiaUnread

	items = _build_status_items(st)
	sounds = _sounds_for_status_items(items, snap)

	# Gdy komunikaty mowy są włączone (filtr wyłączony), statusy są czytane na początku wiersza
	if speakStatuses and items:
		textList.extend([text + "," for _key, text in items])
	textList.extend(contentCells)
	return (" ".join(textList), items, sounds)


def _on_message_row_focus(obj):
	"""Wywoływane z event_gainFocus wiersza: zapamiętuje statusy i odtwarza dźwięki (wątek główny)."""
	global _last_focused_statuses
	_text, items, sounds = _get_row_info(obj)
	_last_focused_statuses = list(items)
	if sounds:
		_schedule_status_sounds(sounds)
	else:
		_stop_pending_sounds()


def _is_image_helper_handling_inline_sounds() -> bool:
	"""
	Sprawdza, czy dodatek Outlook Image Helper jest aktywny i obsługuje
	dźwięki obecności grafik w treści na liście wiadomości.
	Zapobiega podwójnemu odtwarzaniu dźwięku grafiki przez oba dodatki jednocześnie.
	"""
	try:
		oih_running = False
		for p in getattr(globalPluginHandler, "runningPlugins", []):
			mod_name = getattr(p, "__module__", "") or getattr(getattr(p, "__class__", None), "__module__", "")
			if "outlookImageHelper" in mod_name:
				oih_running = True
				break
		if not oih_running:
			return False
		oih_conf = config.conf.get("outlookImageHelper", {})
		if not oih_conf:
			return True
		return _to_bool(oih_conf.get("playSound", True), default=True)
	except Exception:
		return False


class _StatusMaskedItemProxy:
	"""
	Pośrednik elementu Outlooka dla oryginalnej funkcji NVDA getReceivedMessageString:
	ukrywa wyłącznie stan nieprzeczytania i załączniki (wyciszane przez dodatek),
	zachowując wszystkie pozostałe pola i tłumaczenia NVDA.
	"""

	class _NoAttachments:
		count = 0
		Count = 0

	def __init__(self, item):
		object.__setattr__(self, "_osfItem", item)

	def __getattr__(self, name):
		low = name.lower()
		if low == "unread":
			return False
		if low == "attachments":
			return _StatusMaskedItemProxy._NoAttachments()
		return getattr(object.__getattribute__(self, "_osfItem"), name)


def _custom_getReceivedMessageString(obj):
	"""Zastępnik getReceivedMessageString: bez efektów ubocznych, bez zakodowanych na sztywno polskich etykiet."""
	orig = _original_getReceivedMessageString
	if orig is None:
		return ""
	if not _get_conf_snapshot()["enabled"]:
		return orig(obj)
	try:
		return orig(_StatusMaskedItemProxy(obj))
	except Exception:
		log.debugWarning("outlookSpeechFilter: getReceivedMessageString przez pośrednika nie powiodło się", exc_info=True)
		return orig(obj)


_saved_class_methods = {}


def _safe_patch_method(target_obj, attr_name, new_method):
	"""Bezpiecznie podmienia atrybut, zachowując indywidualny oryginał dla danego obiektu."""
	key = (target_obj, attr_name)
	if key not in _saved_class_methods:
		orig = getattr(target_obj, attr_name, None)
		if orig is not None:
			_saved_class_methods[key] = orig
	setattr(target_obj, attr_name, new_method)


def patchOutlookClasses():
	"""
	Zapamiętuje oryginalny UIAGridRow._get_name (używany jako awaryjna nazwa) i podmienia
	funkcję modułową getReceivedMessageString.
	Uwaga: nie podmieniamy już UIAGridRow._get_name na klasie — AutoPropertyType (baseObject)
	wiąże getter z właściwością „name” w chwili tworzenia klasy, więc taka podmiana nie działała.
	Filtrowanie zapewnia nakładka CleanOutlookGridRow.
	"""
	global _original_UIAGridRow_get_name, _original_getReceivedMessageString, _patches_applied
	if _patches_applied:
		return

	if builtinOutlook and hasattr(builtinOutlook, "UIAGridRow"):
		if _original_UIAGridRow_get_name is None:
			_original_UIAGridRow_get_name = builtinOutlook.UIAGridRow._get_name

	appModulePatched = False
	for modName in list(sys.modules.keys()):
		if modName.endswith("appModules.outlook"):
			mod = sys.modules.get(modName)
			if mod is None or not hasattr(mod, "getReceivedMessageString"):
				continue
			if modName == "appModules.outlook":
				appModulePatched = True
			current = getattr(mod, "getReceivedMessageString", None)
			if current is _custom_getReceivedMessageString:
				continue
			if _original_getReceivedMessageString is None:
				_original_getReceivedMessageString = current
			_safe_patch_method(mod, "getReceivedMessageString", _custom_getReceivedMessageString)

	# Zakończ dopiero, gdy załadowany jest moduł aplikacji Outlooka (appModules.outlook)
	_patches_applied = appModulePatched


class CleanOutlookGridRow(builtinOutlook.UIAGridRow if builtinOutlook else object):
	"""Nakładka na wiersze listy wiadomości Outlooka (wyłącznie obiekty UIA o klasach wierszy/nagłówków grup)."""

	def _get_name(self):
		return _custom_UIAGridRow_get_name(self)

	def _get_description(self):
		try:
			rawDesc = super().description
		except Exception:
			return ""
		if _uia_class_name(self) in GROUP_HEADER_CLASSES:
			rawName = _raw_uia_name(self)
			for txt in (rawName, rawDesc):
				if txt and any(k in txt.lower() for k in ("grupuj według", "group by", "kategorie:", "kategoria:")):
					return ""
		return rawDesc

	def event_gainFocus(self):
		try:
			_on_message_row_focus(self)
		except Exception:
			log.debugWarning("outlookSpeechFilter: błąd obsługi fokusu wiersza", exc_info=True)
		superHandler = getattr(super(), "event_gainFocus", None)
		if superHandler:
			superHandler()

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
			title=_t("UPD_DLG_TITLE", title=addon_title),
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER
		)
		self.on_install_callback = on_install_callback

		main_sizer = wx.BoxSizer(wx.VERTICAL)

		# Nagłówek
		lbl_header = wx.StaticText(self, label=_t("UPD_HEADER", title=addon_title))
		font = lbl_header.GetFont()
		font.SetWeight(wx.FONTWEIGHT_BOLD)
		lbl_header.SetFont(font)
		main_sizer.Add(lbl_header, 0, wx.ALL, 12)

		# Informacja o wersjach
		lbl_ver = wx.StaticText(self, label=_t("UPD_VERSIONS", local=local_ver, remote=remote_ver))
		main_sizer.Add(lbl_ver, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		# Etykieta zmian
		lbl_changes = wx.StaticText(self, label=_t("UPD_CHANGES_LABEL"))
		main_sizer.Add(lbl_changes, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 4)

		# Pole tekstowe z opisem zmian (dostępne, wieloliniowe, tylko do odczytu)
		self.txt_changelog = wx.TextCtrl(
			self,
			value=changelog or _t("UPD_NO_CHANGELOG"),
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL,
			size=(520, 160)
		)
		main_sizer.Add(self.txt_changelog, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 12)

		# Komunikat pytający
		lbl_prompt = wx.StaticText(self, label=_t("UPD_PROMPT"))
		main_sizer.Add(lbl_prompt, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 12)

		# Przyciski akcji
		btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
		self.btn_install = wx.Button(self, wx.ID_YES, label=_t("UPD_BTN_INSTALL"))
		self.btn_cancel = wx.Button(self, wx.ID_CANCEL, label=_t("UPD_BTN_LATER"))

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


class UpdateSecurityError(Exception):
	"""Błąd weryfikacji bezpieczeństwa aktualizacji (adres, suma kontrolna, wersja)."""


def _is_allowed_update_url(url: str) -> bool:
	"""Dozwolone są wyłącznie adresy https:// na hostach GitHub z listy UPDATE_ALLOWED_HOSTS."""
	try:
		parts = urllib.parse.urlsplit(str(url or "").strip())
	except Exception:
		return False
	if parts.scheme.lower() != "https":
		return False
	host = (parts.hostname or "").lower()
	return host in UPDATE_ALLOWED_HOSTS


class _GitHubOnlyRedirectHandler(urllib.request.HTTPRedirectHandler):
	"""Podąża za przekierowaniami tylko wtedy, gdy cel jest również dozwolonym adresem HTTPS GitHub."""

	def redirect_request(self, req, fp, code, msg, headers, newurl):
		absolute = urllib.parse.urljoin(req.full_url, newurl)
		if not _is_allowed_update_url(absolute):
			raise UpdateSecurityError(_t("UPD_ERR_URL", url=absolute))
		return super().redirect_request(req, fp, code, msg, headers, absolute)


def _open_update_url(url: str, timeout: int):
	"""Otwiera adres HTTPS GitHub, weryfikując adres początkowy, przekierowania i adres końcowy."""
	if not _is_allowed_update_url(url):
		raise UpdateSecurityError(_t("UPD_ERR_URL", url=url))
	opener = urllib.request.build_opener(_GitHubOnlyRedirectHandler())
	req = urllib.request.Request(url, headers={"User-Agent": "NVDA-Addon-Updater/1.0"})
	resp = opener.open(req, timeout=timeout)
	finalUrl = resp.geturl()
	if not _is_allowed_update_url(finalUrl):
		resp.close()
		raise UpdateSecurityError(_t("UPD_ERR_URL", url=finalUrl))
	return resp


class AddonUpdaterEngine:
	def __init__(self, addon_name: str, addon_title: str, current_version: str, addon_dir: str, config_section: str):
		self.addon_name = addon_name
		self.addon_title = addon_title
		self.current_version = current_version
		self.addon_dir = addon_dir
		self.config_section = config_section

	def get_configured_source(self) -> str:
		"""Adres źródła z konfiguracji — akceptowany tylko, jeśli jest dozwolonym adresem HTTPS GitHub."""
		try:
			src = str(config.conf.get(self.config_section, {}).get("updateSourceUrl", DEFAULT_UPDATE_SOURCE_URL)).strip()
		except Exception:
			src = ""
		if not src or not _is_allowed_update_url(src):
			return DEFAULT_UPDATE_SOURCE_URL
		return src

	def fetch_manifest(self) -> tuple[dict | None, str | None]:
		"""Pobiera informację o aktualizacji wyłącznie z HTTPS GitHub (bez plików lokalnych)."""
		candidates = []
		for s in (self.get_configured_source(), DEFAULT_UPDATE_SOURCE_URL, SECONDARY_UPDATE_SOURCE_URL):
			if s and s not in candidates:
				candidates.append(s)

		raw_content = None
		resolved_source = None
		last_err = None

		for s in candidates:
			try:
				with _open_update_url(s, timeout=10) as resp:
					raw_content = resp.read(1024 * 1024).decode("utf-8-sig")
					resolved_source = s
					break
			except Exception as ex:
				last_err = ex
				continue

		if not raw_content:
			return None, _t("UPD_ERR_NETWORK", err=last_err or "")

		try:
			data = json.loads(raw_content)
			if isinstance(data, dict) and isinstance(data.get(self.addon_name), dict):
				item = dict(data[self.addon_name])
			elif isinstance(data, dict) and "version" in data:
				item = dict(data)
			else:
				return None, _t("UPD_ERR_NO_ENTRY", name=self.addon_name)

			item["_source"] = resolved_source
			return item, None
		except Exception as ex:
			return None, _t("UPD_ERR_JSON", err=ex)

	def check_for_updates(self, manual: bool = False):
		t = threading.Thread(target=self._worker_check, args=(manual,), daemon=True)
		t.start()

	def _record_successful_check(self):
		"""Zapisuje czas udanego sprawdzenia (w wątku głównym)."""
		def _set():
			try:
				config.conf[self.config_section]["lastUpdateCheckTime"] = time.time()
			except Exception:
				pass
		wx.CallAfter(_set)

	def _worker_check(self, manual: bool):
		try:
			info, err = self.fetch_manifest()
			if not info:
				if manual:
					wx.CallAfter(
						showAccessibleMessageBox,
						_t("UPD_CHECK_FAILED", title=self.addon_title, err=err),
						_t("UPD_TITLE", title=self.addon_title),
					)
				return

			self._record_successful_check()
			remote_ver = str(info.get("version", "")).strip()
			if _is_newer(remote_ver, self.current_version):
				download_url = str(info.get("downloadUrl", "") or "")
				changelog = info.get("changelog", "")
				expected_sha = str(info.get("sha256", "") or "").strip().lower()

				def _show_dialog():
					def _on_confirm():
						self.perform_update(download_url, expected_sha)
					dlg = UpdateNotificationDialog(
						parent=gui.mainFrame,
						addon_title=self.addon_title,
						local_ver=self.current_version,
						remote_ver=remote_ver,
						changelog=changelog,
						on_install_callback=_on_confirm
					)
					gui.mainFrame.prePopup()
					try:
						dlg.ShowModal()
					finally:
						dlg.Destroy()
						gui.mainFrame.postPopup()

				wx.CallAfter(_show_dialog)
			else:
				if manual:
					wx.CallAfter(
						showAccessibleMessageBox,
						_t("UPD_UP_TO_DATE", title=self.addon_title, ver=self.current_version),
						_t("UPD_TITLE", title=self.addon_title),
					)
		except Exception as ex:
			log.debugWarning(f"{self.addon_name} updater check error: {ex}", exc_info=True)
			if manual:
				wx.CallAfter(
					showAccessibleMessageBox,
					_t("UPD_UNEXPECTED_ERROR", err=ex),
					_t("UPD_ERROR_TITLE", title=self.addon_title),
				)

	def perform_update(self, download_url: str, expected_sha: str = ""):
		t = threading.Thread(target=self._worker_download_install, args=(download_url, expected_sha), daemon=True)
		t.start()

	def _download_package(self, download_url: str, expected_sha: str) -> str:
		"""Pobiera pakiet z HTTPS GitHub do pliku tymczasowego i weryfikuje sumę SHA-256."""
		if not re.fullmatch(r"[0-9a-f]{64}", expected_sha or ""):
			raise UpdateSecurityError(_t("UPD_ERR_NO_SHA"))
		fd, tmp_file = tempfile.mkstemp(prefix=f"{self.addon_name}_update_", suffix=".nvda-addon")
		os.close(fd)
		hasher = hashlib.sha256()
		total = 0
		try:
			with _open_update_url(download_url, timeout=30) as resp, open(tmp_file, "wb") as out:
				while True:
					chunk = resp.read(64 * 1024)
					if not chunk:
						break
					total += len(chunk)
					if total > UPDATE_MAX_PACKAGE_SIZE:
						raise UpdateSecurityError(_t("UPD_ERR_TOO_BIG"))
					hasher.update(chunk)
					out.write(chunk)
			if hasher.hexdigest().lower() != expected_sha:
				raise UpdateSecurityError(_t("UPD_ERR_SHA_MISMATCH"))
			return tmp_file
		except Exception:
			try:
				os.remove(tmp_file)
			except Exception:
				pass
			raise

	def _worker_download_install(self, download_url: str, expected_sha: str = ""):
		import addonHandler

		wx.CallAfter(ui.message, _t("UPD_DOWNLOADING", title=self.addon_title))
		target_pkg = None

		try:
			target_pkg = self._download_package(download_url, expected_sha)

			bundle = addonHandler.AddonBundle(target_pkg)
			manifest_name = str(bundle.manifest.get("name") or "").strip()
			manifest_ver = str(bundle.manifest.get("version") or "").strip()
			if not manifest_name or not manifest_ver:
				raise ValueError(_t("UPD_ERR_MANIFEST"))
			if manifest_name != self.addon_name:
				raise ValueError(_t("UPD_ERR_NAME", expected=self.addon_name, got=manifest_name))
			if not _is_newer(manifest_ver, self.current_version):
				raise UpdateSecurityError(_t("UPD_ERR_NOT_NEWER", ver=manifest_ver, cur=self.current_version))

			def _do_install():
				installed_ok = False
				install_err_msg = ""
				try:
					try:
						import addonHandler.addonVersionCheck as avc
						if hasattr(avc, "isAddonCompatible") and not avc.isAddonCompatible(bundle):
							canOverride = getattr(bundle, "canOverrideCompatibility", False)
							if callable(canOverride):
								canOverride = canOverride()
							if canOverride and hasattr(bundle, "enableCompatibilityOverride"):
								# Zamiast cichego obejścia zgodności — pytanie do użytkownika
								if showAccessibleConfirmBox(
									_t("UPD_OVERRIDE_CONFIRM", title=self.addon_title, ver=manifest_ver),
									_t("UPD_OVERRIDE_TITLE"),
									defaultYes=False,
								):
									bundle.enableCompatibilityOverride()
								else:
									raise RuntimeError(_t("UPD_OVERRIDE_DECLINED"))
							else:
								raise RuntimeError(_t("UPD_INCOMPATIBLE"))
					except ImportError:
						pass

					prevAddon = None
					try:
						from addonStore.install import _getPreviouslyInstalledAddonById
						prevAddon = _getPreviouslyInstalledAddonById(bundle)
					except Exception:
						try:
							for a in addonHandler.getAvailableAddons():
								if getattr(a, "name", None) == manifest_name:
									prevAddon = a
									break
						except Exception:
							prevAddon = None

					addonObj = None
					if hasattr(addonHandler, "installAddonBundle"):
						addonObj = addonHandler.installAddonBundle(bundle)
					elif hasattr(addonHandler, "installAddonPackage"):
						addonObj = addonHandler.installAddonPackage(bundle)
					elif hasattr(addonHandler, "installAddon"):
						addonObj = addonHandler.installAddon(bundle)

					install_exceptions = getattr(bundle, "_installExceptions", None)
					if addonObj is not None and not install_exceptions:
						if prevAddon is not None:
							try:
								from addonStore.dataManager import addonDataManager
								if addonDataManager:
									addonDataManager._deleteCacheInstalledAddon(prevAddon.name)
							except Exception:
								pass
							try:
								prevAddon.requestRemove()
							except Exception as remEx:
								log.warning(f"{self.addon_name} updater prevAddon.requestRemove warning: {remEx}")
						if hasattr(addonObj, "_cleanupAddonImports"):
							try:
								addonObj._cleanupAddonImports()
							except Exception:
								pass
						installed_ok = True
					else:
						if install_exceptions:
							install_err_msg = "; ".join(str(e) for e in install_exceptions)
						else:
							install_err_msg = _t("UPD_INSTALL_RETURNED_NONE")
				except Exception as ex:
					install_err_msg = str(ex)
					log.error(f"{self.addon_name} updater internal install error: {ex}", exc_info=True)
				finally:
					if target_pkg and os.path.isfile(target_pkg):
						try:
							os.remove(target_pkg)
						except Exception:
							pass

				if installed_ok:
					if showAccessibleConfirmBox(
						_t("UPD_SUCCESS", title=self.addon_title, ver=manifest_ver),
						_t("UPD_SUCCESS_TITLE"),
						defaultYes=True,
					):
						core.restart()
					return

				showAccessibleMessageBox(
					_t("UPD_INSTALL_FAILED", title=self.addon_title, err=install_err_msg or _t("UPD_UNKNOWN_INSTALL_ERROR")),
					_t("UPD_INSTALL_FAILED_TITLE", title=self.addon_title),
				)

			wx.CallAfter(_do_install)
		except Exception as ex:
			if target_pkg and os.path.isfile(target_pkg):
				try:
					os.remove(target_pkg)
				except Exception:
					pass
			log.warning(f"{self.addon_name} updater: {ex}")
			wx.CallAfter(
				showAccessibleMessageBox,
				_t("UPD_DOWNLOAD_FAILED", err=ex),
				_t("UPD_DOWNLOAD_FAILED_TITLE", title=self.addon_title),
			)


class ShareAddonDialog(wx.Dialog):
	"""
	W pełni dostępne okno dialogowe umożliwiające podzielenie się dodatkiem ze znajomymi.
	Zawiera bezpośredni link do pobrania instalatora z GitHub Releases, oficjalną stronę
	projektu na GitHubie oraz gotową treść wiadomości e-mail polecającej dodatek.
	"""
	def __init__(self, parent):
		isPl = (get_effective_language() == "pl")
		title = _t("SHARE_DLG_TITLE")
		super().__init__(
			parent or gui.mainFrame,
			title=title,
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
			size=(640, 480)
		)
		self.SetMinSize((600, 440))
		mainSizer = wx.BoxSizer(wx.VERTICAL)

		lblIntro = wx.StaticText(
			self,
			label=(
				"Skorzystaj z poniższych przycisków, aby skopiować link do pobrania, "
				"skopiować gotową treść lub od razu otworzyć nowy e-mail w programie Outlook:"
				if isPl else
				"Use the buttons below to copy the download link, copy the recommendation text, "
				"or instantly open a new email message in Outlook:"
			)
		)
		lblIntro.Wrap(600)
		mainSizer.Add(lblIntro, 0, wx.ALL, 10)

		downloadUrl = GITHUB_LATEST_DOWNLOAD
		repoUrl = GITHUB_REPO_URL

		shareBody = (
			"Cześć,\n\n"
			"Polecam Ci bezpłatny i w pełni dostępny dodatek dla czytnika ekranu NVDA:\n"
			"Outlook Speech Filter – wyciszanie gadatliwości mowy i nawigator wątków (dla Microsoft Outlook Classic).\n\n"
			"Kluczowe możliwości dodatku:\n"
			"• Błyskawiczne wyciszenie uciążliwych i rozwlekłych prefiksów mowy (NVDA od razu czyta od nadawcy, tematu i daty),\n"
			"• Zastąpienie zbędnych komunikatów dyskretnymi sygnałami dźwiękowymi (earconami) o spójnej logice,\n"
			"• Dedykowane okno konfiguracji i odsłuchu dźwięków (z informacją o częściowym zaznaczeniu),\n"
			"• Zaawansowany nawigator wątków mailowych (NVDA+F7 oraz NVDA+Alt+F7) z chronologicznym podglądem i czystym kopiowaniem (Alt+K),\n"
			"• Szybkie skoki między mailami w wątku (NVDA+Alt+Strzałki / Home / End) bez zamykania okna,\n"
			"• Odczyt wszystkich cech i statusów na żądanie (NVDA+Shift+Z) oraz obsługa zaznaczenia wielu maili naraz,\n"
			"• Inteligentne czyszczenie nagłówków w skrzynkach grupowanych według flag lub kategorii.\n\n"
			f"Bezpośredni link do pobrania najnowszej wersji instalatora (.nvda-addon):\n{downloadUrl}\n\n"
			f"Oficjalna strona projektu na GitHubie:\n{repoUrl}\n\n"
			"Instalacja:\n"
			"Pobierz plik 'outlookSpeechFilter.nvda-addon' i otwórz go klawiszem Enter w Eksploratorze Windows, "
			"a następnie potwierdź instalację w programie NVDA.\n\n"
			"Wymagania: klasyczny program pocztowy Microsoft Outlook Classic (wersje 2013-365) oraz NVDA 2024.1+."
		) if isPl else (
			"Hello,\n\n"
			"I would like to recommend a free and accessible add-on for the NVDA screen reader:\n"
			"Outlook Speech Filter – Mute Verbosity & Thread Navigator (for Microsoft Outlook Classic).\n\n"
			"Key Features:\n"
			"• Mutes verbose message status prefixes (NVDA immediately reads the sender, subject, and date),\n"
			"• Replaces status announcements with discrete, pleasant earcons,\n"
			"• Dedicated event sound customization dialog (with native partially-checked indicator),\n"
			"• Powerful conversation thread navigator (NVDA+F7 / NVDA+Alt+F7) with clean quote copying (Alt+K),\n"
			"• Quick thread navigation jumping shortcuts (NVDA+Alt+Arrows / Home / End),\n"
			"• On-demand comprehensive status report (NVDA+Shift+Z) with multi-selection support,\n"
			"• Automatic cleanup of verbose group headers in flagged or categorized mail views.\n\n"
			f"Direct download link for latest installer (.nvda-addon):\n{downloadUrl}\n\n"
			f"Official GitHub Project Page:\n{repoUrl}\n\n"
			"Installation:\n"
			"Download 'outlookSpeechFilter.nvda-addon', press Enter in Windows Explorer to open it, "
			"and confirm installation in NVDA.\n\n"
			"Compatibility: Microsoft Outlook Classic (2013-365) and NVDA 2024.1+."
		)

		self.txtShare = wx.TextCtrl(
			self,
			value=shareBody,
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_BESTWRAP,
			size=(-1, 230)
		)
		self.txtShare.SetName("Gotowy tekst wiadomości z łączami do pobrania" if isPl else "Share message text with download links")
		mainSizer.Add(self.txtShare, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		btnSizer = wx.WrapSizer(wx.HORIZONTAL)

		btnCopyLink = wx.Button(self, label="&Kopiuj link do pobrania" if isPl else "&Copy download link")
		def onCopyLink(evt):
			if copy_text_to_clipboard(downloadUrl):
				ui.message("Skopiowano bezpośredni link do pobrania dodatku do schowka." if isPl else "Download link copied to clipboard.")
			tones.beep(880, 50)
		btnCopyLink.Bind(wx.EVT_BUTTON, onCopyLink)
		btnSizer.Add(btnCopyLink, 0, wx.ALL, 4)

		btnCopyMsg = wx.Button(self, label="Kopiuj &całą treść" if isPl else "Copy &entire text")
		def onCopyMsg(evt):
			if copy_text_to_clipboard(shareBody):
				ui.message("Skopiowano kompletną treść polecającą do schowka." if isPl else "Recommendation text copied to clipboard.")
			tones.beep(880, 50)
		btnCopyMsg.Bind(wx.EVT_BUTTON, onCopyMsg)
		btnSizer.Add(btnCopyMsg, 0, wx.ALL, 4)

		btnEmail = wx.Button(self, label="&Wyślij e-mail z poleceniem..." if isPl else "&Send email recommendation...")
		def onEmail(evt):
			try:
				import urllib.parse
				subj = "Polecam bezpłatny dodatek: Outlook Speech Filter dla NVDA" if isPl else "Recommended free NVDA add-on: Outlook Speech Filter"
				mailto = f"mailto:?subject={urllib.parse.quote(subj)}&body={urllib.parse.quote(shareBody)}"
				os.startfile(mailto)
			except Exception as ex:
				log.debug(f"Share mailto error: {ex}")
		btnEmail.Bind(wx.EVT_BUTTON, onEmail)
		btnSizer.Add(btnEmail, 0, wx.ALL, 4)

		btnClose = wx.Button(self, wx.ID_CLOSE, label="&Zamknij" if isPl else "&Close")
		btnClose.Bind(wx.EVT_BUTTON, lambda evt: self.EndModal(wx.ID_CLOSE))
		btnSizer.Add(btnClose, 0, wx.ALL, 4)

		mainSizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 8)
		self.SetSizer(mainSizer)
		self.Layout()
		self.CenterOnScreen()

		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)
		wx.CallLater(50, lambda: self.txtShare.SetFocus())

	def onCharHook(self, evt):
		if evt.GetKeyCode() == wx.WXK_ESCAPE:
			self.EndModal(wx.ID_CLOSE)
		else:
			evt.Skip()


def showShareAddonDialog(parent=None):
	dlg = ShareAddonDialog(parent)
	gui.mainFrame.prePopup()
	try:
		dlg.ShowModal()
	finally:
		dlg.Destroy()
		gui.mainFrame.postPopup()


class SoundCustomizationDialog(wx.Dialog):
	"""
	W pełni dostępne okno dialogowe konfiguracji i odsłuchu sygnałów dźwiękowych (earconów).
	Pozwala wybrać motyw dźwięków, dostosować głośność, przetestować dźwięki
	oraz zaznaczyć/odznaczyć dowolne z 11 zdarzeń i natychmiast odsłuchać każdy dźwięk.
	"""
	def __init__(self, parent, current_states: dict, sound_theme: str = "associative", sound_volume: int = 80):
		isPl = (get_effective_language() == "pl")
		dlgTitle = _t("CUSTOMIZE_SOUNDS_DLG_TITLE")
		super().__init__(
			parent or gui.mainFrame,
			title=dlgTitle,
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
			size=(700, 620)
		)
		self.SetMinSize((640, 520))
		self.sound_theme = sound_theme if sound_theme in ("associative", "minimal", "instrumental") else "associative"
		try:
			self.sound_volume = int(sound_volume)
		except (ValueError, TypeError):
			self.sound_volume = 80
		self.states = dict(current_states)
		self.soundCBs = []

		mainSizer = wx.BoxSizer(wx.VERTICAL)

		lblIntro = wx.StaticText(
			self,
			label=_t("CUSTOMIZE_SOUNDS_INTRO")
		)
		lblIntro.Wrap(660)
		mainSizer.Add(lblIntro, 0, wx.ALL, 10)

		# 1. Wybór motywu dźwięków (przeniesiony na górę tego okna)
		themeBox = wx.BoxSizer(wx.HORIZONTAL)
		self.themeLabel = wx.StaticText(self, label=_t("SOUND_THEME_LABEL"))
		themeBox.Add(self.themeLabel, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

		self.themeKeys = ["associative", "minimal", "instrumental"]
		themeLabels = [
			_t("THEME_ASSOCIATIVE"),
			_t("THEME_MINIMAL"),
			_t("THEME_INSTRUMENTAL"),
		]
		self.soundThemeChoice = wx.Choice(self, choices=themeLabels)
		self.soundThemeChoice.SetName(_t("SOUND_THEME_LABEL"))
		themeIdx = 0
		if self.sound_theme in self.themeKeys:
			themeIdx = self.themeKeys.index(self.sound_theme)
		self.soundThemeChoice.SetSelection(themeIdx)
		self.soundThemeChoice.Bind(wx.EVT_CHOICE, self.onSoundThemeChanged)
		themeBox.Add(self.soundThemeChoice, flag=wx.ALIGN_CENTER_VERTICAL)
		mainSizer.Add(themeBox, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		# 2. Suwak głośności dźwięków oraz przycisk testowy (przeniesione na górę tego okna)
		volBox = wx.BoxSizer(wx.HORIZONTAL)
		self.volLabel = wx.StaticText(self, label=_t("VOL_LABEL"))
		volBox.Add(self.volLabel, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

		self.soundVolumeSlider = wx.Slider(
			self,
			value=self.sound_volume,
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

		self.volValueText = wx.StaticText(self, label=f"{self.sound_volume}%")
		volBox.Add(self.volValueText, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

		self.testVolBtn = wx.Button(self, label=_t("TEST_VOL_BTN"))
		self.testVolBtn.Bind(wx.EVT_BUTTON, self.onTestSoundVolume)
		volBox.Add(self.testVolBtn, flag=wx.ALIGN_CENTER_VERTICAL)
		mainSizer.Add(volBox, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		# Separator oddzielający konfigurację ogólną dźwięków od listy
		sepLine = wx.StaticLine(self)
		mainSizer.Add(sepLine, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		# 3. Przyciski zbiorcze
		bulkBox = wx.BoxSizer(wx.HORIZONTAL)
		self.selectAllBtn = wx.Button(self, label=_t("SELECT_ALL_SOUNDS"))
		self.selectAllBtn.Bind(wx.EVT_BUTTON, self.onSelectAll)
		bulkBox.Add(self.selectAllBtn, 0, wx.RIGHT, 10)

		self.deselectAllBtn = wx.Button(self, label=_t("DESELECT_ALL_SOUNDS"))
		self.deselectAllBtn.Bind(wx.EVT_BUTTON, self.onDeselectAll)
		bulkBox.Add(self.deselectAllBtn, 0)
		mainSizer.Add(bulkBox, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		# 4. Panel z przewijaniem dla poszczególnych dźwięków
		scrollPanel = wx.ScrolledWindow(self, style=wx.VSCROLL)
		scrollPanel.SetScrollRate(0, 20)
		scrollSizer = wx.BoxSizer(wx.VERTICAL)

		soundDefinitions = [
			("soundUnread", "unread", "SOUND_UNREAD", "PLAY_UNREAD", True),
			("soundFlags", "flag", "SOUND_FLAG", "PLAY_FLAG", False),
			("soundAttachments", "attachment", "SOUND_ATTACHMENT", "PLAY_ATTACHMENT", True),
			("soundInlineImages", "inline_image", "SOUND_INLINE_IMAGE", "PLAY_INLINE_IMAGE", True),
			("soundReplied", "replied", "SOUND_REPLIED", "PLAY_REPLIED", False),
			("soundForwarded", "forwarded", "SOUND_FORWARDED", "PLAY_FORWARDED", False),
			("soundImportanceHigh", "importance_high", "SOUND_IMPORTANCE_HIGH", "PLAY_IMPORTANCE_HIGH", True),
			("soundImportanceLow", "importance_low", "SOUND_IMPORTANCE_LOW", "PLAY_IMPORTANCE_LOW", False),
			("soundMeetingRequest", "meeting", "SOUND_MEETING", "PLAY_MEETING", True),
			("soundCategories", "category", "SOUND_CATEGORY", "PLAY_CATEGORY", False),
			("soundExternal", "external", "SOUND_EXTERNAL", "PLAY_EXTERNAL", False),
		]

		for confKey, soundName, cbKey, btnKey, defaultVal in soundDefinitions:
			rowBox = wx.BoxSizer(wx.HORIZONTAL)
			cb = wx.CheckBox(scrollPanel, label=_t(cbKey))
			val = self.states.get(confKey, defaultVal)
			cb.SetValue(_to_bool(val, default=defaultVal))
			cb.Bind(wx.EVT_CHECKBOX, self.onCheckBoxChanged)
			rowBox.Add(cb, 1, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 10)

			btn = wx.Button(scrollPanel, label=_t(btnKey))
			btn.Bind(wx.EVT_BUTTON, lambda evt, s=soundName: self._playSound(s))
			rowBox.Add(btn, 0, wx.ALIGN_CENTER_VERTICAL)

			scrollSizer.Add(rowBox, 0, wx.EXPAND | wx.BOTTOM, 6)
			self.soundCBs.append((confKey, cb, soundName))

		scrollPanel.SetSizer(scrollSizer)
		mainSizer.Add(scrollPanel, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		# Dolne przyciski OK / Anuluj
		btnSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.btnOK = wx.Button(self, wx.ID_OK, label="&Zastosuj i zamknij" if isPl else "&Apply and Close")
		self.btnOK.SetDefault()
		btnSizer.Add(self.btnOK, 0, wx.RIGHT, 10)

		self.btnCancel = wx.Button(self, wx.ID_CANCEL, label="&Anuluj" if isPl else "&Cancel")
		btnSizer.Add(self.btnCancel, 0)
		mainSizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

		self.SetSizer(mainSizer)
		self.Layout()
		self.CenterOnScreen()

		self.updateBulkButtonStates()
		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)
		if hasattr(self, "soundThemeChoice") and self.soundThemeChoice:
			wx.CallLater(50, lambda: self.soundThemeChoice.SetFocus())

	def _playSound(self, soundName):
		_stop_pending_sounds()
		vol = self.soundVolumeSlider.GetValue() if hasattr(self, "soundVolumeSlider") and self.soundVolumeSlider else self.sound_volume
		_play_status_sound(soundName, volume=vol, theme=self.sound_theme)

	def onSoundThemeChanged(self, evt):
		idx = self.soundThemeChoice.GetSelection()
		if 0 <= idx < len(self.themeKeys):
			self.sound_theme = self.themeKeys[idx]
			_stop_pending_sounds()
			vol = self.soundVolumeSlider.GetValue() if hasattr(self, "soundVolumeSlider") and self.soundVolumeSlider else self.sound_volume
			_play_status_sound("attachment", volume=vol, theme=self.sound_theme)

	def _updateSliderValue(self, newVal):
		newVal = max(0, min(100, newVal))
		self.soundVolumeSlider.SetValue(newVal)
		self.sound_volume = newVal
		self.volValueText.SetLabel(f"{newVal}%")
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
		if keyCode in (wx.WXK_PAGEUP, getattr(wx, "WXK_NUMPAD_PAGEUP", -1)):
			self._updateSliderValue(currentVal + 10)
			return
		elif keyCode in (wx.WXK_PAGEDOWN, getattr(wx, "WXK_NUMPAD_PAGEDOWN", -1)):
			self._updateSliderValue(currentVal - 10)
			return
		elif keyCode in (wx.WXK_UP, wx.WXK_RIGHT, getattr(wx, "WXK_NUMPAD_UP", -1), getattr(wx, "WXK_NUMPAD_RIGHT", -1)):
			self._updateSliderValue(currentVal + 1)
			return
		elif keyCode in (wx.WXK_DOWN, wx.WXK_LEFT, getattr(wx, "WXK_NUMPAD_DOWN", -1), getattr(wx, "WXK_NUMPAD_LEFT", -1)):
			self._updateSliderValue(currentVal - 1)
			return
		elif keyCode in (wx.WXK_HOME, getattr(wx, "WXK_NUMPAD_HOME", -1)):
			self._updateSliderValue(0)
			return
		elif keyCode in (wx.WXK_END, getattr(wx, "WXK_NUMPAD_END", -1)):
			self._updateSliderValue(100)
			return
		evt.Skip()

	def onSoundVolumeChanged(self, evt):
		val = self.soundVolumeSlider.GetValue()
		self.sound_volume = val
		self.volValueText.SetLabel(f"{val}%")

	def onTestSoundVolume(self, evt):
		_stop_pending_sounds()
		val = self.soundVolumeSlider.GetValue()
		if val <= 0:
			ui.message(_t("SOUNDS_VOL_MUTED"))
		else:
			_schedule_status_sounds_sequence(TEST_SOUNDS_SEQUENCE, volume=val, delay_ms=230, theme=self.sound_theme)

	def onCheckBoxChanged(self, evt):
		self.updateBulkButtonStates()

	def onSelectAll(self, evt):
		for _, cb, _ in self.soundCBs:
			cb.SetValue(True)
		self.updateBulkButtonStates()
		ui.message(_t("SOUNDS_ALL_SELECTED"))

	def onDeselectAll(self, evt):
		for _, cb, _ in self.soundCBs:
			cb.SetValue(False)
		self.updateBulkButtonStates()
		ui.message(_t("SOUNDS_ALL_DESELECTED"))

	def updateBulkButtonStates(self):
		total = len(self.soundCBs)
		selected = sum(1 for _, cb, _ in self.soundCBs if cb.GetValue())
		if hasattr(self, "selectAllBtn") and self.selectAllBtn:
			self.selectAllBtn.Enable(selected < total)
		if hasattr(self, "deselectAllBtn") and self.deselectAllBtn:
			self.deselectAllBtn.Enable(selected > 0)

	def getStates(self):
		return {confKey: cb.GetValue() for confKey, cb, _ in self.soundCBs}

	def getSoundTheme(self):
		idx = self.soundThemeChoice.GetSelection()
		if 0 <= idx < len(self.themeKeys):
			return self.themeKeys[idx]
		return self.sound_theme

	def getSoundVolume(self):
		if hasattr(self, "soundVolumeSlider") and self.soundVolumeSlider:
			return int(self.soundVolumeSlider.GetValue())
		return self.sound_volume

	def onCharHook(self, evt):
		key = evt.GetKeyCode()
		if key == wx.WXK_ESCAPE:
			self.EndModal(wx.ID_CANCEL)
		elif key == wx.WXK_F1:
			showContextHelp("sound_customization", parent=self)
		else:
			evt.Skip()


class OutlookSpeechFilterSettingsPanel(SettingsPanel):
	title = SettingsPanelTitleDescriptor()

	def makeSettings(self, settingsSizer):
		global _activeSettingsPanelRef
		_activeSettingsPanelRef = weakref.ref(self)
		try:
			self._makeSettingsInternal(settingsSizer)
		except Exception as e:
			log.error(f"outlookSpeechFilter: Błąd podczas budowania panelu ustawień: {e}", exc_info=True)
			errLabel = wx.StaticText(self, label=_t("SETTINGS_LOAD_ERROR", err=e))
			settingsSizer.Add(errLabel, flag=wx.ALL, border=15)

	def _makeSettingsInternal(self, settingsSizer):
		sHelper = gui.guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
		try:
			conf = config.conf[CONFIG_SECTION]
		except Exception:
			conf = {}

		# Załaduj aktualny motyw i głośność dźwięków do pamięci panelu
		currentTheme = conf.get("soundTheme", "associative")
		self.soundTheme = currentTheme if currentTheme in ("associative", "minimal", "instrumental") else "associative"
		self.initialSoundTheme = self.soundTheme

		currentVol = conf.get("soundVolume", 80)
		try:
			currentVol = int(currentVol)
		except (ValueError, TypeError):
			currentVol = 80
		self.soundVolume = currentVol
		self.initialSoundVolume = currentVol

		# 1. Filtrowanie mowy w Outlooku
		self.enabledCB = sHelper.addItem(
			wx.CheckBox(self, label=_t("ENABLE_FILTER"))
		)
		self.enabledCB.SetValue(_to_bool(conf.get("enabled", True), default=True))

		# 2. Nawigator wątków
		self.threadNavCB = sHelper.addItem(
			wx.CheckBox(self, label=_t("ENABLE_THREAD_NAV"))
		)
		self.threadNavCB.SetValue(_to_bool(conf.get("threadNavigationEnabled", True), default=True))

		# 3. Stan dźwięków (3-stanowe pole wyboru)
		self.soundStates = {
			"soundUnread": _to_bool(conf.get("soundUnread", conf.get("reportUnread", True)), default=True),
			"soundFlags": _to_bool(conf.get("soundFlags", conf.get("reportFlags", False)), default=False),
			"soundAttachments": _to_bool(conf.get("soundAttachments", conf.get("reportAttachments", True)), default=True),
			"soundInlineImages": _to_bool(conf.get("soundInlineImages", conf.get("reportInlineImages", True)), default=True),
			"soundReplied": _to_bool(conf.get("soundReplied", conf.get("reportReplied", False)), default=False),
			"soundForwarded": _to_bool(conf.get("soundForwarded", conf.get("reportForwarded", False)), default=False),
			"soundImportanceHigh": _to_bool(conf.get("soundImportanceHigh", conf.get("reportImportanceHigh", True)), default=True),
			"soundImportanceLow": _to_bool(conf.get("soundImportanceLow", conf.get("reportImportanceLow", False)), default=False),
			"soundMeetingRequest": _to_bool(conf.get("soundMeetingRequest", conf.get("reportMeetingRequest", True)), default=True),
			"soundCategories": _to_bool(conf.get("soundCategories", conf.get("reportCategories", False)), default=False),
			"soundExternal": _to_bool(conf.get("soundExternal", conf.get("reportExternal", False)), default=False),
		}
		self._lastActiveSoundStates = dict(self.soundStates)
		self.useSoundsEnabled = _to_bool(conf.get("useSounds", True), default=True)

		self.useSoundsCB = sHelper.addItem(
			wx.CheckBox(self, label=_t("PARENT_SOUNDS_LABEL"), style=wx.CHK_3STATE)
		)
		self.useSoundsCB.Bind(wx.EVT_CHECKBOX, self.onUseSoundsChanged)

		# Przycisk otwierający osobne okno konfiguracji dźwięków (Alt+D)
		# (w oknie tym znajduje się motyw dźwięków, regulacja głośności z testem oraz pełna lista sygnałów)
		soundBtnBox = wx.BoxSizer(wx.HORIZONTAL)
		self.btnCustomizeSounds = wx.Button(self, label=_t("CUSTOMIZE_SOUNDS_BTN"))
		self.btnCustomizeSounds.Bind(wx.EVT_BUTTON, self.onCustomizeSounds)
		soundBtnBox.Add(self.btnCustomizeSounds, flag=wx.ALIGN_CENTER_VERTICAL)
		sHelper.addItem(soundBtnBox)

		# 4. Sekcja aktualizacji (automatyczne sprawdzanie i przycisk na żądanie)
		settingsSizer.AddSpacer(6)
		self.chkAutoUpdates = sHelper.addItem(
			wx.CheckBox(self, label=_t("AUTO_UPDATES_CB"))
		)
		self.chkAutoUpdates.SetValue(_to_bool(conf.get("autoCheckUpdates", True), default=True))

		updBtnBox = wx.BoxSizer(wx.HORIZONTAL)
		self.checkUpdatesBtn = wx.Button(self, label=_t("CHECK_UPDATES_BTN"))
		self.checkUpdatesBtn.Bind(wx.EVT_BUTTON, self.onCheckUpdatesNow)
		updBtnBox.Add(self.checkUpdatesBtn, flag=wx.ALIGN_CENTER_VERTICAL)
		sHelper.addItem(updBtnBox)

		settingsSizer.AddSpacer(6)

		# 5. Wybór języka interfejsu dodatku (przeniesiony na dół okna)
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

		settingsSizer.AddSpacer(6)

		# 6. Przyciski narzędziowe: Pomoc i skróty (zgrupowane razem)
		toolsBox1 = wx.BoxSizer(wx.HORIZONTAL)
		self.contextHelpBtn = wx.Button(self, label=_t("CONTEXT_HELP_BTN"))
		self.contextHelpBtn.Bind(wx.EVT_BUTTON, lambda evt: showContextHelp("settings", parent=self))
		toolsBox1.Add(self.contextHelpBtn, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=8)

		self.helpBtn = wx.Button(self, label=_t("HELP_BTN"))
		self.helpBtn.Bind(wx.EVT_BUTTON, self.onOpenHelp)
		toolsBox1.Add(self.helpBtn, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=8)

		self.gesturesBtn = wx.Button(self, label=_t("GESTURES_BTN"))
		self.gesturesBtn.Bind(wx.EVT_BUTTON, self.onOpenInputGestures)
		toolsBox1.Add(self.gesturesBtn, flag=wx.ALIGN_CENTER_VERTICAL)
		sHelper.addItem(toolsBox1)

		# 7. Przyciski społecznościowe i przywracania fabrycznego
		toolsBox2 = wx.BoxSizer(wx.HORIZONTAL)
		self.shareBtn = wx.Button(self, label=_t("SHARE_BTN"))
		self.shareBtn.Bind(wx.EVT_BUTTON, lambda evt: showShareAddonDialog(self))
		toolsBox2.Add(self.shareBtn, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=8)

		self.contactBtn = wx.Button(self, label=_t("CONTACT_BTN"))
		self.contactBtn.Bind(wx.EVT_BUTTON, self.onContactAuthor)
		toolsBox2.Add(self.contactBtn, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=8)

		self.restoreDefaultsBtn = wx.Button(self, label=_t("RESTORE_DEFAULTS_BTN"))
		self.restoreDefaultsBtn.Bind(wx.EVT_BUTTON, self.onRestoreDefaults)
		toolsBox2.Add(self.restoreDefaultsBtn, flag=wx.ALIGN_CENTER_VERTICAL)
		sHelper.addItem(toolsBox2)

		# 8. Informacja o wersji i autorach
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

	def setCleanerEnabledState(self, state):
		if hasattr(self, "enabledCB") and self.enabledCB:
			self.enabledCB.SetValue(_to_bool(state, default=True))

	def setThreadNavEnabledState(self, state):
		if hasattr(self, "threadNavCB") and self.threadNavCB:
			self.threadNavCB.SetValue(_to_bool(state, default=True))

	def setUseSoundsState(self, state):
		if hasattr(self, "useSoundsCB") and self.useSoundsCB:
			self.useSoundsEnabled = _to_bool(state, default=True)
			self.updateParentSoundsLabel()

	def onCustomizeSounds(self, evt):
		dlg = SoundCustomizationDialog(
			self,
			self.soundStates,
			sound_theme=getattr(self, "soundTheme", "associative"),
			sound_volume=getattr(self, "soundVolume", 80)
		)
		gui.mainFrame.prePopup()
		try:
			if dlg.ShowModal() == wx.ID_OK:
				self.soundStates = dlg.getStates()
				self.soundTheme = dlg.getSoundTheme()
				self.soundVolume = dlg.getSoundVolume()
				if any(self.soundStates.values()):
					self._lastActiveSoundStates = dict(self.soundStates)
					self.useSoundsEnabled = True
				else:
					self.useSoundsEnabled = False
				self.updateParentSoundsLabel()
		finally:
			dlg.Destroy()
			gui.mainFrame.postPopup()

	def onUseSoundsChanged(self, evt):
		if not getattr(self, "useSoundsEnabled", True):
			self.useSoundsEnabled = True
			if not any(self.soundStates.values()):
				if hasattr(self, "_lastActiveSoundStates") and any(self._lastActiveSoundStates.values()):
					self.soundStates = dict(self._lastActiveSoundStates)
				else:
					for k in self.soundStates:
						self.soundStates[k] = (k in ("soundUnread", "soundAttachments", "soundInlineImages", "soundImportanceHigh", "soundMeetingRequest"))
			self.updateParentSoundsLabel()
		else:
			self.useSoundsEnabled = False
			self.updateParentSoundsLabel()

	def updateParentSoundsLabel(self):
		total = len(self.soundStates)
		selected = sum(1 for val in self.soundStates.values() if val)
		baseLabel = _t("PARENT_SOUNDS_LABEL")

		if not getattr(self, "useSoundsEnabled", True):
			self.useSoundsCB.Set3StateValue(wx.CHK_UNCHECKED)
			self.useSoundsCB.SetLabel(baseLabel)
			if hasattr(self, "btnCustomizeSounds"):
				self.btnCustomizeSounds.Enable(False)
		elif selected == total:
			self.useSoundsCB.Set3StateValue(wx.CHK_CHECKED)
			self.useSoundsCB.SetLabel(baseLabel)
			if hasattr(self, "btnCustomizeSounds"):
				self.btnCustomizeSounds.Enable(True)
		elif selected == 0:
			self.useSoundsCB.Set3StateValue(wx.CHK_UNCHECKED)
			self.useSoundsCB.SetLabel(baseLabel + _t("PARENT_ALL_DISABLED"))
			if hasattr(self, "btnCustomizeSounds"):
				self.btnCustomizeSounds.Enable(False)
		else:
			self.useSoundsCB.Set3StateValue(wx.CHK_UNDETERMINED)
			self.useSoundsCB.SetLabel(baseLabel + _t("PARENT_PARTIAL"))
			if hasattr(self, "btnCustomizeSounds"):
				self.btnCustomizeSounds.Enable(True)

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
		if hasattr(self, "btnCustomizeSounds") and self.btnCustomizeSounds:
			self.btnCustomizeSounds.SetLabel(_t("CUSTOMIZE_SOUNDS_BTN"))
		if hasattr(self, "chkAutoUpdates") and self.chkAutoUpdates:
			self.chkAutoUpdates.SetLabel(_t("AUTO_UPDATES_CB"))
		if hasattr(self, "checkUpdatesBtn") and self.checkUpdatesBtn:
			self.checkUpdatesBtn.SetLabel(_t("CHECK_UPDATES_BTN"))

		self.updateParentSoundsLabel()

		if hasattr(self, "contextHelpBtn") and self.contextHelpBtn:
			self.contextHelpBtn.SetLabel(_t("CONTEXT_HELP_BTN"))
		if hasattr(self, "helpBtn") and self.helpBtn:
			self.helpBtn.SetLabel(_t("HELP_BTN"))
		if hasattr(self, "shareBtn") and self.shareBtn:
			self.shareBtn.SetLabel(_t("SHARE_BTN"))
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
			# Podgląd języka na żywo (get_effective_language czyta wybór z panelu);
			# zapis do konfiguracji następuje dopiero w onSave, Anuluj przywraca poprzedni język.
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

	def onOpenHelp(self, evt):
		docFile = get_doc_path()
		if docFile and os.path.isfile(docFile):
			try:
				import urllib.parse
				url = f"file:///{urllib.parse.quote(docFile.replace(os.sep, '/'))}"
				import webbrowser
				webbrowser.open(url)
			except Exception:
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

		# Przywracane są wyłącznie kontrolki i stan panelu — zapis nastąpi w onSave (OK/Zastosuj),
		# a Anuluj pozostawia konfigurację bez zmian.
		# 0. Język
		if hasattr(self, "languageChoice"):
			self.languageChoice.SetSelection(0)

		# 0b. Motyw dźwięków i głośność
		self.soundTheme = "associative"
		self.soundVolume = 80

		# 1. Główne przełączniki
		self.enabledCB.SetValue(True)
		if hasattr(self, "threadNavCB") and self.threadNavCB:
			self.threadNavCB.SetValue(True)
		if hasattr(self, "chkAutoUpdates") and self.chkAutoUpdates:
			self.chkAutoUpdates.SetValue(True)

		# 2. Poszczególne dźwięki według oficjalnych ustawień domyślnych
		self.soundStates = {
			"soundUnread": True,
			"soundFlags": False,
			"soundAttachments": True,
			"soundInlineImages": True,
			"soundReplied": False,
			"soundForwarded": False,
			"soundImportanceHigh": True,
			"soundImportanceLow": False,
			"soundMeetingRequest": True,
			"soundCategories": False,
			"soundExternal": False,
		}
		self._lastActiveSoundStates = dict(self.soundStates)
		self.useSoundsEnabled = True

		self.updateParentSoundsLabel()
		self.updateUILabels()
		self.Layout()

		ui.message(_t("DEFAULTS_RESTORED"))

	def onDiscard(self):
		"""Anulowanie: panel niczego nie zapisał do konfiguracji przed onSave, więc wystarczy zwolnić referencję."""
		global _activeSettingsPanelRef
		_activeSettingsPanelRef = None

	def onSave(self):
		conf = config.conf[CONFIG_SECTION]
		try:
			if hasattr(self, "languageChoice") and hasattr(self, "langKeys"):
				idx = self.languageChoice.GetSelection()
				if 0 <= idx < len(self.langKeys):
					chosenLang = self.langKeys[idx]
					conf["language"] = chosenLang
					self.initialLanguage = chosenLang
			if hasattr(self, "soundTheme"):
				conf["soundTheme"] = self.soundTheme
				self.initialSoundTheme = self.soundTheme
			if hasattr(self, "enabledCB") and self.enabledCB:
				conf["enabled"] = bool(self.enabledCB.GetValue())
			if hasattr(self, "threadNavCB") and self.threadNavCB:
				conf["threadNavigationEnabled"] = bool(self.threadNavCB.GetValue())
			if hasattr(self, "useSoundsCB") and self.useSoundsCB:
				conf["useSounds"] = bool(getattr(self, "useSoundsEnabled", True) and any(self.soundStates.values()))
			if hasattr(self, "soundVolume"):
				conf["soundVolume"] = int(self.soundVolume)
				self.initialSoundVolume = int(self.soundVolume)

			for confKey, val in self.soundStates.items():
				conf[confKey] = bool(val)
				conf["report" + confKey[5:]] = bool(val)

			if hasattr(self, "chkAutoUpdates") and self.chkAutoUpdates:
				conf["autoCheckUpdates"] = bool(self.chkAutoUpdates.GetValue())

			# Synchronize legacy keys
			conf["soundForwardReply"] = bool(self.soundStates.get("soundReplied", False) or self.soundStates.get("soundForwarded", False))
			conf["soundImportance"] = bool(self.soundStates.get("soundImportanceHigh", True) or self.soundStates.get("soundImportanceLow", False))
			conf["reportForwardReply"] = bool(conf["soundForwardReply"])
			conf["reportImportance"] = bool(conf["soundImportance"])

			_invalidate_conf_snapshot()
			_refresh_script_descriptions()
			self.updateUILabels()
			try:
				config.conf.save()
			except Exception as e:
				log.warning(f"outlookSpeechFilter: Błąd zapisu konfiguracji do pliku w onSave: {e}")
		except Exception as e:
			log.warning(f"outlookSpeechFilter: Błąd podczas onSave: {e}")


OutlookMessageCleanerSettingsPanel = OutlookSpeechFilterSettingsPanel


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
		import comtypes.client.dynamic
		from comtypes.automation import IDispatch
		from ctypes import POINTER, byref

		# Tak samo jak NVDA (NVDAObjects.window.winword.WordDocument.WinwordWindowObject):
		# wskaźnik IDispatch z OBJID_NATIVEOM opakowany przez comtypes.client.dynamic.Dispatch
		pDispatch = POINTER(IDispatch)()
		hr = ctypes.windll.oleacc.AccessibleObjectFromWindow(
			wintypes.HWND(hwnd),
			ctypes.c_long(getattr(winUser, "OBJID_NATIVEOM", -16)),
			byref(IDispatch._iid_),
			byref(pDispatch),
		)
		if hr == 0 and pDispatch:
			windowObject = comtypes.client.dynamic.Dispatch(pDispatch)
			return getattr(windowObject, "Document", None)
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

		# Typy zgodne z 64-bitowym Windows: BOOL CALLBACK EnumChildProc(HWND, LPARAM)
		WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
		callback = WNDENUMPROC(enumChild)
		ctypes.windll.user32.EnumChildWindows(wintypes.HWND(parentHwnd), callback, wintypes.LPARAM(0))
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
	sel = _getDocumentSelection(doc)
	if sel is not None:
		try:
			return int(sel.Range.Start)
		except Exception:
			pass
	return 0


def _getDocumentSelection(doc):
	"""
	Zaznaczenie w oknie danego dokumentu (doc.ActiveWindow.Selection), a nie globalne
	Application.Selection, które może należeć do innego okna Worda/Outlooka.
	"""
	try:
		win = getattr(doc, "ActiveWindow", None)
		if win is not None:
			sel = getattr(win, "Selection", None)
			if sel is not None:
				return sel
	except Exception:
		pass
	try:
		app = getattr(doc, "Application", None)
		if app is not None:
			return getattr(app, "Selection", None)
	except Exception:
		pass
	return None


def selectWordRange(doc, char_start: int) -> bool:
	"""Przesuwa kursor i widok w Wordzie dokładnie do wskazanego znaku."""
	if not doc:
		return False
	try:
		target_rng = doc.Range(char_start, char_start)
		target_rng.Select()
		sel = _getDocumentSelection(doc)
		if sel is not None:
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


def resolveMessageBodyRangeInWord(doc, msgs: list, msg) -> tuple:
	"""
	Zwraca (start, koniec) treści wiadomości (bez nagłówka) w rzeczywistych pozycjach Worda,
	wyznaczonych tak samo jak przy skokach (resolveMessageRangeInWord), zamiast surowych
	przesunięć z tekstu doc.Content.Text (które dryfują przy tabelach, obrazkach i polach).
	"""
	start = resolveMessageRangeInWord(doc, msg)
	headerLen = 0
	if getattr(msg, "header_end", 0) > 0:
		headerLen = max(0, msg.header_end - msg.range_start)
	bodyStart = start + headerLen
	end = None
	try:
		ordered = sorted(msgs, key=lambda m: m.range_start)
		pos = ordered.index(msg)
		if pos + 1 < len(ordered):
			end = resolveMessageRangeInWord(doc, ordered[pos + 1])
	except Exception:
		end = None
	if end is None:
		try:
			end = int(doc.Content.End)
		except Exception:
			end = msg.range_end
	if end <= bodyStart:
		return (msg.header_end if msg.header_end > 0 else msg.range_start), msg.range_end
	return bodyStart, end


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
		self._filterAnnounceTimer = None

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
		# Bez dodatkowego ui.message — NVDA samo odczytuje zaznaczony element listy

	def onFilterChange(self, event):
		self.populateList(self.txtFilter.GetValue())
		# Po filtrowaniu (z opóźnieniem, by nie przerywać pisania) podaj liczbę wyników
		try:
			if self._filterAnnounceTimer is not None:
				self._filterAnnounceTimer.Stop()
		except Exception:
			pass
		self._filterAnnounceTimer = wx.CallLater(700, self._announceFilterResults)

	def _announceFilterResults(self):
		self._filterAnnounceTimer = None
		try:
			if self.txtFilter.HasFocus():
				ui.message(_t("THREAD_FILTER_RESULTS", count=len(self.filtered_messages)))
		except Exception:
			pass

	def onFilterKeyDown(self, event):
		if event.GetKeyCode() == wx.WXK_DOWN:
			self.listCtrl.SetFocus()
			return
		event.Skip()

	def onSelectionChange(self, event):
		event.Skip()

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
				start_c, end_c = resolveMessageBodyRangeInWord(self.doc, self.messages, target_m)
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

	# Skróty klawiszowe są zdefiniowane wyłącznie w dekoratorach @script(gestures=[...]).
	# „kb:NVDA+...” obejmuje zarówno Insert, jak i Caps Lock (zależnie od ustawień klawisza NVDA),
	# a NVDA normalizuje kolejność modyfikatorów, więc permutacje nie są potrzebne.

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
				if not _get_conf_snapshot()["threadNavigationEnabled"]:
					return None
				g_name = gesture.displayName.lower() if hasattr(gesture, "displayName") else ""
				if "alt" not in g_name and not self._is_outlook_message_active():
					return None
		return script

	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)

		config.conf.spec[CONFIG_SECTION] = confspec
		if CONFIG_SECTION not in config.conf:
			config.conf[CONFIG_SECTION] = {}
		targetSec = config.conf[CONFIG_SECTION]

		# Jednorazowa migracja ustawień ze starej sekcji (kopiowane są tylko brakujące klucze).
		# Nie przepisujemy pozostałych kluczy ani nie zapisujemy konfiguracji przy starcie —
		# typy wartości zapewnia specyfikacja confspec, a odczyty używają _to_bool/_to_int.
		if OLD_CONFIG_SECTION in config.conf:
			try:
				oldSec = config.conf[OLD_CONFIG_SECTION]
				migrated = False
				for k, v in oldSec.items():
					if k in confspec and k not in targetSec:
						targetSec[k] = v
						migrated = True
				if migrated:
					log.info("outlookSpeechFilter: Przeprowadzono migrację konfiguracji ze starej sekcji outlookMessageCleaner")
			except Exception as e:
				log.warning(f"outlookSpeechFilter: Błąd migracji ze starej konfiguracji: {e}")

		_invalidate_conf_snapshot()
		self._registeredConfigHandlers = []
		for extPointName in ("post_configProfileSwitch", "post_configReset"):
			extPoint = getattr(config, extPointName, None)
			if extPoint is None:
				continue
			try:
				extPoint.register(_invalidate_conf_snapshot)
				self._registeredConfigHandlers.append(extPoint)
			except Exception as e:
				log.debugWarning(f"outlookSpeechFilter: nie można zarejestrować {extPointName}: {e}")

		patchOutlookClasses()
		_refresh_script_descriptions()

		# Bezpieczna rejestracja panelu ustawień bez dublowania
		if OutlookSpeechFilterSettingsPanel not in gui.settingsDialogs.NVDASettingsDialog.categoryClasses:
			gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(OutlookSpeechFilterSettingsPanel)

		# Inicjalizacja modułu aktualizacji i opóźnione sprawdzenie w tle
		self._updater = AddonUpdaterEngine(
			addon_name=CONFIG_SECTION,
			addon_title=ADDON_NAME,
			current_version=ADDON_VERSION,
			addon_dir=ADDON_DIR,
			config_section=CONFIG_SECTION
		)
		self._updateTimer = wx.CallLater(15000, self._check_updates_auto)

	def _check_updates_auto(self):
		self._updateTimer = None
		conf = config.conf.get(CONFIG_SECTION, {})
		if not _to_bool(conf.get("autoCheckUpdates", True), default=True):
			return
		try:
			last_time = float(conf.get("lastUpdateCheckTime", 0.0) or 0.0)
		except (TypeError, ValueError):
			last_time = 0.0
		if time.time() - last_time < 86400:
			return
		self._updater.check_for_updates(manual=False)

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		# Nakładka tylko dla obiektów UIA Outlooka o klasach rzeczywistych wierszy wiadomości
		# lub nagłówków grup — nigdy dla obiektów IAccessible ani na podstawie samej roli/nazwy.
		if not builtinOutlook or _UIAObjectClass is None:
			return
		if not isinstance(obj, _UIAObjectClass):
			return
		appMod = getattr(obj, "appModule", None)
		if not appMod or getattr(appMod, "appName", "").lower() != "outlook":
			return
		cachedCls = _uia_class_name(obj)
		if cachedCls in MESSAGE_ROW_CLASSES or cachedCls in GROUP_HEADER_CLASSES:
			patchOutlookClasses()
			if CleanOutlookGridRow not in clsList:
				clsList.insert(0, CleanOutlookGridRow)

	def terminate(self, *args, **kwargs):
		super().terminate(*args, **kwargs)
		_stop_pending_sounds()
		global _original_UIAGridRow_get_name, _original_getReceivedMessageString, _patches_applied

		try:
			if getattr(self, "_updateTimer", None) is not None:
				self._updateTimer.Stop()
		except Exception:
			pass

		for extPoint in getattr(self, "_registeredConfigHandlers", []):
			try:
				extPoint.unregister(_invalidate_conf_snapshot)
			except Exception:
				pass

		# Przywrócenie oryginalnych funkcji tylko wtedy, gdy nadal są to nasze podmiany
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

		# Czyste usunięcie wszystkich instancji panelu z categoryClasses
		while OutlookSpeechFilterSettingsPanel in gui.settingsDialogs.NVDASettingsDialog.categoryClasses:
			try:
				gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(OutlookSpeechFilterSettingsPanel)
			except (ValueError, KeyError):
				break

	@script(
		description=_t("GESTURE_DESC_TOGGLE_FILTER"),
		gestures=[
			"kb:NVDA+control+shift+o",
		],
	)
	def script_toggleOutlookCleaner(self, gesture):
		conf = config.conf[CONFIG_SECTION]
		newState = not _to_bool(conf.get("enabled", True), default=True)
		conf["enabled"] = newState
		_invalidate_conf_snapshot()
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
		description=_t("GESTURE_DESC_TOGGLE_SOUNDS"),
		gestures=[
			"kb:NVDA+control+shift+s",
		],
	)
	def script_toggleSounds(self, gesture):
		conf = config.conf[CONFIG_SECTION]
		newState = not _to_bool(conf.get("useSounds", True), default=True)
		conf["useSounds"] = newState
		_invalidate_conf_snapshot()
		if not newState:
			_stop_pending_sounds()
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
		description=_t("GESTURE_DESC_REPORT_STATUSES"),
		gestures=[
			"kb:NVDA+control+shift+z",
			"kb:NVDA+shift+z",
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
		nativeOm = getattr(appMod, "nativeOm", None)
		if not nativeOm and getattr(focus, "parent", None):
			p = focus.parent
			if getattr(p, "appModule", None):
				nativeOm = getattr(p.appModule, "nativeOm", None)

		selObj = None
		if nativeOm:
			try:
				exp = None
				for mName in ("activeExplorer", "ActiveExplorer"):
					f = getattr(nativeOm, mName, None)
					if f:
						try:
							exp = f() if callable(f) else f
							if exp:
								break
						except Exception:
							pass
				if exp:
					for sName in ("selection", "Selection"):
						s = getattr(exp, sName, None)
						if s is not None:
							selObj = s() if callable(s) else s
							if selObj is not None:
								break
			except Exception:
				pass

		selCount = 0
		if selObj:
			for cName in ("count", "Count"):
				cVal = getattr(selObj, cName, None)
				if cVal is not None:
					try:
						selCount = int(cVal() if callable(cVal) else cVal)
						break
					except Exception:
						pass
			if selCount == 0:
				try:
					selCount = len(selObj)
				except Exception:
					pass

		# Sprawdź UIA Selection Pattern na liście wiadomości (fallback)
		if selCount <= 1 and focus:
			try:
				parentList = focus
				for _ in range(3):
					if not parentList:
						break
					role = getattr(parentList, "role", None)
					if role in (controlTypes.Role.LIST, controlTypes.Role.TABLE, controlTypes.Role.TREEVIEW):
						break
					parentList = getattr(parentList, "parent", None)
				if parentList and hasattr(parentList, "UIASelectionPattern"):
					selItems = parentList.UIASelectionPattern.GetSelection()
					if selItems and len(selItems) > 1:
						selCount = len(selItems)
			except Exception:
				pass

		if selCount > 1:
			unreadCount = 0
			attCount = 0
			flaggedCount = 0
			highCount = 0
			limit = min(selCount, 100)
			for i in range(1, limit + 1):
				try:
					it = None
					for itName in ("item", "Item"):
						itFn = getattr(selObj, itName, None)
						if itFn:
							try:
								it = itFn(i)
								break
							except Exception:
								pass
					if not it and selObj:
						try:
							it = selObj[i - 1]
						except Exception:
							pass

					if it:
						isUnread = bool(getattr(it, "unread", getattr(it, "UnRead", False)))
						if isUnread:
							unreadCount += 1
						atts = getattr(it, "attachments", getattr(it, "Attachments", None))
						attCnt = 0
						if atts:
							for aCount in ("count", "Count"):
								v = getattr(atts, aCount, None)
								if v is not None:
									attCnt = int(v() if callable(v) else v)
									break
						if attCnt > 0:
							attCount += 1
						isTask = bool(getattr(it, "isMarkedAsTask", getattr(it, "IsMarkedAsTask", False)))
						flagSt = int(getattr(it, "flagStatus", getattr(it, "FlagStatus", 0)))
						if isTask or flagSt == 2:
							flaggedCount += 1
						imp = int(getattr(it, "importance", getattr(it, "Importance", 1)))
						if imp == 2:
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

		# 1. Statusy wiersza z ostatniego przetwarzania (lista krotek (klucz_statusu, tekst))
		rowItems = None
		for cand in (focus, getattr(focus, "parent", None)):
			if isinstance(cand, CleanOutlookGridRow):
				try:
					rowItems = _get_row_info(cand)[1]
				except Exception:
					rowItems = None
				break
		if not rowItems and _last_focused_statuses:
			rowItems = _last_focused_statuses

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

		statusItems = _get_message_statuses_report(
			appMod,
			selection,
			windowThreadID=getattr(focus, "windowThreadID", None),
			uiaObj=focus,
		)

		# 3. Uzupełnij statusami wiersza (łączenie po kluczu statusu, bez duplikatów)
		merged = {}
		for key, text in statusItems:
			merged.setdefault(key, text)
		if rowItems:
			for key, text in rowItems:
				merged.setdefault(key, text)

		# 4. Kolejność ściśle według CANONICAL_STATUS_ORDER (dokładne dopasowanie klucza)
		statuses = [merged[k] for k in sorted(merged, key=_status_order_key)]

		if statuses:
			ui.message(", ".join(statuses))
		elif selection or rowItems is not None:
			ui.message(_t("STATUS_NO_ADDITIONAL"))
		else:
			ui.message(_t("STATUS_NO_SELECTION"))

	@script(
		description=_t("GESTURE_DESC_CONTEXT_HELP"),
		gestures=[
			"kb:NVDA+control+shift+h",
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

		if not doc:
			ui.message(_t("THREAD_CANNOT_NAVIGATE"))
			return False

		exact_pos = resolveMessageRangeInWord(doc, target_msg)
		if not selectWordRange(doc, exact_pos):
			# Sukces ogłaszamy tylko wtedy, gdy kursor faktycznie został przeniesiony
			ui.message(_t("THREAD_NAV_ERROR"))
			return False

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
		return True

	@script(
		description=_t("GESTURE_DESC_TOGGLE_THREAD_NAV"),
		gestures=[
			"kb:NVDA+control+shift+t",
		],
	)
	def script_toggleThreadNavigation(self, gesture):
		conf = config.conf[CONFIG_SECTION]
		newState = not _to_bool(conf.get("threadNavigationEnabled", True), default=True)
		conf["threadNavigationEnabled"] = newState
		_invalidate_conf_snapshot()
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
		description=_t("GESTURE_DESC_THREAD_STRUCTURE"),
		gestures=[
			"kb:NVDA+f7",
			"kb:NVDA+alt+f7",
		],
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
		description=_t("GESTURE_DESC_THREAD_NEXT"),
		gestures=[
			"kb:NVDA+alt+downArrow",
		],
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
		description=_t("GESTURE_DESC_THREAD_PREV"),
		gestures=[
			"kb:NVDA+alt+upArrow",
		],
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
		description=_t("GESTURE_DESC_THREAD_FIRST"),
		gestures=[
			"kb:NVDA+alt+home",
		],
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
		description=_t("GESTURE_DESC_THREAD_LAST"),
		gestures=[
			"kb:NVDA+alt+end",
		],
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


# Opisy skryptów w oknie „Zdarzenia wejścia” — klucze tłumaczeń _t, odświeżane po zmianie języka dodatku
SCRIPT_DESCRIPTION_KEYS = {
	"script_toggleOutlookCleaner": "GESTURE_DESC_TOGGLE_FILTER",
	"script_toggleSounds": "GESTURE_DESC_TOGGLE_SOUNDS",
	"script_reportAllMessageStatuses": "GESTURE_DESC_REPORT_STATUSES",
	"script_showContextHelp": "GESTURE_DESC_CONTEXT_HELP",
	"script_toggleThreadNavigation": "GESTURE_DESC_TOGGLE_THREAD_NAV",
	"script_showEmailThreadStructure": "GESTURE_DESC_THREAD_STRUCTURE",
	"script_nextThreadMessage": "GESTURE_DESC_THREAD_NEXT",
	"script_prevThreadMessage": "GESTURE_DESC_THREAD_PREV",
	"script_firstThreadMessage": "GESTURE_DESC_THREAD_FIRST",
	"script_lastThreadMessage": "GESTURE_DESC_THREAD_LAST",
}


def _refresh_script_descriptions():
	"""Ustawia opisy skryptów (__doc__) zgodnie z aktualnym językiem dodatku."""
	for scriptName, key in SCRIPT_DESCRIPTION_KEYS.items():
		func = GlobalPlugin.__dict__.get(scriptName)
		if func is None:
			continue
		try:
			func.__doc__ = _t(key)
		except Exception:
			pass

