# -*- coding: utf-8 -*-
# installTasks.py for outlookSpeechFilter
# Obsługa instalacji, aktualizacji i automatycznego usuwania starej wersji (outlookMessageCleaner).

from __future__ import annotations
import os
import shutil
import addonHandler
import globalVars
from logHandler import log

LEGACY_ADDON_NAME = "outlookMessageCleaner"


def _to_bool(val, default=True):
	"""Bezpieczna konwersja wartości na typ bool."""
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
	"""Bezpieczna konwersja wartości na int w zadanym zakresie."""
	try:
		iv = int(val)
		return max(min_val, min(max_val, iv))
	except (ValueError, TypeError):
		return default


def onInstall():
	"""
	Wywoływane automatycznie przez NVDA podczas instalacji dodatku outlookSpeechFilter.
	Wykrywa i usuwa starą wersję o nazwie 'outlookMessageCleaner', dzięki czemu
	po instalacji nie występują zdublowane dodatki ani zdublowane opcje w ustawieniach.
	"""
	log.info("outlookSpeechFilter installTasks: rozpoczęcie zadania instalacyjnego (onInstall)")

	# 1. Zgłoszenie żądania usunięcia starego dodatku do menedżera dodatków NVDA (addonHandler)
	try:
		for addon in addonHandler.getAvailableAddons():
			if addon.name == LEGACY_ADDON_NAME:
				log.info(f"outlookSpeechFilter installTasks: Wykryto stary dodatek '{addon.name}', żądanie usunięcia...")
				try:
					addon.requestRemove()
					log.info(f"outlookSpeechFilter installTasks: Pomyślnie zgłoszono requestRemove() dla '{addon.name}'. Usunięcie nastąpi bezpiecznie przy restarcie NVDA.")
				except Exception as e:
					log.warning(f"outlookSpeechFilter installTasks: addon.requestRemove() nie powiodło się: {e}")
	except Exception as e:
		log.warning(f"outlookSpeechFilter installTasks: Błąd podczas przeszukiwania dostępnych dodatków: {e}")

	# 2. Migracja ustawień konfiguracyjnych ze starej sekcji do nowej (z zachowaniem kopii zapasowej)
	try:
		import config
		if "outlookMessageCleaner" in config.conf:
			oldSec = config.conf["outlookMessageCleaner"]
			if "outlookSpeechFilter" not in config.conf:
				config.conf["outlookSpeechFilter"] = {}
			newSec = config.conf["outlookSpeechFilter"]
			for k, v in oldSec.items():
				if k not in newSec:
					if k == "soundVolume":
						newSec[k] = _to_int(v, 80)
					elif k == "language":
						newSec[k] = str(v)
					else:
						newSec[k] = _to_bool(v)
			log.info("outlookSpeechFilter installTasks: Pomyślnie przeniesiono konfigurację ze starej sekcji outlookMessageCleaner.")
	except Exception as e:
		log.warning(f"outlookSpeechFilter installTasks: Błąd podczas migracji konfiguracji: {e}")

	# 3. Język dodatku: NIE nadpisujemy ustawienia „auto” — wartość „auto” oznacza dynamiczne
	# dopasowanie do języka NVDA (languageHandler.getLanguage()) przy każdym uruchomieniu.
	# Naprawiamy jedynie wartości zapisane jako tekst (np. "True"/"80"), zapisując tylko przy zmianie.
	try:
		import config
		if "outlookSpeechFilter" in config.conf:
			sec = config.conf["outlookSpeechFilter"]
			changed = False
			for k, v in list(sec.items()):
				if not isinstance(v, str):
					continue
				if k == "soundVolume":
					newVal = _to_int(v, 80)
				elif v.strip().lower() in ("true", "false"):
					newVal = _to_bool(v)
				else:
					continue
				if newVal != v:
					sec[k] = newVal
					changed = True
			if changed:
				try:
					config.conf.save()
					log.info("outlookSpeechFilter installTasks: Konfiguracja pomyślnie zapisana.")
				except Exception as e:
					log.warning(f"outlookSpeechFilter installTasks: Błąd zapisu konfiguracji: {e}")
	except Exception as e:
		log.warning(f"outlookSpeechFilter installTasks: Nie udało się sprawdzić konfiguracji: {e}")

	log.info("outlookSpeechFilter installTasks: zakończenie zadania onInstall")

