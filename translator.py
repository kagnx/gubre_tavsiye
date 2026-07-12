"""Dil destegi - Ceviri yonetimi."""
import json
import os
from config import kaynak_yolu

CEVIRILER = {}
MEVCUT_DIL = "tr"


def cevirileri_yukle():
    """Tum ceviri dosyalarini yukler."""
    global CEVIRILER
    ceviri_klasoru = kaynak_yolu("translations")
    for dil in ["tr", "en"]:
        dosya = os.path.join(ceviri_klasoru, f"{dil}.json")
        if os.path.exists(dosya):
            try:
                with open(dosya, "r", encoding="utf-8") as f:
                    CEVIRILER[dil] = json.load(f)
            except Exception:
                CEVIRILER[dil] = {}
        else:
            CEVIRILER[dil] = {}


def dil_ayarla(dil_kodu):
    """Aktif dili ayarlar."""
    global MEVCUT_DIL
    if dil_kodu in CEVIRILER:
        MEVCUT_DIL = dil_kodu


def cevir(anahtar, varsayilan=None):
    """Verilen anahtar icin ceviriyi dondurur."""
    if MEVCUT_DIL in CEVIRILER:
        return CEVIRILER[MEVCUT_DIL].get(anahtar, varsayilan or anahtar)
    return varsayilan or anahtar


def mevcut_dil():
    """Aktif dil kodunu dondurur."""
    return MEVCUT_DIL


# Baslat
cevirileri_yukle()
