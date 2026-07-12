"""Uygulama ayarlari ve sabitleri."""
import os
import sys

SURUM = "3.0"
SURUM_TARIHI = "Temmuz 2026"
UYGULAMA_ADI = "Gubre Tavsiyesi"
YAZILIMCI = "Oguz Kaan FIRAT"

BEKLENEN_UZUNLUK = 19
N_ARALIK_SAYISI = 4
P_ARALIK_SAYISI = 10
K_ARALIK_SAYISI = 5

BAKLAGIL_BITKILERI = [
    "Yonca", "Nohut", "Mercimek", "Fiğ", "Korunga",
    "K. Fasulye", "Soya", "Fıstık",
]

KISAYOLLAR = {
    "Ctrl+H": "hesapla",
    "Ctrl+P": "pdf_kaydet",
    "Ctrl+E": "eposta_gonder",
    "Ctrl+O": "dosya_yukle",
    "Ctrl+G": "gecmis_goster",
    "Ctrl+L": "konum_bul",
    "Ctrl+T": "tema_degistir",
    "Ctrl+D": "dil_degistir",
}

# Varsayilan ayarlar
VARSAYILAN_AYARLAR = {
    "tema": "acik",
    "dil": "tr",
    "smtp_sunucu": "smtp.gmail.com",
    "smtp_port": 587,
    "smtp_kullanici": "",
    "smtp_sifre": "",
    "son_bolge": "",
    "son_bitki": "",
    "son_tarim": "",
    "otomatik_kaydet": True,
}

def kaynak_yolu(dosya_adi):
    """PyInstaller icin resource yolu cozumu."""
    if getattr(sys, 'frozen', False):
        base_yol = sys._MEIPASS
    else:
        base_yol = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_yol, dosya_adi)

def veri_klasoru():
    """Veri klasoru yolunu dondurur."""
    if getattr(sys, 'frozen', False):
        base_yol = os.path.dirname(sys.executable)
    else:
        base_yol = os.path.dirname(os.path.abspath(__file__))
    return base_yol

def ayarlari_yukle():
    """Kayitli ayarlari yukler."""
    import json
    ayar_dosyasi = os.path.join(veri_klasoru(), "ayarlar.json")
    ayarlar = dict(VARSAYILAN_AYARLAR)
    if os.path.exists(ayar_dosyasi):
        try:
            with open(ayar_dosyasi, "r", encoding="utf-8") as f:
                kayitli = json.load(f)
                ayarlar.update(kayitli)
        except Exception:
            pass
    return ayarlar

def ayarlari_kaydet(ayarlar):
    """Ayarlari dosyaya kaydeder."""
    import json
    ayar_dosyasi = os.path.join(veri_klasoru(), "ayarlar.json")
    try:
        with open(ayar_dosyasi, "w", encoding="utf-8") as f:
            json.dump(ayarlar, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
