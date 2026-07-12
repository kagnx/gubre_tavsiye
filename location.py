"""Coğrafi konum tespiti - Ucretsiz geocoder servisi."""
import socket


def internet_var_mi():
    """Internet baglantisi olup olmadigini kontrol eder."""
    try:
        socket.create_connection(("8.8.8.8", 53), timeout=3)
        return True
    except (OSError, socket.timeout):
        return False


def konum_al():
    """Mevcut konumu alir (ucretsiz IP tabanli).

    Returns:
        (enlem, boylam, adres) veya (None, None, hata_mesaji)
    """
    if not internet_var_mi():
        return None, None, "Internet baglantisi yok"

    # 1. Deneme: ip-api.com
    try:
        import urllib.request
        import json
        url = "http://ip-api.com/json/?fields=lat,lon,city,country"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8) as response:
            data = json.loads(response.read().decode())
            if data.get("status") == "success":
                enlem = data["lat"]
                boylam = data["lon"]
                sehir = data.get("city", "")
                ulke = data.get("country", "")
                return enlem, boylam, f"{sehir}, {ulke}"
    except Exception:
        pass

    # 2. Deneme: ipinfo.io
    try:
        import urllib.request
        import json
        url = "https://ipinfo.io/json"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8) as response:
            data = json.loads(response.read().decode())
            loc = data.get("loc", "")
            if loc and "," in loc:
                parts = loc.split(",")
                enlem = float(parts[0])
                boylam = float(parts[1])
                sehir = data.get("city", "")
                ulke = data.get("country", "")
                return enlem, boylam, f"{sehir}, {ulke}"
    except Exception:
        pass

    # 3. Deneme: ipwhois
    try:
        import urllib.request
        import json
        url = "https://ipwho.is/"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8) as response:
            data = json.loads(response.read().decode())
            if data.get("success"):
                enlem = data["latitude"]
                boylam = data["longitude"]
                sehir = data.get("city", "")
                ulke = data.get("country", "")
                return enlem, boylam, f"{sehir}, {ulke}"
    except Exception:
        pass

    return None, None, "Konum alinamadi - internet kontrol edin"


# Turkiye bolgelerinin koordinat araliklari (yaklasik)
BOLGE_KOORDINATLARI = {
    "Trakya": {"enlem": (40.0, 42.1), "boylam": (26.0, 29.6)},
    "Marmara": {"enlem": (39.5, 41.0), "boylam": (28.0, 30.5)},
    "Karadeniz": {"enlem": (40.0, 41.5), "boylam": (32.0, 41.5)},
    "Orta Anadolu": {"enlem": (38.5, 40.5), "boylam": (30.0, 36.0)},
    "Akdeniz": {"enlem": (36.0, 37.5), "boylam": (29.0, 36.5)},
    "Adalar": {"enlem": (36.0, 37.5), "boylam": (26.0, 28.5)},
    "Dogu Anadolu": {"enlem": (38.0, 41.0), "boylam": (40.0, 45.0)},
    "Guneydogu Anadolu": {"enlem": (36.5, 38.5), "boylam": (37.0, 44.5)},
    "Goller": {"enlem": (37.5, 39.0), "boylam": (29.5, 32.5)},
}


def bolgeyi_tahmin_et(enlem, boylam):
    """Verilen koordinatlara gore bolgeyi tahmin eder."""
    if enlem is None or boylam is None:
        return None

    for bolge, sinirlar in BOLGE_KOORDINATLARI.items():
        if (sinirlar["enlem"][0] <= enlem <= sinirlar["enlem"][1] and
                sinirlar["boylam"][0] <= boylam <= sinirlar["boylam"][1]):
            return bolge

    return None


def konumdan_bolge_bul():
    """Konum bilgisinden bolgeyi otomatik bulur.

    Returns:
        (bolge_adi, konum_bilgisi) veya (None, hata_mesaji)
    """
    enlem, boylam, adres = konum_al()
    if enlem is None:
        return None, adres  # hata mesaji

    bolge = bolgeyi_tahmin_et(enlem, boylam)
    if bolge:
        return bolge, f"{adres} -> {bolge}"
    else:
        return None, f"{adres} - Bolge eslesmedi"
