"""
TARIM-GUBRE v3.0 - Turkiye Profesyonel Gubre Tavsiye Sistemi
==============================================================
Yeni ozellikler v3.0:
  - Tema destegi (Acik/Karanlik)
  - Dil secenegi (Turkce/Ingilizce)
  - Kisayol tuslari
  - Otomatik kaydetme
  - Excel/CSV yukleme
  - Gecmis kayitlari (SQLite)
  - NPK grafikleri (matplotlib)
  - Karsilastirmali analiz
  - E-posta gonderme (SMTP)
  - Coğrafi konum ile bolge tespiti
  - Cevrimdisi kontrol
  - Baski ciktisi
"""

import sys
import os
import ctypes
import json
from datetime import datetime

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QComboBox, QLabel, QDoubleSpinBox, QPushButton, QTableWidget,
    QTableWidgetItem, QGroupBox, QHeaderView, QMessageBox, QTabWidget,
    QCheckBox, QFileDialog, QFrame, QSplitter, QSizePolicy,
    QMenuBar, QMenu, QStatusBar, QDialog, QLineEdit, QTextEdit,
    QGridLayout, QProgressBar, QSpinBox
)
from PyQt6.QtCore import Qt, QTimer, QSettings
from PyQt6.QtGui import QColor, QFont, QPalette, QAction, QIcon, QKeySequence

# PDF icin reportlab importlari
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm, cm
from reportlab.lib.colors import HexColor, white, black
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Yerel moduller
from config import (
    SURUM, SURUM_TARIHI, UYGULAMA_ADI, YAZILIMCI,
    BEKLENEN_UZUNLUK, N_ARALIK_SAYISI, P_ARALIK_SAYISI, K_ARALIK_SAYISI,
    BAKLAGIL_BITKILERI, kaynak_yolu, veri_klasoru,
    ayarlari_yukle, ayarlari_kaydet, VARSAYILAN_AYARLAR
)
import database
import translator
from translator import cevir
from data_loader import dosya_yukle
from email_sender import eposta_gonder, smtp_test_et
from location import konumdan_bolge_bul, internet_var_mi
from charts import NPKPastalGrafik, NPKCubukGrafik

# =============================================================================
# 1. VERI TABANI
# =============================================================================

def veri_tabani_olustur():
    ham_veri = {
        "Trakya": {
            "Bugday": {"Sulu": [17,16,15,12,13,11,9,8,7,5,4,3,0,0,12,9,6,4,0], "Kuru": [14,13,12,10,10,9,8,7,6,5,3,2,0,0,9,7,4,0,0]},
            "Arpa": {"Kuru": [10,9,8,7,9,9,8,7,6,5,3,2,0,0,0,0,0,0,0]},
            "Misir": {"Sulu": [19,18,16,14,16,15,12,10,8,6,4,3,0,0,14,10,6,4,0], "Kuru": [14,13,12,10,10,9,7,6,5,4,3,2,0,0,0,0,0,0,0]},
            "Aycicegi": {"Sulu": [18,14,12,10,13,12,11,10,9,7,5,4,0,0,0,0,0,0,0], "Kuru": [10,9,8,7,10,9,8,7,6,5,4,3,0,0,12,10,8,5,0]},
            "Seker Pancari": {"Sulu": [15,14,12,10,12,11,10,9,8,7,5,4,0,0,14,12,10,8,0], "Kuru": [11,10,9,8,9,8,7,6,5,4,4,3,0,0,12,10,8,6,0]},
            "Patates": {"Sulu": [17,16,14,12,12,11,10,9,8,7,5,4,0,0,14,12,10,8,0]},
            "Kabak": {"Sulu": [10,9,8,7,10,9,8,7,6,5,0,3,0,0,0,0,0,0,0]},
            "Tutun": {"Kuru": [5,4,3,3,6,5,4,4,3,3,0,0,0,0,0,0,0,0,0]},
            "Zeytin": {"Kuru": [10,9,8,7,10,9,8,7,6,5,3,3,0,0,0,0,0,0,0]},
            "Nohut": {"Kuru": [5,4,3,3,10,9,8,7,6,5,4,3,0,0,0,0,0,0,0]},
            "Mercimek": {"Kuru": [5,4,3,3,10,9,8,7,6,5,4,3,0,0,0,0,0,0,0]},
            "Kekik": {"Kuru": [4,4,3,3,6,5,4,4,3,3,0,0,0,0,0,0,0,0,0]},
            "Kimyon": {"Kuru": [8,7,6,5,8,7,6,5,4,3,0,0,0,0,0,0,0,0,0]},
            "Bag": {"Sulu": [14,12,11,10,11,10,9,8,7,6,5,4,3,0,14,12,10,7,0], "Kuru": [10,9,8,7,8,7,6,5,4,3,3,0,0,0,12,10,8,6,0]},
            "Domates (Sera)": {"Sulu": [18,16,14,12,12,11,10,9,8,7,6,5,0,0,14,12,8,6,0]},
            "Salatalik (Sera)": {"Sulu": [16,14,12,10,10,9,8,7,6,5,4,3,0,0,12,10,8,6,0]},
            "Elma": {"Sulu": [12,10,8,7,11,10,9,8,7,6,5,4,3,0,10,8,6,4,0]},
            "Armut": {"Sulu": [12,10,8,7,11,10,9,8,7,6,5,4,3,0,10,8,6,4,0]},
            "Cilek": {"Sulu": [12,10,9,7,9,8,7,6,5,4,3,0,0,0,10,8,6,4,0]},
            "Antep Fistigi": {"Kuru": [10,8,7,5,9,8,7,6,5,4,4,3,0,0,10,8,6,4,0]},
        },
        "Marmara": {
            "Bugday": {"Sulu": [17,16,15,12,11,10,8,7,6,5,4,3,0,0,13,10,7,5,0], "Kuru": [10,9,8,7,9,8,7,6,5,4,3,0,0,0,10,8,6,4,0]},
            "Misir": {"Sulu": [21,20,19,17,12,11,9,8,7,6,5,4,3,0,16,14,12,10,7], "Kuru": [14,13,12,10,10,9,7,6,5,4,3,0,0,0,0,0,0,0,0]},
            "Aycicegi": {"Sulu": [18,12,11,9,13,12,11,9,7,5,3,3,0,0,12,10,7,5,0], "Kuru": [10,9,8,7,10,9,8,7,6,4,3,3,0,0,0,0,0,0,0]},
            "Findik": {"Kuru": [20,18,16,14,15,14,13,11,9,7,5,3,0,0,0,0,0,0,0]},
            "Zeytin": {"Kuru": [10,9,8,7,10,9,8,7,6,5,4,3,0,0,8,6,4,0,0]},
            "Seker Pancari": {"Sulu": [18,17,16,14,13,11,10,9,7,6,5,4,3,0,16,12,8,6,0]},
            "Sogan": {"Sulu": [14,13,12,10,10,9,8,7,6,5,4,3,0,0,0,0,0,0,0], "Kuru": [9,8,7,6,8,7,6,5,5,4,4,3,0,0,0,0,0,0,0]},
            "Sarmsak": {"Sulu": [15,14,13,11,10,9,8,7,6,5,4,3,0,0,0,0,0,0,0]},
            "Bag": {"Sulu": [15,13,12,10,12,10,9,8,7,6,5,4,0,0,15,12,9,6,0], "Kuru": [10,9,8,7,10,9,7,6,5,4,3,0,0,0,12,10,7,5,0]},
            "Seftali": {"Sulu": [14,12,10,8,12,11,10,9,8,7,6,5,0,0,12,10,8,6,0]},
            "Kiraz": {"Sulu": [12,10,8,6,11,10,9,8,7,6,5,4,0,0,10,8,6,4,0]},
            "Tutun (Kalite)": {"Kuru": [6,5,4,3,7,6,5,4,4,3,3,0,0,0,0,8,6,5,4]},
            "Balkabagi": {"Kuru": [11,10,9,8,11,10,9,8,7,6,5,4,3,0,0,0,0,0,0]},
            "Yonca": {"Sulu": [4,4,3,3,16,14,12,10,8,7,6,5,4,0,12,9,7,5,0]},
        },
        "Karadeniz": {
            "Bugday": {"Sulu": [17,16,15,12,12,11,9,8,7,6,5,4,3,0,0,0,0,0,0], "Kuru": [14,13,12,10,10,9,7,6,5,4,3,0,0,0,14,11,7,0,0]},
            "Misir": {"Sulu": [19,17,15,12,14,13,11,8,7,5,3,0,0,0,0,0,0,0,0]},
            "Cay": {"Kuru": [24,22,20,18,9,8,7,6,5,4,3,2,0,0,14,9,6,4,0]},
            "Findik": {"Kuru": [20,18,16,14,16,14,12,10,9,7,5,0,0,0,12,8,5,3,0]},
            "Nohut": {"Kuru": [5,4,3,3,10,9,8,7,6,5,4,3,0,0,0,0,0,0,0]},
            "Yonca": {"Sulu": [5,4,3,3,16,14,12,10,8,6,5,4,0,0,12,10,7,5,0]},
            "Kivi": {"Sulu": [10,9,8,7,10,9,8,7,6,5,4,3,0,0,0,0,0,0,0]},
            "Laz Biberi": {"Kuru": [8,7,6,5,8,7,6,5,4,3,0,0,0,0,0,0,0,0,0]},
            "Misir (Silajlik)": {"Sulu": [22,20,18,15,14,13,11,8,7,5,3,0,0,0,0,0,0,0,0]},
            "Patates": {"Sulu": [17,15,13,12,13,12,10,9,8,7,5,4,3,0,13,11,7,5,0]},
        },
        "Orta Anadolu": {
            "Bugday": {"Sulu": [16,15,14,12,11,10,8,7,6,5,4,3,0,0,11,8,6,4,0], "Kuru": [9,8,7,5,9,8,7,6,5,4,3,2,0,0,8,6,4,0,0]},
            "Arpa": {"Sulu": [15,14,13,11,10,9,8,7,6,5,3,0,0,0,0,0,0,0,0], "Kuru": [8,7,6,5,8,7,6,5,4,3,3,0,0,0,0,0,0,0,0]},
            "Misir": {"Sulu": [17,16,14,10,12,11,10,9,8,7,5,4,3,0,13,10,7,5,0]},
            "Seker Pancari": {"Sulu": [15,14,12,10,12,10,8,7,6,5,4,3,3,0,15,13,9,6,0]},
            "Nohut": {"Kuru": [5,4,3,3,8,7,5,4,4,3,2,0,0,0,0,0,0,0,0]},
            "Mercimek": {"Kuru": [5,4,3,3,9,8,7,6,5,4,3,0,0,0,0,0,0,0,0]},
            "Fig": {"Kuru": [3,3,0,0,8,7,6,5,4,3,3,2,0,0,0,0,0,0,0]},
            "Patates": {"Sulu": [16,16,14,12,14,12,10,9,8,6,5,4,3,0,15,12,8,6,0]},
            "Elma": {"Sulu": [10,9,7,6,12,10,8,7,6,5,4,3,0,0,10,8,6,4,0]},
            "Kayisi": {"Sulu": [12,10,8,6,12,10,8,7,6,5,4,3,0,0,12,10,8,6,0]},
            "Havuc": {"Sulu": [19,18,17,14,13,12,11,9,8,7,6,5,4,0,0,0,0,0,0]},
            "Sogan": {"Sulu": [14,13,12,10,10,9,8,7,6,5,4,3,0,0,0,0,0,0,0]},
            "Kimyon": {"Kuru": [11,10,9,7,10,9,8,7,6,5,4,3,0,0,0,0,0,0,0]},
            "Bag": {"Sulu": [15,13,12,11,10,9,7,6,5,4,3,0,0,0,15,12,7,5,0], "Kuru": [10,8,7,6,9,8,6,5,4,4,3,0,0,0,12,9,6,4,0]},
        },
        "Akdeniz": {
            "Bugday": {"Sulu": [18,17,16,14,12,10,8,7,6,5,4,3,0,0,12,9,6,4,0], "Kuru": [14,13,12,10,9,8,7,6,5,4,3,0,0,0,9,7,4,0,0]},
            "Misir": {"Sulu": [19,18,17,15,13,12,11,8,7,5,4,3,0,0,14,10,6,4,0]},
            "Pamuk": {"Sulu": [17,16,15,13,9,8,7,6,5,4,3,0,0,0,13,10,6,4,0]},
            "Narenciye": {"Sulu": [18,17,16,14,12,10,9,8,7,6,5,4,3,0,12,10,7,5,0]},
            "Muz": {"Sulu": [19,18,17,15,14,12,10,9,8,7,6,5,4,0,0,0,0,0,0]},
            "Zeytin": {"Kuru": [10,9,8,6,10,9,8,7,6,5,4,3,0,0,0,0,0,0,0]},
            "Aycicegi": {"Sulu": [20,19,18,16,11,10,9,8,7,6,5,3,0,0,12,10,8,5,0], "Kuru": [13,12,10,8,9,8,7,6,5,4,3,0,0,0,10,8,6,4,0]},
            "Domates": {"Sulu": [19,17,15,13,14,12,11,10,8,6,5,4,0,0,0,0,0,0,0]},
            "Cilek": {"Sulu": [11,10,9,7,9,8,7,6,5,4,3,0,0,0,0,0,0,0,0]},
            "Marul": {"Sulu": [18,17,16,14,18,17,15,13,11,9,7,5,3,0,0,0,0,0,0]},
            "Sogan": {"Sulu": [13,12,11,8,11,9,8,7,6,5,4,3,0,0,0,0,0,0,0]},
            "Sarmsak": {"Sulu": [13,12,11,8,10,9,8,7,6,5,4,3,0,0,0,0,0,0,0]},
            "Karpuz": {"Sulu": [16,14,12,10,12,10,9,8,7,6,5,4,0,0,0,0,0,0,0]},
            "Yerfistigi": {"Sulu": [6,5,4,3,10,9,8,7,6,5,4,0,0,0,13,9,6,4,0]},
            "Susam": {"Sulu": [9,8,6,5,9,8,7,6,5,4,3,0,0,0,10,8,5,0,0]},
            "Avokado": {"Sulu": [14,12,10,8,12,10,9,8,7,6,5,4,0,0,12,10,8,6,0]},
            "Bag": {"Sulu": [16,15,14,12,12,10,9,8,7,6,5,4,3,0,15,11,7,5,0], "Kuru": [11,10,9,8,9,8,7,6,5,4,3,0,0,0,10,8,5,4,0]},
        },
        "Adalar": {
            "Bugday": {"Sulu": [17,16,15,12,12,10,8,7,6,5,4,3,0,0,12,10,7,4,0], "Kuru": [13,12,11,8,9,8,7,6,5,4,3,2,0,0,8,6,4,0,0]},
            "Misir": {"Sulu": [19,18,17,15,13,12,11,8,7,5,3,0,0,0,13,10,8,5,0]},
            "Pamuk": {"Sulu": [16,14,13,11,9,8,7,6,5,4,3,0,0,0,13,10,7,0,0]},
            "Tutun (Kalite)": {"Kuru": [6,5,4,3,8,7,6,5,4,3,3,0,0,0,12,10,8,5,0]},
            "Zeytin": {"Kuru": [10,9,8,6,10,9,8,7,6,5,4,3,0,0,8,6,5,3,0]},
            "Incir": {"Kuru": [12,10,8,7,12,11,10,9,8,7,5,4,0,0,0,0,0,0,0]},
            "Aycicegi": {"Sulu": [18,13,11,9,12,11,10,8,7,6,5,4,3,0,13,10,8,6,0], "Kuru": [10,9,8,7,9,8,7,6,5,4,3,0,0,0,12,10,8,4,0]},
            "Seker Pancari": {"Sulu": [17,16,14,12,13,11,9,8,7,6,5,4,3,0,14,11,8,5,0]},
            "Anason": {"Sulu": [9,8,7,5,9,8,7,6,5,4,3,0,0,0,0,0,0,0,0]},
            "Susam": {"Sulu": [9,8,6,5,9,8,7,6,5,4,3,0,0,0,9,6,5,0,0]},
            "Kekik": {"Kuru": [5,4,3,3,7,6,5,4,4,3,0,0,0,0,0,0,0,0,0]},
            "Lavanta": {"Kuru": [4,4,3,3,6,5,4,4,3,3,0,0,0,0,0,0,0,0,0]},
            "Bag (Sultaniye)": {"Sulu": [14,13,11,9,11,10,9,8,7,6,5,4,3,0,12,10,8,5,0], "Kuru": [11,10,9,8,8,7,6,5,4,3,3,0,0,0,10,8,6,5,0]},
            "Mandalina": {"Sulu": [18,16,14,12,12,11,10,9,8,7,6,5,3,0,12,8,6,4,0]},
        },
        "Guneydogu Anadolu": {
            "Bugday": {"Sulu": [16,15,14,12,13,12,10,9,8,6,5,4,3,0,11,9,7,5,0], "Kuru": [9,8,7,5,10,9,7,5,4,3,3,0,0,0,9,7,5,0,0]},
            "Pamuk": {"Sulu": [15,14,12,10,10,8,7,6,5,4,3,3,0,0,13,11,8,6,0]},
            "Misir": {"Sulu": [17,16,14,10,12,11,10,8,6,4,3,0,0,0,13,10,8,5,0], "2. Urun": [18,16,14,12,13,12,11,8,7,5,3,0,0,0,0,0,0,0,0]},
            "Antep Fistigi": {"Kuru": [12,9,6,3,9,8,7,6,5,4,4,3,0,0,10,8,6,5,0]},
            "Fistik (Yer)": {"Sulu": [6,5,4,4,8,7,6,5,4,4,3,0,0,0,0,0,0,0,0]},
            "Soya": {"Sulu": [6,5,4,4,10,9,8,7,6,5,4,3,0,0,0,0,0,0,0]},
            "Nohut": {"Kuru": [5,4,4,3,10,9,8,7,6,5,4,3,0,0,0,0,0,0,0]},
            "Mercimek": {"Kuru": [5,4,4,3,10,9,8,7,6,5,4,3,0,0,0,0,0,0,0]},
            "Karpuz": {"Sulu": [14,12,10,8,10,9,8,7,6,5,4,3,0,0,0,0,0,0,0]},
            "Zeytin": {"Kuru": [9,8,7,6,10,9,7,6,5,4,3,0,0,0,8,7,6,4,0]},
            "Seker Pancari": {"Sulu": [15,14,12,10,13,12,10,9,8,6,5,4,0,0,14,13,10,7,0]},
            "Susam": {"Sulu": [9,8,6,6,9,8,7,6,5,4,3,0,0,0,0,0,0,0,0]},
            "Bag": {"Sulu": [16,14,12,10,11,9,8,7,6,5,4,3,0,0,12,11,8,6,0], "Kuru": [11,10,9,7,9,8,7,6,5,4,3,2,0,0,10,9,7,4,0]},
        },
        "Dogu Anadolu": {
            "Bugday": {"Sulu": [13,11,9,8,10,9,7,5,4,4,3,3,0,0,12,8,6,4,0], "Kuru": [9,8,6,5,8,7,6,5,4,3,3,0,0,0,9,7,4,0,0]},
            "Arpa": {"Sulu": [12,10,8,6,9,8,7,6,5,4,3,3,0,0,0,0,0,0,0], "Kuru": [8,7,5,4,8,7,6,5,4,4,3,0,0,0,0,0,0,0,0]},
            "Patates": {"Sulu": [15,14,12,10,13,11,9,8,7,6,5,4,3,0,13,11,8,6,0]},
            "Seker Pancari": {"Sulu": [15,14,12,10,14,13,11,9,8,7,6,5,4,0,14,12,8,6,0]},
            "Nohut": {"Kuru": [5,4,3,3,8,7,6,5,4,3,3,2,0,0,0,0,0,0,0]},
            "Mercimek": {"Kuru": [5,4,3,3,8,6,5,4,4,3,3,2,0,0,0,0,0,0,0]},
            "Fig": {"Sulu": [6,5,4,3,11,10,9,8,7,6,5,4,0,0,0,0,0,0,0], "Kuru": [4,4,3,3,8,7,6,5,4,3,3,2,0,0,0,0,0,0,0]},
            "Yonca": {"Sulu": [6,5,4,3,18,16,14,12,10,8,6,4,0,0,11,9,7,5,0]},
            "Cayir (Yem)": {"Sulu": [19,18,17,15,11,10,9,8,7,6,5,4,0,0,0,0,0,0,0]},
            "Meyve": {"Sulu": [12,10,8,7,11,10,8,7,6,5,4,3,0,0,10,8,5,3,0]},
            "Bag": {"Sulu": [13,12,10,7,11,10,9,8,7,6,4,3,0,0,0,0,0,0,0], "Kuru": [10,9,8,7,8,7,6,5,4,3,0,0,0,0,0,0,0,0,0]},
        },
        "Goller": {
            "Bugday": {"Sulu": [17,16,15,12,12,11,9,8,7,6,5,4,0,0,11,9,6,4,0], "Kuru": [10,9,8,8,9,8,7,5,4,3,2,0,0,0,9,7,4,0,0]},
            "Misir": {"Sulu": [18,17,16,12,12,11,10,9,8,7,5,4,3,0,14,12,7,5,0]},
            "Misir (Silajlik)": {"Sulu": [20,18,16,13,12,11,10,9,8,7,5,4,3,0,0,0,0,0,0]},
            "Gul (Isparta)": {"Sulu": [16,12,8,4,8,7,5,4,3,2,0,0,0,0,13,11,9,7,5]},
            "Kiraz": {"Sulu": [14,12,10,8,12,10,9,8,7,6,5,4,0,0,0,0,0,0,0]},
            "Elma": {"Sulu": [12,11,9,7,12,10,9,7,5,4,3,0,0,0,10,8,6,4,0]},
            "Seker Pancari": {"Sulu": [18,16,14,12,13,12,10,9,8,7,5,4,0,0,14,12,10,6,0]},
            "Anason": {"Sulu": [9,8,7,6,9,8,7,6,5,4,3,0,0,0,0,0,0,0,0]},
            "Kimyon": {"Kuru": [11,10,8,6,10,9,8,7,6,5,4,3,0,0,0,0,0,0,0]},
            "Lavanta": {"Kuru": [5,4,3,3,7,6,5,4,3,3,0,0,0,0,0,0,0,0,0]},
            "Mercimek": {"Sulu": [5,4,3,3,9,8,7,6,5,4,2,0,0,0,0,0,0,0,0]},
            "Bag": {"Sulu": [13,12,10,8,12,10,8,7,6,5,4,3,0,0,15,12,8,5,0], "Kuru": [10,8,7,6,8,7,6,5,4,4,3,0,0,0,12,10,6,4,0]},
        },
    }
    # Dogrulama ve eksik teknik tamamlama
    for bolge, bitkiler in ham_veri.items():
        for bitki, tarim_sekileri in bitkiler.items():
            for tarim_sekli, degerler in tarim_sekileri.items():
                while len(degerler) < BEKLENEN_UZUNLUK:
                    degerler.append(0)
                tarim_sekileri[tarim_sekli] = degerler[:BEKLENEN_UZUNLUK]

            # Eksik yetistirme tekniğini tamamla
            if "Sulu" in tarim_sekileri and "Kuru" not in tarim_sekileri:
                tarim_sekileri["Kuru"] = list(tarim_sekileri["Sulu"])
            elif "Kuru" in tarim_sekileri and "Sulu" not in tarim_sekileri:
                tarim_sekileri["Sulu"] = list(tarim_sekileri["Kuru"])

    return ham_veri


# =============================================================================
# 2. GUBRE VERILERI
# =============================================================================

INORGANIK_GUBRELER = {
    "Azot (N)": {
        "Amonyum Sulfat (AS) %21 N": {"kat": 100.0/21, "ek_not": "Asit karakterlidir, kukurt ihtiyacini karsilar."},
        "Kalsiyum Amonyum Nitrat (CAN) %26 N": {"kat": 100.0/26, "ek_not": "Topragi asitlestirmez, bazik topraklar icin ideal."},
        "Amonyum Nitrat (AN) %33 N": {"kat": 100.0/33, "ek_not": "Celtik haric tum bitkiler icin uygundur."},
        "Ure %46 N": {"kat": 100.0/46, "ek_not": "En yogun N kaynagi. Yuzeyden verilecekse hemen surulmelidir."},
        "Diamonyum Fosfat (DAP) %18 N": {"kat": 100.0/18, "ek_not": "Azot ve Fosfor birlikte verilecekse tercih edilir."},
    },
    "Fosfor (P2O5)": {
        "Triple Super Fosfat (TSP) %46 P2O5": {"kat": 100.0/46, "ek_not": "En yaygin fosfor kaynagi. Banda verme onerilir."},
        "Diamonyum Fosfat (DAP) %46 P2O5": {"kat": 100.0/46, "ek_not": "DAP ile N ihtiyaci da %18 oraninda karsilanir."},
        "Monoamonyum Fosfat (MAP) %52 P2O5": {"kat": 100.0/52, "ek_not": "Yapraktan gubreleme ve sera icin mukemmel."},
    },
    "Potasyum (K2O)": {
        "Potasyum Sulfat (PS) %50 K2O": {"kat": 100.0/50, "ek_not": "Tutun, patates, sebze ve meyvelerde sarttir."},
        "Potasyum Nitrat (PN) %44 K2O": {"kat": 100.0/44, "ek_not": "Sera ve yapraktan uygulamada N ve K birlikte verir."},
        "Potasyum Klorur (KCl) %60 K2O": {"kat": 100.0/60, "ek_not": "Tutun ve patates haric tum tarla bitkilerinde kullanilir."},
    },
}

ORGANIK_GUBRELER = {
    "Ciftlik Gubresi": {"N": 0.50, "P2O5": 0.30, "K2O": 0.50, "ton_kg": 1000, "aciklama": "2-4 Ton/da kullanilir."},
    "Tavuk Gubresi (Kuru)": {"N": 3.50, "P2O5": 1.50, "K2O": 2.50, "ton_kg": 100, "aciklama": "200-500 kg/da."},
    "Sivi Tavuk Gubresi": {"N": 0.60, "P2O5": 0.20, "K2O": 0.40, "ton_kg": 1000, "aciklama": "Fertigasyona uygundur."},
    "Kompost": {"N": 1.20, "P2O5": 0.80, "K2O": 1.00, "ton_kg": 100, "aciklama": "1-3 Ton/da."},
    "Tutun Tozu": {"N": 2.00, "P2O5": 0.50, "K2O": 2.00, "ton_kg": 100, "aciklama": "Hizli etki eden organik azot kaynagi."},
}

# =============================================================================
# TURKIYE GUBRELEME REHBERI - BITKI BAZLI DETAYLI VERILER
# =============================================================================

GUBRELEME_REHBERI = {
    "Bugday": {
        "zaman": {
            "taban_gubresi": "Ekim oncesi (Ekim-Kasim)",
            "ilk_bahar": "Subat-Mart (Vegetasyon baslangici)",
            "dapr": "Mart-Nisan (Sap boylanmasi)"
        },
        "sekli": "Taban gubresi topraga serpilip 10-15 cm islenir. Dapr yuzeyden serpilir veya kar ile birlikte verilir.",
        "organik_oneri": "Dekar basina 2-3 ton ciftlik gubresi ekimden 15-20 gun once verilmelidir.",
        "kompost_oneri": "1.5-2 ton/da kompost topraga islenerek verilir.",
        "onemli_not": "Azotun yarisi ekimde, kalan yarisi subat-mart aylarinda DAP olarak verilir. Fosfor ve potasyumun tamami ekim oncesi verilir.",
        "uygulama_detay": "Ekim oncesi: %50 N + %100 P + %100 K. DAP: %50 N (subat-mart)."
    },
    "Arpa": {
        "zaman": {
            "taban_gubresi": "Ekim oncesi (Ekim-Kasim)",
            "ilk_bahar": "Subat-Mart"
        },
        "sekli": "Taban gubresi topraga serpilip 10-15 cm islenir.",
        "organik_oneri": "1.5-2.5 ton/da ciftlik gubresi.",
        "kompost_oneri": "1-2 ton/da kompost.",
        "onemli_not": "Bugdaya kiyasla %15-20 daha az azot gerektirir. Fosfor ve potasyum ekim oncesi verilir.",
        "uygulama_detay": "Ekim oncesi: %50 N + %100 P + %100 K. DAP: %50 N (subat-mart)."
    },
    "Misir": {
        "zaman": {
            "taban_gubresi": "Ekim oncesi (Nisan-Mayis)",
            "ilk_bahar": "Mayis (Fideler 4-5 yaprakli iken)",
            "dapr": "Haziran (Gozlenme donemi)"
        },
        "sekli": "Taban gubresi banta verilir. Dapr yuzeyden serpilir veya sulama suyuyla verilir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi ekimden 20-30 gun once.",
        "kompost_oneri": "2-3 ton/da kompost topraga islenir.",
        "onemli_not": "Misir azot tuketimi yuksek bir bitkidir. Azotun %30'u tabanda, %70'i vegetasyon doneminde verilir.",
        "uygulama_detay": "Taban: %30 N + %100 P + %100 K. DAP: %40 N (Mayis). Son DAP: %30 N (Haziran)."
    },
    "Aycicegi": {
        "zaman": {
            "taban_gubresi": "Ekim oncesi (Subat-Mart)",
            "ilk_bahar": "Nisan (Yapraklanma donemi)"
        },
        "sekli": "Taban gubresi banta verilir. Dapr yuzeyden serpilir.",
        "organik_oneri": "1.5-2 ton/da ciftlik gubresi.",
        "kompost_oneri": "1-1.5 ton/da kompost.",
        "onemli_not": "Fosfor kokunun gelismesi icin onemlidir. Banda verme onerilir.",
        "uygulama_detay": "Taban: %40 N + %100 P + %100 K. DAP: %60 N (Nisan)."
    },
    "Seker Pancari": {
        "zaman": {
            "taban_gubresi": "Ekim oncesi (Nisan-Mayis)",
            "ilk_bahar": "Mayis-Haziran",
            "dapr": "Temmuz (Kok Buyume Donemi)"
        },
        "sekli": "Taban gubresi banta verilir. Dapr sulama suyuyla veya yuzeyden verilir.",
        "organik_oneri": "3-4 ton/da ciftlik gubresi.",
        "kompost_oneri": "2-3 ton/da kompost.",
        "onemli_not": "Seker pancari cok yuksek potasyum ister. Potasyumun %50'si tabanda, %50'si vegetasyonda verilir.",
        "uygulama_detay": "Taban: %30 N + %100 P + %50 K. DAP: %40 N (Mayis). Son DAP: %30 N + %50 K (Temmuz)."
    },
    "Pamuk": {
        "zaman": {
            "taban_gubresi": "Ekim oncesi (Nisan-Mayis)",
            "ilk_bahar": "Mayis (Fideler 4-5 yaprakli)",
            "dapr": "Haziran-Temmuz (Ciceklenme oncesi)"
        },
        "sekli": "Taban gubresi banta verilir. Dapr sulama suyuyla verilir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "2-3 ton/da kompost.",
        "onemli_not": "Pamuk azot ve potasyum tuketimi yuksek bir bitkidir. Azotun %40'i tabanda, %60'i vegetasyonda verilir.",
        "uygulama_detay": "Taban: %40 N + %100 P + %50 K. DAP: %30 N (Mayis). Son DAP: %30 N + %50 K (Haziran-Temmuz)."
    },
    "Domates": {
        "zaman": {
            "taban_gubresi": "Ekim oncesi (Mart-Nisan)",
            "dikim": "Dikim zamani (Mayis-Haziran)",
            "ilk_hasat": "Ilk ciceklenme oncesi",
            "dapr": "Her hasat sonrasi (2-3 hafta aralikla)"
        },
        "sekli": "Taban gubresi banta verilir. Dapr damla sulama ile fertigasyon seklinde verilir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi veya 1.5-2 ton/da kompost.",
        "kompost_oneri": "2-3 ton/da kompost.",
        "onemli_not": "Domates sik sulama ve sik gubreme ister. Fertigasyon ile verim cok artar.",
        "uygulama_detay": "Taban: %30 N + %100 P + %50 K. DAP: %70 N + %50 K (sulama suyyla, hasat boyunca)."
    },
    "Salatalik": {
        "zaman": {
            "taban_gubresi": "Ekim oncesi (Mart-Nisan)",
            "dikim": "Dikim zamani (Mayis)",
            "dapr": "Her hasat sonrasi (2 hafta aralikla)"
        },
        "sekli": "Taban gubresi banta verilir. Dapr damla sulama ile fertigasyon.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "2 ton/da kompost.",
        "onemli_not": "Salatalik neme cok duyarlidir. Duzenli sulama ve gubreme gerekir.",
        "uygulama_detay": "Taban: %30 N + %100 P + %50 K. DAP: %70 N + %50 K (sulama suyyla)."
    },
    "Patates": {
        "zaman": {
            "taban_gubresi": "Ekim oncesi (Mart-Nisan)",
            "ilk_bahar": "Dikimden 2-3 hafta once",
            "dapr": "Dikimden 30-40 gun sonra (goguslenme donemi)"
        },
        "sekli": "Taban gubresi banta verilir. Dapr yuzeyden serpilir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Patates potasyum ve azot ister. Fazla azot kaliteyi dusurur.",
        "uygulama_detay": "Taban: %30 N + %100 P + %70 K. Son DAP: %70 N + %30 K (goguslenme oncesi)."
    },
    "Findik": {
        "zaman": {
            "taban_gubresi": "Ekim-Kasim (Sonbahar)",
            "ilk_bahar": "Subat-Mart (Vegetasyon baslangici)"
        },
        "sekli": "Dekar basina 3-5 kg gubre, agacin taç alti alana serpilir ve 10-15 cm islenir.",
        "organik_oneri": "3-4 ton/da ciftlik gubresi.",
        "kompost_oneri": "2-3 ton/da kompost.",
        "onemli_not": "Findik da azot ve potasyum ister. Fosfor tuylenme doneminde onemlidir.",
        "uygulama_detay": "Sonbahar: %50 N + %100 P + %50 K. Bahar: %50 N + %50 K."
    },
    "Zeytin": {
        "zaman": {
            "taban_gubresi": "Ekim-Kasim (Sonbahar)",
            "ilk_bahar": "Subat-Mart"
        },
        "sekli": "Agacin taç projeksiyonu altina serpilip 15-20 cm islenir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Zeytin verim yili ile dinlenme yili arasinda gubreme miktarini degistirin.",
        "uygulama_detay": "Sonbahar: %40 N + %100 P + %50 K. Bahar: %60 N + %50 K."
    },
    "Bag": {
        "zaman": {
            "taban_gubresi": "Ekim-Kasim (Sonbahar)",
            "ilk_bahar": "Subat-Mart (Kis uykusundan uyanma oncesi)",
            "dapr": "Mayis-Haziran (Ciceklenme oncesi)"
        },
        "sekli": "Taç altina serpilip 15-20 cm islenir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Bagda azot cok fazla verilmemelidir. fazlasi yaprak buyumesini tesvik eder.",
        "uygulama_detay": "Sonbahar: %30 N + %100 P + %50 K. Bahar: %70 N + %50 K."
    },
    "Nohut": {
        "zaman": {
            "taban_gubresi": "Ekim oncesi (Subat-Mart)"
        },
        "sekli": "Taban gubresi banta verilir.",
        "organik_oneri": "1-1.5 ton/da ciftlik gubresi.",
        "kompost_oneri": "1 ton/da kompost.",
        "onemli_not": "Nohut baklagildir ve kendi azotunu uretir. Azot gubresi sadece fideler cok zayif ise verilir.",
        "uygulama_detay": "Azot gerekmez. Sadece %100 P + %100 K tabanda verilir."
    },
    "Mercimek": {
        "zaman": {
            "taban_gubresi": "Ekim oncesi (Subat-Mart)"
        },
        "sekli": "Taban gubresi banta verilir.",
        "organik_oneri": "1-1.5 ton/da ciftlik gubresi.",
        "kompost_oneri": "1 ton/da kompost.",
        "onemli_not": "Mercimek baklagildir. Azot gubresi genellikle gerekmez.",
        "uygulama_detay": "Azot gerekmez. Sadece %100 P + %100 K tabanda verilir."
    },
    "Cay": {
        "zaman": {
            "ilk_bahar": "Subat-Mart (Ilk hasat oncesi)",
            "dapr": "Her hasat sonrasi (3-4 hasat doneminde)"
        },
        "sekli": "Bitki siralari arasina serpilip 5-10 cm islenir.",
        "organik_oneri": "3-4 ton/da ciftlik gubresi yilin basinda.",
        "kompost_oneri": "2-3 ton/da kompost.",
        "onemli_not": "Cay surekli hasat edilen bir bitkidir. Sik sik azot verilmesi gerekir.",
        "uygulama_detay": "Her hasat sonrasi: 20-25 kg N/da (yil icinde toplam 100-120 kg N/da)."
    },
    "Seker Pancari": {
        "zaman": {
            "taban_gubresi": "Ekim oncesi (Nisan-Mayis)",
            "ilk_bahar": "Mayis-Haziran",
            "dapr": "Temmuz (Kok Buyume Donemi)"
        },
        "sekli": "Taban gubresi banta verilir. Dapr sulama suyuyla veya yuzeyden verilir.",
        "organik_oneri": "3-4 ton/da ciftlik gubresi.",
        "kompost_oneri": "2-3 ton/da kompost.",
        "onemli_not": "Seker pancari cok yuksek potasyum ister. Potasyumun %50'si tabanda, %50'si vegetasyonda verilir.",
        "uygulama_detay": "Taban: %30 N + %100 P + %50 K. DAP: %40 N (Mayis). Son DAP: %30 N + %50 K (Temmuz)."
    },
    "Tutun": {
        "zaman": {
            "taban_gubresi": "Ekim oncesi (Nisan-Mayis)",
            "dikim": "Dikim zamani (Mayis-Haziran)",
            "dapr": "Dikimden 30-40 gun sonra"
        },
        "sekli": "Taban gubresi banta verilir. Dapr yuzeyden serpilir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "2 ton/da kompost.",
        "onemli_not": "Tutunda azot fazlasi kaliteyi dusurur. Dikkatli dozaj gerekir.",
        "uygulama_detay": "Taban: %30 N + %100 P + %50 K. DAP: %70 N (dikimden 30-40 gun sonra)."
    },
    "Kekik": {
        "zaman": {
            "taban_gubresi": "Ekim oncesi (Mart-Nisan)"
        },
        "sekli": "Sira uzerine serpilir.",
        "organik_oneri": "1-1.5 ton/da ciftlik gubresi.",
        "kompost_oneri": "1 ton/da kompost.",
        "onemli_not": "Kekik cok az gubre ister. Fazlasi yag kalitesini dusurur.",
        "uygulama_detay": "Sadece taban gubresi: %100 P + %100 K. Azot sadece cok zaysa."
    },
    "Kimyon": {
        "zaman": {
            "taban_gubresi": "Ekim oncesi (Subat-Mart)"
        },
        "sekli": "Sira uzerine serpilir.",
        "organik_oneri": "1-1.5 ton/da ciftlik gubresi.",
        "kompost_oneri": "1 ton/da kompost.",
        "onemli_not": "Kimyon az gubreme ister. Fazlasi kokusunu ve lezzetini etkiler.",
        "uygulama_detay": "Taban: %100 P + %100 K. Azot sadece cok zaysa."
    },
    "Narenciye": {
        "zaman": {
            "taban_gubresi": "Ekim-Kasim (Sonbahar)",
            "ilk_bahar": "Subat-Mart (Ciceklenme oncesi)"
        },
        "sekli": "Taç altina serpilip 15-20 cm islenir.",
        "organik_oneri": "3-4 ton/da ciftlik gubresi.",
        "kompost_oneri": "2-3 ton/da kompost.",
        "onemli_not": "Narenciyede potasyum meyve kalitesini cok etkiler.",
        "uygulama_detay": "Sonbahar: %40 N + %100 P + %50 K. Bahar: %60 N + %50 K."
    },
    "Muz": {
        "zaman": {
            "taban_gubresi": "Dikim oncesi (Yil boyunca dikim yapilabilir)",
            "dapr": "Her hasat sonrasi (aylik)"
        },
        "sekli": "Sulanma kanallarina verilir.",
        "organik_oneri": "4-5 ton/da ciftlik gubresi.",
        "kompost_oneri": "3-4 ton/da kompost.",
        "onemli_not": "Muz cok yogun gubreme ister. Aylik duzenli gubreme gerekir.",
        "uygulama_detay": "Aylik: 30-40 kg N + 10-15 kg P + 40-50 kg K (damla sulamayla)."
    },
    "Elma": {
        "zaman": {
            "taban_gubresi": "Ekim-Kasim (Sonbahar)",
            "ilk_bahar": "Subat-Mart (Kis uykusundan uyanma)"
        },
        "sekli": "Taç altina serpilip 15-20 cm islenir.",
        "organik_oneri": "3-4 ton/da ciftlik gubresi.",
        "kompost_oneri": "2-3 ton/da kompost.",
        "onemli_not": "Elmada azot fazlasi agac buyumesini tesvik eder, meyve kalitesini dusurur.",
        "uygulama_detay": "Sonbahar: %30 N + %100 P + %50 K. Bahar: %70 N + %50 K."
    },
    "Antep Fistigi": {
        "zaman": {
            "taban_gubresi": "Ekim-Kasim (Sonbahar)",
            "ilk_bahar": "Subat-Mart"
        },
        "sekli": "Taç altina serpilip 15-20 cm islenir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Antep fistigi uzun sureli gubreme gerektirir. Potasyum cok onemlidir.",
        "uygulama_detay": "Sonbahar: %30 N + %100 P + %60 K. Bahar: %70 N + %40 K."
    },
    "Anason": {
        "zaman": {"taban_gubresi": "Ekim oncesi (Subat-Mart)"},
        "sekli": "Sira uzerine serpilir, 5-8 cm islenir.",
        "organik_oneri": "1-1.5 ton/da ciftlik gubresi.",
        "kompost_oneri": "1 ton/da kompost.",
        "onemli_not": "Anason sicak ve kurak iklimi sever. Fazla gubre yag kalitesini dusurur.",
        "uygulama_detay": "Taban: %100 P + %100 K. Azot sadece cok zaysa %50 oraninda verilir."
    },
    "Armut": {
        "zaman": {"taban_gubresi": "Ekim-Kasim (Sonbahar)", "ilk_bahar": "Subat-Mart (Kis uykusundan uyanma)"},
        "sekli": "Taç altina serpilip 15-20 cm islenir.",
        "organik_oneri": "3-4 ton/da ciftlik gubresi.",
        "kompost_oneri": "2-3 ton/da kompost.",
        "onemli_not": "Armut agaclari uzun omurludur. Duzenli gubreme onemlidir.",
        "uygulama_detay": "Sonbahar: %30 N + %100 P + %50 K. Bahar: %70 N + %50 K."
    },
    "Avokado": {
        "zaman": {"taban_gubresi": "Nisan-Mayis (Ilkbahar)", "dapr": "Her 2 ayda bir"},
        "sekli": "Taç altina serpilip sulama ile verilir.",
        "organik_oneri": "3-4 ton/da ciftlik gubresi.",
        "kompost_oneri": "2-3 ton/da kompost.",
        "onemli_not": "Avokado soguga cok duyarlidir. Potasyum dona karsi korur.",
        "uygulama_detay": "Baslangic: %30 N + %100 P + %50 K. DAP: %70 N + %50 K (2 ay aralikla)."
    },
    "Bag (Sultaniye)": {
        "zaman": {"taban_gubresi": "Ekim-Kasim (Sonbahar)", "ilk_bahar": "Subat-Mart", "dapr": "Mayis-Haziran (Ciceklenme oncesi)"},
        "sekli": "Taç altina serpilip 15-20 cm islenir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Sultaniye cok hizli buyuyen bir bag cesididir. Azota dikkat edilmeli.",
        "uygulama_detay": "Sonbahar: %30 N + %100 P + %50 K. Bahar: %70 N + %50 K."
    },
    "Balkabagi": {
        "zaman": {"taban_gubresi": "Ekim oncesi (Nisan-Mayis)", "dapr": "Goguslenme doneminde"},
        "sekli": "Taban gubresi banta verilir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Balkabagi yuksek potasyum ister. Meyve eti kalinligi icin onemlidir.",
        "uygulama_detay": "Taban: %30 N + %100 P + %60 K. DAP: %70 N + %40 K (goguslenme doneminde)."
    },
    "Cayir (Yem)": {
        "zaman": {"taban_gubresi": "Ekim oncesi", "dapr": "Her hasat sonrasi"},
        "sekli": "Yuzeyden serpilir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Yemlik bitkilerde azot verimi dogrudan etkiler.",
        "uygulama_detay": "Taban: %40 N + %100 P + %50 K. DAP: %60 N (her hasat sonrasi)."
    },
    "Cilek": {
        "zaman": {"taban_gubresi": "Agustos-Eylul (Dikim oncesi)", "dapr": "Mart-Temmuz (Hasat doneminde)"},
        "sekli": "Damla sulama ile fertigasyon veya serpmeli.",
        "organik_oneri": "1.5-2 ton/da ciftlik gubresi.",
        "kompost_oneri": "1-1.5 ton/da kompost.",
        "onemli_not": "Cilek sicak mevsimde sik sulama ve gubreme ister.",
        "uygulama_detay": "Taban: %30 N + %100 P + %50 K. DAP: %70 N + %50 K (hasat boyunca)."
    },
    "Domates (Sera)": {
        "zaman": {"taban_gubresi": "Dikim oncesi", "dikim": "Dikim zamani", "dapr": "Her hasat sonrasi (2 hafta aralikla)"},
        "sekli": "Damla sulama ile fertigasyon onerilir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "2 ton/da kompost.",
        "onemli_not": "Sera kosullarinda sik sulama ve gubreme gerekir.",
        "uygulama_detay": "Taban: %30 N + %100 P + %50 K. DAP: %70 N + %50 K (sulama suyuyla)."
    },
    "Fig": {
        "zaman": {"taban_gubresi": "Ekim-Kasim (Sonbahar)", "ilk_bahar": "Subat-Mart"},
        "sekli": "Taç altina serpilip 15 cm islenir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Incir agaclari az gubreme ister. Fazlasi dallanmayi engeller.",
        "uygulama_detay": "Sonbahar: %30 N + %100 P + %50 K. Bahar: %70 N + %50 K."
    },
    "Fistik (Yer)": {
        "zaman": {"taban_gubresi": "Ekim oncesi (Nisan-Mayis)", "dapr": "Ciceklenme doneminde"},
        "sekli": "Banda verilir.",
        "organik_oneri": "1-1.5 ton/da ciftlik gubresi.",
        "kompost_oneri": "1 ton/da kompost.",
        "onemli_not": "Yer fistigi baklagildir. Azot ihtiyaci dusuktur.",
        "uygulama_detay": "Taban: %100 P + %100 K. Azot sadece cok zaysa verilir."
    },
    "Gul (Isparta)": {
        "zaman": {"taban_gubresi": "Ekim-Kasim (Sonbahar)", "ilk_bahar": "Subat-Mart", "dapr": "Her hasat sonrasi"},
        "sekli": "Sira uzerine serpilip sulama ile verilir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Gul yagi uretimi icin ozel gubreme gerekir. Azot ve potasyum onceliklidir.",
        "uygulama_detay": "Taban: %30 N + %100 P + %50 K. DAP: %70 N + %50 K (hasat sonrasi)."
    },
    "Havuc": {
        "zaman": {"taban_gubresi": "Ekim oncesi (Nisan-Mayis)", "dapr": "Kok buyume doneminde"},
        "sekli": "Banda verilir.",
        "organik_oneri": "1.5-2 ton/da ciftlik gubresi.",
        "kompost_oneri": "1-1.5 ton/da kompost.",
        "onemli_not": "Havucta fazla azot yan dallanmayi tesvik eder.",
        "uygulama_detay": "Taban: %40 N + %100 P + %50 K. DAP: %60 N (kok buyume doneminde)."
    },
    "Incir": {
        "zaman": {"taban_gubresi": "Ekim-Kasim (Sonbahar)", "ilk_bahar": "Subat-Mart"},
        "sekli": "Taç altina serpilip 15 cm islenir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Incir kurutma icin yetistiriliyorsa potasyum cok onemlidir.",
        "uygulama_detay": "Sonbahar: %30 N + %100 P + %60 K. Bahar: %70 N + %40 K."
    },
    "Kabak": {
        "zaman": {"taban_gubresi": "Ekim oncesi (Nisan-Mayis)", "dapr": "Ciceklenme ve meyve tutma doneminde"},
        "sekli": "Banda verilir veya damla sulama ile.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Kabak ciceklenme doneminde potasyum ister.",
        "uygulama_detay": "Taban: %30 N + %100 P + %50 K. DAP: %70 N + %50 K (ciceklenme sonrasi)."
    },
    "Karpuz": {
        "zaman": {"taban_gubresi": "Ekim oncesi (Nisan-Mayis)", "dapr": "Goguslenme ve meyve buyume doneminde"},
        "sekli": "Banda verilir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Karpuzda son donem azot verimi kaliteyi dusurur.",
        "uygulama_detay": "Taban: %30 N + %100 P + %50 K. DAP: %70 N + %50 K (goguslenme oncesi)."
    },
    "Kayisi": {
        "zaman": {"taban_gubresi": "Ekim-Kasim (Sonbahar)", "ilk_bahar": "Subat-Mart"},
        "sekli": "Taç altina serpilip 15 cm islenir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Kayisi agaclari az gubreme ister. Fazlasi yanik hastaligini tetikler.",
        "uygulama_detay": "Sonbahar: %30 N + %100 P + %50 K. Bahar: %70 N + %50 K."
    },
    "Kiraz": {
        "zaman": {"taban_gubresi": "Ekim-Kasim (Sonbahar)", "ilk_bahar": "Subat-Mart"},
        "sekli": "Taç altina serpilip 15-20 cm islenir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Kirazda fazla azot yaprak leke hastaligini artirir.",
        "uygulama_detay": "Sonbahar: %30 N + %100 P + %50 K. Bahar: %70 N + %50 K."
    },
    "Kivi": {
        "zaman": {"taban_gubresi": "Ekim-Kasim (Sonbahar)", "ilk_bahar": "Subat-Mart", "dapr": "Mayis-Haziran"},
        "sekli": "Taç altina serpilip sulama ile verilir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Kivi yuksek nem ve sulama ister. Potasyum meyve kalitesini artirir.",
        "uygulama_detay": "Sonbahar: %30 N + %100 P + %50 K. Bahar: %70 N + %50 K."
    },
    "Lavanta": {
        "zaman": {"taban_gubresi": "Ekim oncesi (Mart-Nisan)"},
        "sekli": "Sira uzerine serpilir.",
        "organik_oneri": "1-1.5 ton/da ciftlik gubresi.",
        "kompost_oneri": "1 ton/da kompost.",
        "onemli_not": "Lavanta cok az gubre ister. Fazlasi yag kalitesini ve miktarini dusurur.",
        "uygulama_detay": "Taban: %100 P + %100 K. Azot sadece cok zaysa verilir."
    },
    "Laz Biberi": {
        "zaman": {"taban_gubresi": "Ekim oncesi (Nisan-Mayis)", "dapr": "Dikimden 30-40 gun sonra"},
        "sekli": "Sira uzerine serpilir.",
        "organik_oneri": "1.5-2 ton/da ciftlik gubresi.",
        "kompost_oneri": "1-1.5 ton/da kompost.",
        "onemli_not": "Aciliyi artirmak icin son donem azot verilmemeli.",
        "uygulama_detay": "Taban: %30 N + %100 P + %50 K. DAP: %70 N (dikimden 30-40 gun sonra)."
    },
    "Mandalina": {
        "zaman": {"taban_gubresi": "Ekim-Kasim (Sonbahar)", "ilk_bahar": "Subat-Mart"},
        "sekli": "Taç altina serpilip 15-20 cm islenir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Mandalinada potasyum kabuk kalinligini ve rengini etkiler.",
        "uygulama_detay": "Sonbahar: %30 N + %100 P + %50 K. Bahar: %70 N + %50 K."
    },
    "Marul": {
        "zaman": {"taban_gubresi": "Ekim oncesi (Agustos-Eylul)", "dapr": "Dikimden 20-30 gun sonra"},
        "sekli": "Banda verilir veya sulama suyuyla.",
        "organik_oneri": "1.5-2 ton/da ciftlik gubresi.",
        "kompost_oneri": "1-1.5 ton/da kompost.",
        "onemli_not": "Marul hizli buyuyen bir sebzedir. Sik gubreme gerekir.",
        "uygulama_detay": "Taban: %30 N + %100 P + %50 K. DAP: %70 N (dikimden 20-30 gun sonra)."
    },
    "Meyve": {
        "zaman": {"taban_gubresi": "Ekim-Kasim (Sonbahar)", "ilk_bahar": "Subat-Mart"},
        "sekli": "Taç altina serpilip 15-20 cm islenir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Genel meyve agaclari icin dengeli gubreme onemlidir.",
        "uygulama_detay": "Sonbahar: %30 N + %100 P + %50 K. Bahar: %70 N + %50 K."
    },
    "Misir (Silajlik)": {
        "zaman": {"taban_gubresi": "Ekim oncesi (Nisan-Mayis)", "dapr": "Mayis-Haziran"},
        "sekli": "Taban gubresi banta verilir. Dapr yuzeyden serpilir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "2 ton/da kompost.",
        "onemli_not": "Silajlik misirda kitlesel verim onceliklidir. Azot yuksek verilmeli.",
        "uygulama_detay": "Taban: %30 N + %100 P + %50 K. DAP: %70 N (Mayis-Haziran)."
    },
    "Salatalik (Sera)": {
        "zaman": {"taban_gubresi": "Dikim oncesi", "dapr": "Her hasat sonrasi (2 hafta aralikla)"},
        "sekli": "Damla sulama ile fertigasyon onerilir.",
        "organik_oneri": "2 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5 ton/da kompost.",
        "onemli_not": "Sera satisinda sik sulama ve gubreme gerekir.",
        "uygulama_detay": "Taban: %30 N + %100 P + %50 K. DAP: %70 N + %50 K (sulama suyuyla)."
    },
    "Sarmsak": {
        "zaman": {"taban_gubresi": "Ekim-Kasim (Sonbahar)"},
        "sekli": "Sira uzerine serpilip 5-8 cm islenir.",
        "organik_oneri": "1.5-2 ton/da ciftlik gubresi.",
        "kompost_oneri": "1-1.5 ton/da kompost.",
        "onemli_not": "Sarmsak az gubre ister. Fazlasi basak boyunu uzatir, dis sayisini azaltir.",
        "uygulama_detay": "Taban: %100 P + %100 K. Azot sadece cok zaysa verilir."
    },
    "Seftali": {
        "zaman": {"taban_gubresi": "Ekim-Kasim (Sonbahar)", "ilk_bahar": "Subat-Mart"},
        "sekli": "Taç altina serpilip 15-20 cm islenir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "1.5-2 ton/da kompost.",
        "onemli_not": "Seftalide azot fazlasi yaprak biti ve unlu bitlenme riskini artirir.",
        "uygulama_detay": "Sonbahar: %30 N + %100 P + %50 K. Bahar: %70 N + %50 K."
    },
    "Sogan": {
        "zaman": {"taban_gubresi": "Ekim oncesi (Subat-Mart)", "dapr": "Soğanlanma doneminde"},
        "sekli": "Banda verilir.",
        "organik_oneri": "1.5-2 ton/da ciftlik gubresi.",
        "kompost_oneri": "1-1.5 ton/da kompost.",
        "onemli_not": "Soganda son donem azot verimi depolamayi olumsuz etkiler.",
        "uygulama_detay": "Taban: %40 N + %100 P + %50 K. DAP: %60 N (soğanlanma oncesi)."
    },
    "Soya": {
        "zaman": {"taban_gubresi": "Ekim oncesi (Nisan-Mayis)", "dapr": "Ciceklenme doneminde"},
        "sekli": "Banda verilir.",
        "organik_oneri": "1-1.5 ton/da ciftlik gubresi.",
        "kompost_oneri": "1 ton/da kompost.",
        "onemli_not": "Soya baklagildir ve kendi azotunu uretir.",
        "uygulama_detay": "Taban: %100 P + %100 K. Azot sadece cok zaysa verilir."
    },
    "Susam": {
        "zaman": {"taban_gubresi": "Ekim oncesi (Nisan-Mayis)", "dapr": "Dikimden 30-40 gun sonra"},
        "sekli": "Sira uzerine serpilir.",
        "organik_oneri": "1-1.5 ton/da ciftlik gubresi.",
        "kompost_oneri": "1 ton/da kompost.",
        "onemli_not": "Susam cok az gubre ister. Fazlasi yag oranini dusurur.",
        "uygulama_detay": "Taban: %30 N + %100 P + %50 K. DAP: %70 N (dikimden 30-40 gun sonra)."
    },
    "Tutun (Kalite)": {
        "zaman": {"taban_gubresi": "Ekim oncesi (Nisan-Mayis)", "dapr": "Dikimden 30-40 gun sonra"},
        "sekli": "Banda verilir.",
        "organik_oneri": "2-3 ton/da ciftlik gubresi.",
        "kompost_oneri": "2 ton/da kompost.",
        "onemli_not": "Kalite tutununda azot cok dikkatli verilmelidir. fazlasi kaliteyi dusurur.",
        "uygulama_detay": "Taban: %20 N + %100 P + %50 K. DAP: %80 N (dikimden 30-40 gun sonra)."
    },
    "Yerfistigi": {
        "zaman": {"taban_gubresi": "Ekim oncesi (Nisan-Mayis)", "dapr": "Ciceklenme ve kok olusturma doneminde"},
        "sekli": "Banda verilir.",
        "organik_oneri": "1-1.5 ton/da ciftlik gubresi.",
        "kompost_oneri": "1 ton/da kompost.",
        "onemli_not": "Yer fistigi baklagildir. Azot ihtiyaci dusuktur.",
        "uygulama_detay": "Taban: %100 P + %100 K. Azot sadece cok zaysa verilir."
    },
    "Yonca": {
        "zaman": {"taban_gubresi": "Ekim oncesi (Eylul-Ekim)", "dapr": "Ilkbahar (Subat-Mart)"},
        "sekli": "Yuzeyden serpilir.",
        "organik_oneri": "1-1.5 ton/da ciftlik gubresi.",
        "kompost_oneri": "1 ton/da kompost.",
        "onemli_not": "Yonca baklagildir ve kendi azotunu uretir. Azot gubresi gerekmez.",
        "uygulama_detay": "Taban: %100 P + %100 K. Azot verilmez."
    },
}


# =============================================================================
# 3. YARDIMCI FONKSIYONLAR
# =============================================================================

def n_index(d):
    if d <= 1.0: return 0
    elif d <= 2.0: return 1
    elif d <= 3.0: return 2
    else: return 3

def p_index(d):
    for i, limit in enumerate([1.0,2.0,3.0,4.0,5.0,6.0,7.0,8.0,9.0]):
        if d <= limit: return i
    return 9

def k_index(d):
    for i, limit in enumerate([10.0,20.0,25.0,30.0]):
        if d <= limit: return i
    return 4

def is_baklagil(bitki_adi):
    alt = bitki_adi.lower()
    return any(b.lower() in alt for b in BAKLAGIL_BITKILERI)


# =============================================================================
# 4. TEMA YONETIMI
# =============================================================================

ACIK_TEMA = """
QMainWindow { background-color: #F5F5F5; }
QGroupBox { font-weight: bold; font-size: 13px; border: 2px solid #B0BEC5; border-radius: 8px; margin-top: 12px; padding-top: 14px; }
QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 6px; }
QComboBox { padding: 5px; font-size: 12px; min-height: 24px; }
QDoubleSpinBox, QSpinBox { padding: 4px; font-size: 12px; }
QTableWidget { font-size: 11px; gridline-color: #CFD8DC; selection-background-color: #BBDEFB; }
QTableWidget::item { padding: 4px; }
QHeaderView::section { background-color: #2E7D32; color: white; font-weight: bold; font-size: 11px; padding: 6px; border: 1px solid #1B5E20; }
QTabBar::tab { font-size: 12px; padding: 8px 16px; margin-right: 2px; }
QTabBar::tab:selected { background-color: #2E7D32; color: white; font-weight: bold; }
QTabBar::tab:!selected { background-color: #E8F5E9; }
QMenuBar { background-color: #2E7D32; color: white; font-size: 12px; }
QMenuBar::item:selected { background-color: #1B5E20; }
QMenu { background-color: white; color: #212121; border: 1px solid #B0BEC5; }
QMenu::item:selected { background-color: #E8F5E9; }
QStatusBar { background-color: #2E7D32; color: white; }
QLineEdit { padding: 4px; font-size: 12px; border: 1px solid #B0BEC5; border-radius: 4px; }
QTextEdit { font-size: 12px; border: 1px solid #B0BEC5; border-radius: 4px; }
QPushButton { font-size: 12px; padding: 6px 12px; }
"""

KARANLIK_TEMA = """
QMainWindow { background-color: #1E1E1E; }
QGroupBox { font-weight: bold; font-size: 13px; border: 2px solid #555; border-radius: 8px; margin-top: 12px; padding-top: 14px; color: #E0E0E0; }
QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 6px; }
QComboBox { padding: 5px; font-size: 12px; min-height: 24px; background-color: #2D2D2D; color: #E0E0E0; border: 1px solid #555; }
QDoubleSpinBox, QSpinBox { padding: 4px; font-size: 12px; background-color: #2D2D2D; color: #E0E0E0; border: 1px solid #555; }
QTableWidget { font-size: 11px; gridline-color: #444; background-color: #2D2D2D; color: #E0E0E0; selection-background-color: #2E7D32; }
QTableWidget::item { padding: 4px; }
QHeaderView::section { background-color: #2E7D32; color: white; font-weight: bold; font-size: 11px; padding: 6px; border: 1px solid #1B5E20; }
QTabBar::tab { font-size: 12px; padding: 8px 16px; margin-right: 2px; background-color: #2D2D2D; color: #E0E0E0; }
QTabBar::tab:selected { background-color: #2E7D32; color: white; font-weight: bold; }
QTabBar::tab:!selected { background-color: #3A3A3A; }
QMenuBar { background-color: #2E7D32; color: white; font-size: 12px; }
QMenuBar::item:selected { background-color: #1B5E20; }
QMenu { background-color: #2D2D2D; color: #E0E0E0; border: 1px solid #555; }
QMenu::item:selected { background-color: #2E7D32; }
QStatusBar { background-color: #2E7D32; color: white; }
QLabel { color: #E0E0E0; }
QLineEdit { padding: 4px; font-size: 12px; border: 1px solid #555; border-radius: 4px; background-color: #2D2D2D; color: #E0E0E0; }
QTextEdit { font-size: 12px; border: 1px solid #555; border-radius: 4px; background-color: #2D2D2D; color: #E0E0E0; }
QPushButton { font-size: 12px; padding: 6px 12px; }
QCheckBox { color: #E0E0E0; }
"""


# =============================================================================
# 5. EMAIL AYARLARI DIALOGU
# =============================================================================

class EmailAyarDialog(QDialog):
    def __init__(self, ayarlar, ebeveyn=None):
        super().__init__(ebeveyn)
        self.setWindowTitle(cevir("eposta_ayarlari"))
        self.setMinimumWidth(400)
        self.ayarlar = ayarlar
        self._init_ui()

    def _init_ui(self):
        layout = QGridLayout(self)

        layout.addWidget(QLabel(cevir("smtp_sunucu")), 0, 0)
        self.txt_sunucu = QLineEdit(self.ayarlar.get("smtp_sunucu", "smtp.gmail.com"))
        layout.addWidget(self.txt_sunucu, 0, 1)

        layout.addWidget(QLabel(cevir("smtp_port")), 1, 0)
        self.spin_port = QSpinBox()
        self.spin_port.setRange(1, 65535)
        self.spin_port.setValue(self.ayarlar.get("smtp_port", 587))
        layout.addWidget(self.spin_port, 1, 1)

        layout.addWidget(QLabel(cevir("smtp_kullanici")), 2, 0)
        self.txt_kullanici = QLineEdit(self.ayarlar.get("smtp_kullanici", ""))
        layout.addWidget(self.txt_kullanici, 2, 1)

        layout.addWidget(QLabel(cevir("smtp_sifre")), 3, 0)
        self.txt_sifre = QLineEdit(self.ayarlar.get("smtp_sifre", ""))
        self.txt_sifre.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.txt_sifre, 3, 1)

        btn_layout = QHBoxLayout()
        btn_test = QPushButton(cevir("gonder"))
        btn_test.clicked.connect(self._test_et)
        btn_layout.addWidget(btn_test)

        btn_tamam = QPushButton(cevir("tamam"))
        btn_tamam.clicked.connect(self.accept)
        btn_layout.addWidget(btn_tamam)

        btn_iptal = QPushButton(cevir("iptal"))
        btn_iptal.clicked.connect(self.reject)
        btn_layout.addWidget(btn_iptal)

        layout.addLayout(btn_layout, 4, 0, 1, 2)

    def _test_et(self):
        basarili, mesaj = smtp_test_et(
            self.txt_sunucu.text(), self.spin_port.value(),
            self.txt_kullanici.text(), self.txt_sifre.text()
        )
        if basarili:
            QMessageBox.information(self, cevir("tamam"), mesaj)
        else:
            QMessageBox.warning(self, cevir("uyari"), mesaj)

    def ayarlari_al(self):
        return {
            "smtp_sunucu": self.txt_sunucu.text(),
            "smtp_port": self.spin_port.value(),
            "smtp_kullanici": self.txt_kullanici.text(),
            "smtp_sifre": self.txt_sifre.text(),
        }


# =============================================================================
# 6. EPOSTA GONDERME DIALOGU
# =============================================================================

class EpostaDialog(QDialog):
    def __init__(self, alici="", konu="", icerik="", ek_dosya=None, ebeveyn=None):
        super().__init__(ebeveyn)
        self.setWindowTitle(cevir("eposta_gonder"))
        self.setMinimumWidth(500)
        self.ek_dosya = ek_dosya
        self._init_ui(alici, konu, icerik)

    def _init_ui(self, alici, konu, icerik):
        layout = QVBoxLayout(self)

        layout.addWidget(QLabel("Alici E-posta:"))
        self.txt_alici = QLineEdit(alici)
        layout.addWidget(self.txt_alici)

        layout.addWidget(QLabel("Konu:"))
        self.txt_konu = QLineEdit(konu)
        layout.addWidget(self.txt_konu)

        layout.addWidget(QLabel("Icerik:"))
        self.txt_icerik = QTextEdit(icerik)
        self.txt_icerik.setMinimumHeight(150)
        layout.addWidget(self.txt_icerik)

        if self.ek_dosya:
            layout.addWidget(QLabel(f"Eklenti: {os.path.basename(self.ek_dosya)}"))

        btn_layout = QHBoxLayout()
        btn_gonder = QPushButton(cevir("gonder"))
        btn_gonder.clicked.connect(self.accept)
        btn_layout.addWidget(btn_gonder)
        btn_iptal = QPushButton(cevir("iptal"))
        btn_iptal.clicked.connect(self.reject)
        btn_layout.addWidget(btn_iptal)
        layout.addLayout(btn_layout)

    def bilgileri_al(self):
        return self.txt_alici.text(), self.txt_konu.text(), self.txt_icerik.toPlainText()


# =============================================================================
# 7. ANA UYGULAMA
# =============================================================================

class TarimApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ayarlar = ayarlari_yukle()
        self.veri_tabani = veri_tabani_olustur()
        self.sonuc_kayitlar = []
        self.dil = self.ayarlar.get("dil", "tr")
        translator.dil_ayarla(self.dil)
        self.tema = self.ayarlar.get("tema", "acik")

        self.setWindowTitle(cevir("pencere_basligi"))
        self.setMinimumSize(1200, 750)

        # Ikon
        icon_yolu = kaynak_yolu("app.ico")
        if os.path.exists(icon_yolu):
            self.setWindowIcon(QIcon(icon_yolu))

        self._menu_olustur()
        self._init_ui()
        self._kisayol_ayarla()
        self._tema_uygula()
        self._statusbar_olustur()
        self._otomatik_kaydedici_baslat()
        self._guncelle_bitkiler(self.cmb_bolge.currentText())

    def _statusbar_olustur(self):
        """Durum cubugu olusturur."""
        self.statusBar().showMessage(cevir("tamam"))

    def _otomatik_kaydedici_baslat(self):
        """Otomatik kaydedici timer baslatir."""
        self.otomatik_timer = QTimer(self)
        self.otomatik_timer.timeout.connect(self._otomatik_kaydet)
        self.otomatik_timer.start(60000)  # Her 60 saniyede bir

    def _otomatik_kaydet(self):
        """Son secimleri otomatik kaydeder."""
        if self.ayarlar.get("otomatik_kaydet", True):
            self.ayarlar["son_bolge"] = self.cmb_bolge.currentText()
            self.ayarlar["son_bitki"] = self.cmb_bitki.currentText()
            self.ayarlar["son_tarim"] = self.cmb_tarim.currentText()
            ayarlari_kaydet(self.ayarlar)

    def _kisayol_ayarla(self):
        """Kisayol tuslarini ayarlar."""
        self.shortcut_hesapla = QAction(self)
        self.shortcut_hesapla.setShortcut(QKeySequence("Ctrl+H"))
        self.shortcut_hesapla.triggered.connect(self._hesapla)
        self.addAction(self.shortcut_hesapla)

        self.shortcut_pdf = QAction(self)
        self.shortcut_pdf.setShortcut(QKeySequence("Ctrl+P"))
        self.shortcut_pdf.triggered.connect(self._kaydet_pdf)
        self.addAction(self.shortcut_pdf)

        self.shortcut_eposta = QAction(self)
        self.shortcut_eposta.setShortcut(QKeySequence("Ctrl+E"))
        self.shortcut_eposta.triggered.connect(self._eposta_gonder)
        self.addAction(self.shortcut_eposta)

        self.shortcut_dosya = QAction(self)
        self.shortcut_dosya.setShortcut(QKeySequence("Ctrl+O"))
        self.shortcut_dosya.triggered.connect(self._dosya_yukle)
        self.addAction(self.shortcut_dosya)

        self.shortcut_gecmis = QAction(self)
        self.shortcut_gecmis.setShortcut(QKeySequence("Ctrl+G"))
        self.shortcut_gecmis.triggered.connect(lambda: self.tabs.setCurrentIndex(6))
        self.addAction(self.shortcut_gecmis)

        self.shortcut_konum = QAction(self)
        self.shortcut_konum.setShortcut(QKeySequence("Ctrl+L"))
        self.shortcut_konum.triggered.connect(self._konum_bul)
        self.addAction(self.shortcut_konum)

        self.shortcut_tema = QAction(self)
        self.shortcut_tema.setShortcut(QKeySequence("Ctrl+T"))
        self.shortcut_tema.triggered.connect(self._tema_degistir)
        self.addAction(self.shortcut_tema)

        self.shortcut_dil = QAction(self)
        self.shortcut_dil.setShortcut(QKeySequence("Ctrl+D"))
        self.shortcut_dil.triggered.connect(self._dil_degistir)
        self.addAction(self.shortcut_dil)

    def _tema_degistir(self):
        """Tema degistirir."""
        if self.tema == "acik":
            self.tema = "karanlik"
        else:
            self.tema = "acik"
        self.ayarlar["tema"] = self.tema
        ayarlari_kaydet(self.ayarlar)
        self._tema_uygula()

    def _tema_uygula(self):
        """Mevcut temayi uygular."""
        if self.tema == "karanlik":
            self.setStyleSheet(KARANLIK_TEMA)
        else:
            self.setStyleSheet(ACIK_TEMA)

    def _dil_degistir(self):
        """Dili degistirir."""
        if self.dil == "tr":
            self.dil = "en"
        else:
            self.dil = "tr"
        self.ayarlar["dil"] = self.dil
        translator.dil_ayarla(self.dil)
        ayarlari_kaydet(self.ayarlar)
        self._ui_guncelle()

    def _ui_guncelle(self):
        """UI metinlerini gunceller (dil degisikligi sonrasi)."""
        self.setWindowTitle(cevir("pencere_basligi"))
        self.lbl_bolge.setText(cevir("bolge_secimi"))
        self.lbl_bitki.setText(cevir("bitki_secimi"))
        self.lbl_tarim.setText(cevir("tarim_sekli"))
        self.grp_secim.setTitle(f"1. {cevir('bolge_secimi')}")
        self.grp_analiz.setTitle(f"2. {cevir('toprak_analizi')}")
        self.btn_hesapla.setText(cevir("hesapla"))
        self.btn_kaydet.setText(cevir("pdf_kaydet"))
        self.tabs.setTabText(0, cevir("kimyasal_gubre"))
        self.tabs.setTabText(1, cevir("organik_gubre"))
        self.tabs.setTabText(2, cevir("n_dagilimi"))
        self.tabs.setTabText(3, cevir("karsilastirma"))
        self.tabs.setTabText(4, cevir("gecmis"))
        self.tabs.setTabText(5, "Gubre Rehberi")

    def _menu_olustur(self):
        menubar = self.menuBar()

        # Ayarlar menusu
        ayarlar_menu = menubar.addMenu(cevir("ayarlar"))

        tema_action = QAction(cevir("tema") + " (Ctrl+T)", self)
        tema_action.triggered.connect(self._tema_degistir)
        ayarlar_menu.addAction(tema_action)

        dil_menu = ayarlar_menu.addMenu(cevir("dil"))
        tr_action = QAction(cevir("turkce"), self)
        tr_action.triggered.connect(lambda: self._dil_ayarla("tr"))
        dil_menu.addAction(tr_action)
        en_action = QAction(cevir("ingilizce"), self)
        en_action.triggered.connect(lambda: self._dil_ayarla("en"))
        dil_menu.addAction(en_action)

        ayarlar_menu.addSeparator()
        smtp_action = QAction(cevir("eposta_ayarlari"), self)
        smtp_action.triggered.connect(self._email_ayarlarini_ac)
        ayarlar_menu.addAction(smtp_action)

        # Yardim menusu
        yardim_menu = menubar.addMenu(cevir("hakkinda"))
        hakkinda_action = QAction(cevir("hakkinda"), self)
        hakkinda_action.triggered.connect(self._hakkinda_goster)
        yardim_menu.addAction(hakkinda_action)
        yardim_menu.addSeparator()
        cikis_action = QAction(cevir("cikis"), self)
        cikis_action.setShortcut(QKeySequence("Ctrl+Q"))
        cikis_action.triggered.connect(self.close)
        yardim_menu.addAction(cikis_action)

    def _dil_ayarla(self, dil):
        self.dil = dil
        self.ayarlar["dil"] = dil
        translator.dil_ayarla(dil)
        ayarlari_kaydet(self.ayarlar)
        self._ui_guncelle()

    def _email_ayarlarini_ac(self):
        dialog = EmailAyarDialog(self.ayarlar, self)
        if dialog.exec():
            yeni_ayarlar = dialog.ayarlari_al()
            self.ayarlar.update(yeni_ayarlar)
            ayarlari_kaydet(self.ayarlar)

    def _init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        ana_layout = QHBoxLayout(central)
        ana_layout.setSpacing(10)

        # SOL PANEL
        sol = QVBoxLayout()
        sol.setSpacing(8)

        self.grp_secim = QGroupBox(f"1. {cevir('bolge_secimi')}")
        lay_secim = QVBoxLayout()

        self.lbl_bolge = QLabel(cevir("bolge_secimi"))
        lay_secim.addWidget(self.lbl_bolge)
        self.cmb_bolge = QComboBox()
        self.cmb_bolge.addItems(sorted(self.veri_tabani.keys()))
        self.cmb_bolge.currentTextChanged.connect(self._guncelle_bitkiler)
        lay_secim.addWidget(self.cmb_bolge)

        self.lbl_bitki = QLabel(cevir("bitki_secimi"))
        lay_secim.addWidget(self.lbl_bitki)
        self.cmb_bitki = QComboBox()
        self.cmb_bitki.currentTextChanged.connect(self._guncelle_tarim)
        lay_secim.addWidget(self.cmb_bitki)

        self.lbl_tarim = QLabel(cevir("tarim_sekli"))
        lay_secim.addWidget(self.lbl_tarim)
        self.cmb_tarim = QComboBox()
        lay_secim.addWidget(self.cmb_tarim)

        self.grp_secim.setLayout(lay_secim)
        sol.addWidget(self.grp_secim)

        # Toprak Analizi
        self.grp_analiz = QGroupBox(f"2. {cevir('toprak_analizi')} (kg/da)")
        lay_analiz = QVBoxLayout()

        lay_analiz.addWidget(QLabel(cevir("organik_madde")))
        self.spin_om = QDoubleSpinBox()
        self.spin_om.setRange(0, 10)
        self.spin_om.setDecimals(1)
        self.spin_om.setSuffix(" %")
        self.spin_om.setValue(2.0)
        lay_analiz.addWidget(self.spin_om)

        lay_analiz.addWidget(QLabel(cevir("fosfor")))
        self.spin_p = QDoubleSpinBox()
        self.spin_p.setRange(0, 50)
        self.spin_p.setDecimals(1)
        self.spin_p.setSuffix(" kg P2O5")
        self.spin_p.setValue(5.0)
        lay_analiz.addWidget(self.spin_p)

        lay_analiz.addWidget(QLabel(cevir("potasyum")))
        self.spin_k = QDoubleSpinBox()
        self.spin_k.setRange(0, 60)
        self.spin_k.setDecimals(1)
        self.spin_k.setSuffix(" kg K2O")
        self.spin_k.setValue(15.0)
        lay_analiz.addWidget(self.spin_k)

        self.grp_analiz.setLayout(lay_analiz)
        sol.addWidget(self.grp_analiz)

        # Butonlar
        self.btn_hesapla = QPushButton(cevir("hesapla"))
        self.btn_hesapla.setStyleSheet("background-color: #2E7D32; color: white; font-size: 14px; padding: 10px; border-radius: 5px; font-weight: bold;")
        self.btn_hesapla.clicked.connect(self._hesapla)
        sol.addWidget(self.btn_hesapla)

        self.btn_kaydet = QPushButton(cevir("pdf_kaydet"))
        self.btn_kaydet.setStyleSheet("background-color: #1565C0; color: white; font-size: 12px; padding: 8px; border-radius: 5px;")
        self.btn_kaydet.clicked.connect(self._kaydet_pdf)
        sol.addWidget(self.btn_kaydet)

        self.btn_eposta = QPushButton(cevir("eposta_gonder"))
        self.btn_eposta.setStyleSheet("background-color: #7B1FA2; color: white; font-size: 12px; padding: 8px; border-radius: 5px;")
        self.btn_eposta.clicked.connect(self._eposta_gonder)
        sol.addWidget(self.btn_eposta)

        self.btn_dosya = QPushButton(cevir("dosya_yukle"))
        self.btn_dosya.setStyleSheet("background-color: #E65100; color: white; font-size: 12px; padding: 8px; border-radius: 5px;")
        self.btn_dosya.clicked.connect(self._dosya_yukle)
        sol.addWidget(self.btn_dosya)

        self.btn_konum = QPushButton(cevir("konum_bul"))
        self.btn_konum.setStyleSheet("background-color: #00695C; color: white; font-size: 12px; padding: 8px; border-radius: 5px;")
        self.btn_konum.clicked.connect(self._konum_bul)
        sol.addWidget(self.btn_konum)

        self.lbl_uyari = QLabel("")
        self.lbl_uyari.setWordWrap(True)
        self.lbl_uyari.setStyleSheet("color: #D32F2F; font-style: italic; padding: 4px; background-color: #FFEBEE; border-radius: 4px;")
        sol.addWidget(self.lbl_uyari)
        sol.addStretch()

        # SAG PANEL
        sag = QVBoxLayout()
        sag.setSpacing(8)

        self.tabs = QTabWidget()

        # Sekme 1: Kimyasal
        self.tab_inorg = QWidget()
        lay_inorg = QVBoxLayout(self.tab_inorg)
        self.lbl_saf = QLabel(cevir("saf_madde"))
        self.lbl_saf.setStyleSheet("font-size: 13px; font-weight: bold; color: #1565C0; padding: 8px; background-color: #E3F2FD; border-radius: 5px;")
        self.lbl_saf.setWordWrap(True)
        lay_inorg.addWidget(self.lbl_saf)

        self.tbl_inorg = QTableWidget()
        self.tbl_inorg.setColumnCount(4)
        self.tbl_inorg.setHorizontalHeaderLabels(["Ihtiyac", "Onerilen Gubre", "Miktar (kg/da)", "Uygulama Notu"])
        self.tbl_inorg.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tbl_inorg.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        lay_inorg.addWidget(self.tbl_inorg)
        self.tabs.addTab(self.tab_inorg, cevir("kimyasal_gubre"))

        # Sekme 2: Organik
        self.tab_org = QWidget()
        lay_org = QVBoxLayout(self.tab_org)
        self.chk_org = QCheckBox(cevir("sadece_organik"))
        lay_org.addWidget(self.chk_org)
        self.lbl_org_bilgi = QLabel("")
        self.lbl_org_bilgi.setWordWrap(True)
        self.lbl_org_bilgi.setStyleSheet("color: #2E7D32; font-style: italic; padding: 4px; background-color: #E8F5E9; border-radius: 4px;")
        lay_org.addWidget(self.lbl_org_bilgi)

        self.tbl_org_detay = QTableWidget()
        self.tbl_org_detay.setColumnCount(3)
        self.tbl_org_detay.setHorizontalHeaderLabels(["Organik Kaynak", "Onerilen Miktar", "Aciklama"])
        self.tbl_org_detay.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tbl_org_detay.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        lay_org.addWidget(self.tbl_org_detay)
        self.tabs.addTab(self.tab_org, cevir("organik_gubre"))

        # Sekme 3: Grafikler
        self.tab_grafik = QWidget()
        lay_grafik = QVBoxLayout(self.tab_grafik)
        grafik_splitter = QSplitter(Qt.Orientation.Horizontal)
        self.pasta_grafik = NPKPastalGrafik()
        self.cubuk_grafik = NPKCubukGrafik()
        grafik_splitter.addWidget(self.pasta_grafik)
        grafik_splitter.addWidget(self.cubuk_grafik)
        lay_grafik.addWidget(grafik_splitter)
        self.tabs.addTab(self.tab_grafik, cevir("n_dagilimi"))

        # Sekme 4: Karsilastirma
        self.tab_karsilastirma = QWidget()
        lay_kars = QVBoxLayout(self.tab_karsilastirma)

        self.lbl_kars_baslik = QLabel("Karsilastirmali Analiz - Farkli bolge/bitki/toprak degerleri ile hesaplayin, "
                                      "sonra 'Karsilastir' butonuna basin.")
        self.lbl_kars_baslik.setWordWrap(True)
        self.lbl_kars_baslik.setStyleSheet("font-size: 12px; color: #555; padding: 6px; background-color: #F1F8E9; border-radius: 5px;")
        lay_kars.addWidget(self.lbl_kars_baslik)

        self.tbl_karsilastirma = QTableWidget()
        self.tbl_karsilastirma.setColumnCount(9)
        self.tbl_karsilastirma.setHorizontalHeaderLabels([
            "#", "Bolge", "Bitki", "Yetistirme", "N (kg)", "P (kg)", "K (kg)", "Toplam Maliyet", "Onerilen Gubre"
        ])
        self.tbl_karsilastirma.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tbl_karsilastirma.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.tbl_karsilastirma.setColumnWidth(0, 35)
        self.tbl_karsilastirma.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tbl_karsilastirma.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tbl_karsilastirma.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        lay_kars.addWidget(self.tbl_karsilastirma)

        # Butonlar satiri
        kars_btn_layout = QHBoxLayout()

        self.btn_kars_karsilastir = QPushButton("Karsilastir ve Degerlendir")
        self.btn_kars_karsilastir.setStyleSheet(
            "background-color: #2E7D32; color: white; font-size: 12px; padding: 8px 16px; border-radius: 5px; font-weight: bold;"
        )
        self.btn_kars_karsilastir.clicked.connect(self._karsilastirma_degerlendir)
        kars_btn_layout.addWidget(self.btn_kars_karsilastir)

        btn_kars_sil = QPushButton("Secili Satiri Sil")
        btn_kars_sil.setStyleSheet("background-color: #D32F2F; color: white; font-size: 11px; padding: 6px 12px; border-radius: 5px;")
        btn_kars_sil.clicked.connect(self._karsilastirma_satir_sil)
        kars_btn_layout.addWidget(btn_kars_sil)

        btn_kars_temizle = QPushButton(cevir("temizle"))
        btn_kars_temizle.setStyleSheet("background-color: #757575; color: white; font-size: 11px; padding: 6px 12px; border-radius: 5px;")
        btn_kars_temizle.clicked.connect(self._karsilastirma_temizle)
        kars_btn_layout.addWidget(btn_kars_temizle)

        lay_kars.addLayout(kars_btn_layout)

        # Sonuc paneli
        self.lbl_kars_sonuc = QLabel("")
        self.lbl_kars_sonuc.setWordWrap(True)
        self.lbl_kars_sonuc.setStyleSheet(
            "font-size: 13px; padding: 10px; background-color: #E8F5E9; border: 2px solid #4CAF50; "
            "border-radius: 8px; color: #1B5E20; min-height: 60px;"
        )
        lay_kars.addWidget(self.lbl_kars_sonuc)

        self.tabs.addTab(self.tab_karsilastirma, cevir("karsilastirma"))

        # Sekme 5: Gecmis
        self.tab_gecmis = QWidget()
        lay_gecmis = QVBoxLayout(self.tab_gecmis)

        arama_layout = QHBoxLayout()
        self.txt_arama = QLineEdit()
        self.txt_arama.setPlaceholderText(cevir("filtrele"))
        self.txt_arama.textChanged.connect(self._gecmis_yenile)
        arama_layout.addWidget(self.txt_arama)

        btn_gecmis_yenile = QPushButton(cevir("yukle"))
        btn_gecmis_yenile.clicked.connect(self._gecmis_yenile)
        arama_layout.addWidget(btn_gecmis_yenile)

        btn_gecmis_temizle = QPushButton(cevir("temizle"))
        btn_gecmis_temizle.clicked.connect(self._gecmis_temizle)
        arama_layout.addWidget(btn_gecmis_temizle)

        lay_gecmis.addLayout(arama_layout)

        self.tbl_gecmis = QTableWidget()
        self.tbl_gecmis.setColumnCount(8)
        self.tbl_gecmis.setHorizontalHeaderLabels(["ID", cevir("tarih"), cevir("bolge"), cevir("bitki"), "Tarim", "N", "P", "K"])
        self.tbl_gecmis.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tbl_gecmis.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tbl_gecmis.doubleClicked.connect(self._gecmisten_yukle)
        lay_gecmis.addWidget(self.tbl_gecmis)

        btn_gecmis_sil = QPushButton(cevir("sil"))
        btn_gecmis_sil.clicked.connect(self._gecmis_kayit_sil)
        lay_gecmis.addWidget(btn_gecmis_sil)

        self.tabs.addTab(self.tab_gecmis, cevir("gecmis"))

        # Sekme 6: Gubre Rehberi (Detayli Tavsiye)
        self.tab_rehber = QWidget()
        lay_rehber = QVBoxLayout(self.tab_rehber)

        self.lbl_rehber_baslik = QLabel("Bitki secimi yapin, 'Rehberi Goster' butonuna basin.")
        self.lbl_rehber_baslik.setStyleSheet("font-size: 12px; font-weight: bold; color: #2E7D32; padding: 6px; background-color: #E8F5E9; border-radius: 5px;")
        lay_rehber.addWidget(self.lbl_rehber_baslik)

        btn_rehber_goster = QPushButton("Gubreleme Rehberini Goster")
        btn_rehber_goster.setStyleSheet("background-color: #2E7D32; color: white; font-size: 12px; padding: 8px; border-radius: 5px; font-weight: bold;")
        btn_rehber_goster.clicked.connect(self._rehber_goster)
        lay_rehber.addWidget(btn_rehber_goster)

        self.txt_rehber = QTextEdit()
        self.txt_rehber.setReadOnly(True)
        self.txt_rehber.setStyleSheet("font-size: 12px; padding: 8px; background-color: #FAFAFA; border: 1px solid #E0E0E0; border-radius: 5px;")
        lay_rehber.addWidget(self.txt_rehber)

        self.tabs.addTab(self.tab_rehber, "Gubre Rehberi")

        sag.addWidget(self.tabs)

        # Birlestir
        sol_frame = QFrame()
        sol_frame.setLayout(sol)
        sol_frame.setMaximumWidth(400)
        ana_layout.addWidget(sol_frame)
        ana_layout.addLayout(sag, 1)

        # Son secimleri yukle
        son_bolge = self.ayarlar.get("son_bolge", "")
        if son_bolge and son_bolge in self.veri_tabani.keys():
            self.cmb_bolge.setCurrentText(son_bolge)

    def _guncelle_bitkiler(self, bolge):
        self.cmb_bitki.blockSignals(True)
        self.cmb_bitki.clear()
        if bolge in self.veri_tabani:
            self.cmb_bitki.addItems(sorted(self.veri_tabani[bolge].keys()))
        self.cmb_bitki.blockSignals(False)
        self._guncelle_tarim(self.cmb_bitki.currentText())

    def _guncelle_tarim(self, bitki):
        self.cmb_tarim.blockSignals(True)
        self.cmb_tarim.clear()
        bolge = self.cmb_bolge.currentText()
        if bolge in self.veri_tabani and bitki in self.veri_tabani[bolge]:
            self.cmb_tarim.addItems(self.veri_tabani[bolge][bitki].keys())
        self.cmb_tarim.blockSignals(False)

    def _tablo_doldur(self, tablo, satirlar, bold_col=None):
        tablo.setRowCount(len(satirlar))
        for row, items in enumerate(satirlar):
            for col, val in enumerate(items):
                item = QTableWidgetItem(str(val))
                if bold_col is not None and col == bold_col:
                    item.setFont(QFont("Arial", 10, QFont.Weight.Bold))
                tablo.setItem(row, col, item)

    def _rehber_goster(self):
        bitki = self.cmb_bitki.currentText()
        if not bitki:
            self.txt_rehber.setText("Lutfen once bir bitki secin.")
            return

        rehber = GUBRELEME_REHBERI.get(bitki)
        if not rehber:
            self.txt_rehber.setText(f"{bitki} icin detayli gubreleme rehberi bulunmamaktadir.\n\n"
                                   "Genel oneriler:\n"
                                   "- Toprak analizi yaptirin\n"
                                   "- Organik gubreye oncelik verin\n"
                                   "- Kimyasal gubreyi topraga isleyin\n"
                                   "- Azotu bolumler halinde verin")
            return

        # Mevcut degerleri al
        n_saf = 0
        p_saf = 0
        k_saf = 0
        try:
            bolge = self.cmb_bolge.currentText()
            tarim = self.cmb_tarim.currentText()
            veri = self.veri_tabani[bolge][bitki][tarim]
            om_idx = n_index(self.spin_om.value())
            p_idx = p_index(self.spin_p.value())
            k_idx = k_index(self.spin_k.value())
            n_saf = veri[om_idx]
            p_saf = veri[N_ARALIK_SAYISI + p_idx]
            k_saf = veri[N_ARALIK_SAYISI + P_ARALIK_SAYISI + k_idx]
        except:
            pass

        metin = f"{'='*60}\n"
        metin += f"  TURKIYE GUBRELEME REHBERI - {bitki.upper()}\n"
        metin += f"{'='*60}\n\n"

        # Guvenlik Notu
        metin += "  ONEMLI: Bu bilgiler genel rehber niteligindedir.\n"
        metin += "  Kesin dozajlar icin bolgenizdeki Tarim ve Orman\n"
        metin += "  Md.'ne veya Ziraat Muhendisine danisiniz.\n\n"

        metin += f"{'-'*60}\n"
        metin += f"  1. GUBRELEME ZAMANLARI\n"
        metin += f"{'-'*60}\n"

        zaman = rehber.get("zaman", {})
        for zaman_adi, aciklama in zaman.items():
            metin += f"  {zaman_adi.replace('_', ' ').title()}: {aciklama}\n"

        metin += f"\n{'-'*60}\n"
        metin += f"  2. UYGULAMA SEKLI\n"
        metin += f"{'-'*60}\n"
        metin += f"  {rehber.get('sekli', 'Bilgi mevcut degil.')}\n"

        metin += f"\n{'-'*60}\n"
        metin += f"  3. ORGANIK GUBRE ONERILERI\n"
        metin += f"{'-'*60}\n"
        metin += f"  {rehber.get('organik_oneri', 'Bilgi mevcut degil.')}\n"

        metin += f"\n{'-'*60}\n"
        metin += f"  4. KOMPOST ONERILERI\n"
        metin += f"{'-'*60}\n"
        metin += f"  {rehber.get('kompost_oneri', 'Bilgi mevcut degil.')}\n"

        metin += f"\n{'-'*60}\n"
        metin += f"  5. ONEMLI NOTLAR\n"
        metin += f"{'-'*60}\n"
        metin += f"  {rehber.get('onemli_not', 'Bilgi mevcut degil.')}\n"

        metin += f"\n{'-'*60}\n"
        metin += f"  6. UYGULAMA DETAYI\n"
        metin += f"{'-'*60}\n"
        metin += f"  {rehber.get('uygulama_detay', 'Bilgi mevcut degil.')}\n"

        # Hesaplama sonuclari
        if n_saf > 0 or p_saf > 0 or k_saf > 0:
            metin += f"\n{'-'*60}\n"
            metin += f"  7. MEVCUT TOPRAK ANALIZI SONUCUNA GORE\n"
            metin += f"{'-'*60}\n"
            metin += f"  Organik Madde: %{self.spin_om.value():.1f}\n"
            metin += f"  Fosfor (P2O5): {self.spin_p.value():.1f} kg/da\n"
            metin += f"  Potasyum (K2O): {self.spin_k.value():.1f} kg/da\n\n"
            metin += f"  KIMYASAL IHTIYAC:\n"
            metin += f"  Azot (N): {n_saf} kg/da\n"
            metin += f"  Fosfor (P2O5): {p_saf} kg/da\n"
            metin += f"  Potasyum (K2O): {k_saf} kg/da\n\n"

            # Onerilen gubreler
            metin += f"  ONERILEN KIMYASAL GUBRELER:\n"
            if n_saf > 0:
                ure_kg = n_saf * INORGANIK_GUBRELER["Azot (N)"]["Ure %46 N"]["kat"]
                metin += f"  - Ure %46 N: {ure_kg:.0f} kg/da\n"
            if p_saf > 0:
                dap_kg = p_saf * INORGANIK_GUBRELER["Fosfor (P2O5)"]["Diamonyum Fosfat (DAP) %46 P2O5"]["kat"]
                metin += f"  - DAP %46 P2O5: {dap_kg:.0f} kg/da\n"
            if k_saf > 0:
                ps_kg = k_saf * INORGANIK_GUBRELER["Potasyum (K2O)"]["Potasyum Sulfat (PS) %50 K2O"]["kat"]
                metin += f"  - Potasyum Sulfat %50: {ps_kg:.0f} kg/da\n"

            metin += f"\n  ORGANIK GUBRE ONERILERI:\n"
            for isim, icerik in ORGANIK_GUBRELER.items():
                n_kg = (n_saf / (icerik["N"] / 100)) if n_saf > 0 and icerik["N"] > 0 else 0
                p_kg = (p_saf / (icerik["P2O5"] / 100)) if p_saf > 0 and icerik["P2O5"] > 0 else 0
                k_kg = (k_saf / (icerik["K2O"] / 100)) if k_saf > 0 and icerik["K2O"] > 0 else 0
                toplam_kg = max(n_kg, p_kg, k_kg)
                if toplam_kg > 0:
                    metin += f"  - {isim}: {toplam_kg:.0f} kg/da ({icerik['aciklama']})\n"

        metin += f"\n{'='*60}\n"
        metin += f"  Not: Gubreleme islemlerinden once toprak analizi\n"
        metin += f"  yaptirmaniz onemle tavsiye edilir.\n"
        metin += f"{'='*60}\n"

        self.txt_rehber.setText(metin)
        self.lbl_rehber_baslik.setText(f"{bitki} - Gubreleme Rehberi")
        self.tabs.setCurrentIndex(5)  # Rehber sekmesine gec

    # --- ANA HESAPLAMA ---
    def _hesapla(self):
        try:
            self._hesapla_ic()
        except Exception as e:
            QMessageBox.critical(self, "Hesaplama Hatasi", f"Beklenmeyen bir hata olustu:\n{str(e)}\n\nLutfen tekrar deneyin.")

    def _hesapla_ic(self):
        self.lbl_uyari.setText("")
        self.lbl_org_bilgi.setText("")

        bolge = self.cmb_bolge.currentText()
        bitki = self.cmb_bitki.currentText()
        tarim = self.cmb_tarim.currentText()

        if not bolge or not bitki or not tarim:
            QMessageBox.warning(self, cevir("uyari"), "Lutfen Bolge, Bitki ve Tarim Secenegini eksiksiz secin.")
            return

        try:
            veri = self.veri_tabani[bolge][bitki][tarim]
        except KeyError:
            QMessageBox.critical(self, "Hata", f"{bolge}/{bitki}/{tarim} icin veri bulunamadi.")
            return

        if len(veri) != BEKLENEN_UZUNLUK:
            QMessageBox.critical(self, "Hata", f"Veri yapisi hatali.")
            return

        om_idx = n_index(self.spin_om.value())
        p_idx = p_index(self.spin_p.value())
        k_idx = k_index(self.spin_k.value())

        n_saf = veri[om_idx]
        p_saf = veri[N_ARALIK_SAYISI + p_idx]
        k_saf = veri[N_ARALIK_SAYISI + P_ARALIK_SAYISI + k_idx]

        self.lbl_saf.setText(
            f"{cevir('saf_madde')}:\n"
            f"  Azot (N): {n_saf} kg/da  |  "
            f"Fosfor (P2O5): {p_saf} kg/da  |  "
            f"Potasyum (K2O): {k_saf} kg/da"
        )

        # Uyari sistemi
        uyari_metinleri = []
        if is_baklagil(bitki):
            uyari_metinleri.append(f"Uyari: {bitki} baklagildir.")
        if n_saf == 0 and p_saf == 0 and k_saf == 0:
            uyari_metinleri.append("Kimyasal gubre ilavesine gerek bulunmamaktadir.")
        if self.spin_om.value() < 1.0:
            uyari_metinleri.append("Organik madde cok dusuk. Organik gubreleme oncelikli olmalidir.")
        if self.spin_p.value() > 20:
            uyari_metinleri.append("Fosfor seviyesi yuksek.")
        self.lbl_uyari.setText("\n".join(uyari_metinleri))

        # INORGANIK TABLO
        satirlar_inorg = []
        if n_saf > 0:
            for gubre, oz in INORGANIK_GUBRELER["Azot (N)"].items():
                satirlar_inorg.append([f"{n_saf} kg N", gubre, f"{n_saf * oz['kat']:.1f} kg", oz["ek_not"]])
        if p_saf > 0:
            for gubre, oz in INORGANIK_GUBRELER["Fosfor (P2O5)"].items():
                satirlar_inorg.append([f"{p_saf} kg P2O5", gubre, f"{p_saf * oz['kat']:.1f} kg", oz["ek_not"]])
        if k_saf > 0:
            for gubre, oz in INORGANIK_GUBRELER["Potasyum (K2O)"].items():
                satirlar_inorg.append([f"{k_saf} kg K2O", gubre, f"{k_saf * oz['kat']:.1f} kg", oz["ek_not"]])
        self._tablo_doldur(self.tbl_inorg, satirlar_inorg, bold_col=2)

        # ORGANIK TABLO
        satirlar_org_detay = []
        for isim, icerik in ORGANIK_GUBRELER.items():
            n_kg = (n_saf / (icerik["N"] / 100)) if n_saf > 0 and icerik["N"] > 0 else 0
            p_kg = (p_saf / (icerik["P2O5"] / 100)) if p_saf > 0 and icerik["P2O5"] > 0 else 0
            k_kg = (k_saf / (icerik["K2O"] / 100)) if k_saf > 0 and icerik["K2O"] > 0 else 0
            toplam_kg = max(n_kg, p_kg, k_kg)
            if toplam_kg > 0:
                miktar_str = f"{toplam_kg:.0f} kg/da"
                satirlar_org_detay.append([isim, miktar_str, icerik["aciklama"]])
        self._tablo_doldur(self.tbl_org_detay, satirlar_org_detay)

        # GRAFIKLER
        self.pasta_grafik.guncelle(n_saf, p_saf, k_saf)

        # Cubuk grafik: NPK ihtiyac vs onerilen gubre miktarlari
        n_oneri = n_saf * INORGANIK_GUBRELER["Azot (N)"]["Ure %46 N"]["kat"] if n_saf > 0 else 0
        p_oneri = p_saf * INORGANIK_GUBRELER["Fosfor (P2O5)"]["Diamonyum Fosfat (DAP) %46 P2O5"]["kat"] if p_saf > 0 else 0
        k_oneri = k_saf * INORGANIK_GUBRELER["Potasyum (K2O)"]["Potasyum Sulfat (PS) %50 K2O"]["kat"] if k_saf > 0 else 0
        self.cubuk_grafik.guncelle(n_saf, p_saf, k_saf, n_oneri, p_oneri, k_oneri)

        # En uygun gubre onerisi (karsilastirma icin)
        en_uygun_gubre = ""
        if n_saf > 0:
            ure_kat = INORGANIK_GUBRELER["Azot (N)"]["Ure %46 N"]["kat"]
            en_uygun_gubre = f"Ure {n_saf * ure_kat:.0f} kg"
        elif p_saf > 0:
            dap_kat = INORGANIK_GUBRELER["Fosfor (P2O5)"]["Diamonyum Fosfat (DAP) %46 P2O5"]["kat"]
            en_uygun_gubre = f"DAP {p_saf * dap_kat:.0f} kg"
        elif k_saf > 0:
            ps_kat = INORGANIK_GUBRELER["Potasyum (K2O)"]["Potasyum Sulfat (PS) %50 K2O"]["kat"]
            en_uygun_gubre = f"PS {k_saf * ps_kat:.0f} kg"

        # KARSILASTIRMAYA EKLE
        satir_sayisi = self.tbl_karsilastirma.rowCount()
        self.tbl_karsilastirma.insertRow(satir_sayisi)
        num_item = QTableWidgetItem(str(satir_sayisi + 1))
        num_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tbl_karsilastirma.setItem(satir_sayisi, 0, num_item)
        self.tbl_karsilastirma.setItem(satir_sayisi, 1, QTableWidgetItem(bolge))
        self.tbl_karsilastirma.setItem(satir_sayisi, 2, QTableWidgetItem(bitki))
        self.tbl_karsilastirma.setItem(satir_sayisi, 3, QTableWidgetItem(tarim))
        self.tbl_karsilastirma.setItem(satir_sayisi, 4, QTableWidgetItem(f"{n_saf:.0f}"))
        self.tbl_karsilastirma.setItem(satir_sayisi, 5, QTableWidgetItem(f"{p_saf:.0f}"))
        self.tbl_karsilastirma.setItem(satir_sayisi, 6, QTableWidgetItem(f"{k_saf:.0f}"))

        # Yaklasik maliyet hesapla (kg bazinda)
        ure_kg = n_saf * INORGANIK_GUBRELER["Azot (N)"]["Ure %46 N"]["kat"] if n_saf > 0 else 0
        dap_kg = p_saf * INORGANIK_GUBRELER["Fosfor (P2O5)"]["Diamonyum Fosfat (DAP) %46 P2O5"]["kat"] if p_saf > 0 else 0
        ps_kg = k_saf * INORGANIK_GUBRELER["Potasyum (K2O)"]["Potasyum Sulfat (PS) %50 K2O"]["kat"] if k_saf > 0 else 0
        toplam_gubre_kg = ure_kg + dap_kg + ps_kg
        self.tbl_karsilastirma.setItem(satir_sayisi, 7, QTableWidgetItem(f"{toplam_gubre_kg:.0f} kg/da"))
        self.tbl_karsilastirma.setItem(satir_sayisi, 8, QTableWidgetItem(en_uygun_gubre if en_uygun_gubre else "Gubre Gerekmiyor"))

        # GECMISE KAYDET
        self.sonuc_kayitlar = []
        for s in satirlar_inorg:
            self.sonuc_kayitlar.append({
                "Tur": "Kimyasal", "Bolge": bolge, "Bitki": bitki,
                "Tarim_Sekli": tarim, "Bilgi": s[0], "Gubre": s[1],
                "Miktar": s[2], "Not": s[3],
            })

        database.hesaplama_kaydet(
            bolge, bitki, tarim,
            self.spin_om.value(), self.spin_p.value(), self.spin_k.value(),
            n_saf, p_saf, k_saf, en_uygun_gubre, "", ""
        )

        self.statusBar().showMessage(f"Hesaplama tamamlandi: {bolge}/{bitki} - N:{n_saf} P:{p_saf} K:{k_saf}")
        self._gecmis_yenile()

    # --- KARSILASTIRMA ---
    def _karsilastirma_temizle(self):
        self.tbl_karsilastirma.setRowCount(0)
        self.lbl_kars_sonuc.setText("")

    def _karsilastirma_satir_sil(self):
        secili = self.tbl_karsilastirma.currentRow()
        if secili >= 0:
            self.tbl_karsilastirma.removeRow(secili)
            # Numaralari guncelle
            for i in range(self.tbl_karsilastirma.rowCount()):
                num_item = QTableWidgetItem(str(i + 1))
                num_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tbl_karsilastirma.setItem(i, 0, num_item)

    def _karsilastirma_degerlendir(self):
        satir_sayisi = self.tbl_karsilastirma.rowCount()
        if satir_sayisi < 2:
            self.lbl_kars_sonuc.setText("Karsilastirma icin en az 2 hesaplama yapmaniz gerekir.\n"
                                        "Farkli bolge, bitki veya toprak degerleri ile hesaplayarak listeye ekleyin.")
            return

        # Renkleri temizle
        for row in range(satir_sayisi):
            for col in range(self.tbl_karsilastirma.columnCount()):
                item = self.tbl_karsilastirma.item(row, col)
                if item:
                    item.setBackground(QColor("white"))
                    item.setFont(QFont("Arial", 10))

        # Her satir icin skor hesapla
        skorlar = []
        for row in range(satir_sayisi):
            try:
                n_deger = float(self.tbl_karsilastirma.item(row, 4).text())
                p_deger = float(self.tbl_karsilastirma.item(row, 5).text())
                k_deger = float(self.tbl_karsilastirma.item(row, 6).text())
                toplam_gubre_str = self.tbl_karsilastirma.item(row, 7).text().replace(" kg/da", "").strip()
                toplam_gubre = float(toplam_gubre_str) if toplam_gubre_str else 0
            except (ValueError, AttributeError):
                continue

            npk_toplam = n_deger + p_deger + k_deger
            if toplam_gubre > 0:
                verimlilik = npk_toplam / toplam_gubre
                maliyet_skoru = 100 - min(toplam_gubre * 0.5, 50)
            else:
                verimlilik = 10
                maliyet_skoru = 100

            toplam_skor = (verimlilik * 30) + (maliyet_skoru * 30) + (npk_toplam * 0.4)
            skorlar.append((row, toplam_skor, npk_toplam, toplam_gubre))

        if not skorlar:
            self.lbl_kars_sonuc.setText("Gecerli veri bulunamadi.")
            return

        # Skorlari sirala
        skorlar.sort(key=lambda x: x[1], reverse=True)

        # Elesitlik toleransi (floating point karsilastirmasi icin)
        ESITLIK_TOLERANS = 0.01
        en_yuksek_skor = skorlar[0][1]
        en_dusuk_skor = skorlar[-1][1]

        # Gruplari belirle: esit skorlu olanlari ayni gruba koy
        en_iyi_grup = []    # En yuksek skorlu tum secenekler
        en_kotu_grup = []   # En dusuk skorlu tum secenekler
        orta_grup = []      # Diger secenekler

        for item in skorlar:
            row, skor, npk, gubre = item
            if abs(skor - en_yuksek_skor) < ESITLIK_TOLERANS:
                en_iyi_grup.append(item)
            elif abs(skor - en_dusuk_skor) < ESITLIK_TOLERANS:
                en_kotu_grup.append(item)
            else:
                orta_grup.append(item)

        # Renkleri uygula
        # En iyi gruba koyu yesil
        for item in en_iyi_grup:
            for col in range(self.tbl_karsilastirma.columnCount()):
                cell = self.tbl_karsilastirma.item(item[0], col)
                if cell:
                    cell.setBackground(QColor("#C8E6C9"))
                    cell.setFont(QFont("Arial", 10, QFont.Weight.Bold))

        # Ortadaki gruplara hafif yesil
        for item in orta_grup:
            for col in range(self.tbl_karsilastirma.columnCount()):
                cell = self.tbl_karsilastirma.item(item[0], col)
                if cell:
                    cell.setBackground(QColor("#E8F5E9"))

        # En dusuk gruba kirmizi (ama en iyi ile ayni skorsa kirmizi koyma)
        if en_yuksek_skor != en_dusuk_skor:
            for item in en_kotu_grup:
                for col in range(self.tbl_karsilastirma.columnCount()):
                    cell = self.tbl_karsilastirma.item(item[0], col)
                    if cell:
                        cell.setBackground(QColor("#FFEBEE"))

        # Sonuc metni olustur
        en_iyi_sayisi = len(en_iyi_grup)
        ilk_en_iyi = en_iyi_grup[0]
        en_iyi_bolge = self.tbl_karsilastirma.item(ilk_en_iyi[0], 1).text()
        en_iyi_bitki = self.tbl_karsilastirma.item(ilk_en_iyi[0], 2).text()
        en_iyi_tarim = self.tbl_karsilastirma.item(ilk_en_iyi[0], 3).text()
        en_iyi_npk = f"{ilk_en_iyi[2]:.0f}"
        en_iyi_gubre = self.tbl_karsilastirma.item(ilk_en_iyi[0], 8).text()

        if en_iyi_sayisi == 1:
            sonuc_baslik = f"EN UYGUN SECIM:  {en_iyi_bolge} / {en_iyi_bitki} ({en_iyi_tarim})"
        else:
            bolge_listesi = []
            for item in en_iyi_grup:
                b = self.tbl_karsilastirma.item(item[0], 1).text()
                bolge_listesi.append(b)
            sonuc_baslik = f"ESIT DEGERLER - {en_iyi_sayisi} SECENEK EN IYI:  {', '.join(bolge_listesi)}"

        sonuc = (
            f"{sonuc_baslik}\n"
            f"NPK Degeri: {en_iyi_npk} kg/da  |  Gubre: {en_iyi_gubre}  |  "
            f"Toplam Gubre: {ilk_en_iyi[3]:.0f} kg/da\n\n"
            f"Karsilastirma Sonucu: {satir_sayisi} secenek degerlendirildi. "
        )

        if en_yuksek_skor == en_dusuk_skor:
            sonuc += "Tum secenekler esit degerlere sahiptir."
        elif en_iyi_sayisi > 1:
            sonuc += f"{en_iyi_sayisi} secenek ayni skora sahip (en iyi). Kirmizi ile en dusuk performansli secenek isaretlenmistir."
        else:
            sonuc += "Yesil ile en iyi, acik yesil ile orta, kirmizi ile en dusuk performansli secenek isaretlenmistir."

        self.lbl_kars_sonuc.setText(sonuc)

    # --- GECMIS ---
    def _gecmis_yenile(self):
        filtre = self.txt_arama.text().strip() if hasattr(self, 'txt_arama') else ""
        kayitlar = database.gecmisi_listele(filtre=filtre if filtre else None)
        self.tbl_gecmis.setRowCount(len(kayitlar))
        for row, kayit in enumerate(kayitlar):
            for col, deger in enumerate(kayit):
                self.tbl_gecmis.setItem(row, col, QTableWidgetItem(str(deger or "")))

    def _gecmis_temizle(self):
        cevap = QMessageBox.question(self, cevir("temizle"), "Tum gecmis kayitlari silinecek. Emin misiniz?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if cevap == QMessageBox.StandardButton.Yes:
            database.gecmisi_temizle()
            self._gecmis_yenile()

    def _gecmis_kayit_sil(self):
        secili = self.tbl_gecmis.currentRow()
        if secili >= 0:
            kayit_id = self.tbl_gecmis.item(secili, 0).text()
            database.gecmis_sil(int(kayit_id))
            self._gecmis_yenile()

    def _gecmisten_yukle(self):
        secili = self.tbl_gecmis.currentRow()
        if secili >= 0:
            bolge = self.tbl_gecmis.item(secili, 2).text()
            bitki = self.tbl_gecmis.item(secili, 3).text()
            tarim = self.tbl_gecmis.item(secili, 4).text()
            self.cmb_bolge.setCurrentText(bolge)
            self.cmb_bitki.setCurrentText(bitki)
            self.cmb_tarim.setCurrentText(tarim)

    # --- DOSYA YUKLEME ---
    def _dosya_yukle(self):
        dosya_yolu, _ = QFileDialog.getOpenFileName(
            self, cevir("dosya_sec"), "",
            f"{cevir('excel_dosyalari')} (*.xlsx *.csv);;{cevir('tum_dosyalar')} (*)"
        )
        if dosya_yolu:
            sonuclar, hata = dosya_yukle(dosya_yolu)
            if hata:
                QMessageBox.critical(self, cevir("kayit_hatasi"), hata)
                return
            if sonuclar:
                ilkbir = sonuclar[0]
                if "bolge" in ilkbir:
                    self.cmb_bolge.setCurrentText(ilkbir["bolge"])
                if "bitki" in ilkbir:
                    self.cmb_bitki.setCurrentText(ilkbir["bitki"])
                if "tarim_sekli" in ilkbir:
                    self.cmb_tarim.setCurrentText(ilkbir["tarim_sekli"])
                if "om" in ilkbir:
                    self.spin_om.setValue(ilkbir["om"])
                if "fosfor" in ilkbir:
                    self.spin_p.setValue(ilkbir["fosfor"])
                if "potasyum" in ilkbir:
                    self.spin_k.setValue(ilkbir["potasyum"])
                self.statusBar().showMessage(f"{len(sonuclar)} kayit yuklendi")

    # --- KONUM ---
    def _konum_bul(self):
        if not internet_var_mi():
            self.statusBar().showMessage(cevir("offline"))
            QMessageBox.warning(self, cevir("uyari"), cevir("konum_hatasi"))
            return

        self.statusBar().showMessage("Konum aliniyor...")
        QApplication.processEvents()
        bolge, bilgi = konumdan_bolge_bul()
        if bolge:
            self.cmb_bolge.setCurrentText(bolge)
            self.statusBar().showMessage(f"{cevir('otomatik_bolge')}: {bilgi}")
        else:
            self.statusBar().showMessage(cevir("konum_hatasi"))
            QMessageBox.warning(self, cevir("uyari"), bilgi)

    # --- E-POSTA ---
    def _eposta_gonder(self):
        if not self.sonuc_kayitlar:
            QMessageBox.information(self, cevir("uyari"), "Once bir hesaplama yapin.")
            return

        # PDF olustur ve gecici dosyaya kaydet
        gecici_pdf = os.path.join(veri_klasoru(), "gecici_rapor.pdf")
        try:
            self._pdf_olustur(gecici_pdf)
        except Exception as e:
            QMessageBox.critical(self, cevir("kayit_hatasi"), str(e))
            return

        dialog = EpostaDialog(
            konu=f"Gübre Tavsiyesi - {self.cmb_bitki.currentText()}",
            icerik=f"Sayın Yetkili,\n\n"
                   f"{self.cmb_bolge.currentText()} bölgesi, "
                   f"{self.cmb_bitki.currentText()} için gübre tavsiye raporu ekte sunulmuştur.\n\n"
                   f"Lütfen ekteki raporu inceleyiniz.\n\n"
                   f"Saygılarımızla,\n"
                   f"Gübre Tavsiyesi Uygulaması v{SURUM}\n"
                   f"Tarımsal Araştırma ve Danışmanlık",
            ek_dosya=gecici_pdf,
            ebeveyn=self
        )
        if dialog.exec():
            alici, konu, icerik = dialog.bilgileri_al()
            if not alici:
                QMessageBox.warning(self, cevir("uyari"), "Alici e-posta adresi gerekli.")
                return
            basarili, hata = eposta_gonder(
                alici, konu, icerik, gecici_pdf,
                smtp_sunucu=self.ayarlar.get("smtp_sunucu", "smtp.gmail.com"),
                smtp_port=self.ayarlar.get("smtp_port", 587),
                gonderen_email=self.ayarlar.get("smtp_kullanici", ""),
                sifre=self.ayarlar.get("smtp_sifre", "")
            )
            if basarili:
                QMessageBox.information(self, cevir("kayit_basarili"), "E-posta gonderildi.")
            else:
                QMessageBox.critical(self, cevir("kayit_hatasi"), hata)

    # --- PDF ---
    def _kaydet_pdf(self):
        if not self.sonuc_kayitlar:
            QMessageBox.information(self, cevir("uyari"), "Once bir hesaplama yapin.")
            return
        dosya_adi, _ = QFileDialog.getSaveFileName(
            self, "PDF Olarak Kaydet",
            f"gubre_bilgi_karti_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            "PDF Dosyalari (*.pdf)"
        )
        if dosya_adi:
            try:
                self._pdf_olustur(dosya_adi)
                QMessageBox.information(self, cevir("kayit_basarili"), f"PDF kaydedildi:\n{dosya_adi}")
            except Exception as e:
                QMessageBox.critical(self, cevir("kayit_hatasi"), str(e))

    def _pdf_olustur(self, dosya_yolu):
        bolge = self.cmb_bolge.currentText()
        bitki = self.cmb_bitki.currentText()
        tarim = self.cmb_tarim.currentText()
        veri = self.veri_tabani[bolge][bitki][tarim]
        om_idx = n_index(self.spin_om.value())
        p_idx = p_index(self.spin_p.value())
        k_idx = k_index(self.spin_k.value())
        n_saf = veri[om_idx]
        p_saf = veri[N_ARALIK_SAYISI + p_idx]
        k_saf = veri[N_ARALIK_SAYISI + P_ARALIK_SAYISI + k_idx]

        if n_saf > 0:
            ure_kat = INORGANIK_GUBRELER["Azot (N)"]["Ure %46 N"]["kat"]
            en_uygun_gubre = "Üre %46 N"
            en_uygun_miktar = f"{n_saf * ure_kat:.0f} kg/da"
        elif p_saf > 0:
            dap_kat = INORGANIK_GUBRELER["Fosfor (P2O5)"]["Diamonyum Fosfat (DAP) %46 P2O5"]["kat"]
            en_uygun_gubre = "DAP %46"
            en_uygun_miktar = f"{p_saf * dap_kat:.0f} kg/da"
        elif k_saf > 0:
            ps_kat = INORGANIK_GUBRELER["Potasyum (K2O)"]["Potasyum Sulfat (PS) %50 K2O"]["kat"]
            en_uygun_gubre = "Potasyum Sülfat %50"
            en_uygun_miktar = f"{k_saf * ps_kat:.0f} kg/da"
        else:
            en_uygun_gubre = "Gübre Gerekmiyor"
            en_uygun_miktar = "0 kg/da"

        c = pdf_canvas.Canvas(dosya_yolu, pagesize=A4)
        genislik, yukseklik = A4

        windows_font_yolu = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts")
        try:
            # Tahoma fontu Turkce karakterleri tam destekler
            pdfmetrics.registerFont(TTFont("Tahoma", os.path.join(windows_font_yolu, "tahoma.ttf")))
            pdfmetrics.registerFont(TTFont("TahomaBold", os.path.join(windows_font_yolu, "tahomabd.ttf")))
            fn, fb = "Tahoma", "TahomaBold"
        except:
            try:
                pdfmetrics.registerFont(TTFont("Arial", os.path.join(windows_font_yolu, "arial.ttf")))
                pdfmetrics.registerFont(TTFont("ArialBold", os.path.join(windows_font_yolu, "arialbd.ttf")))
                fn, fb = "Arial", "ArialBold"
            except:
                fn, fb = "Helvetica", "Helvetica-Bold"

        koyu_yesil = HexColor("#2E7D32")
        acik_yesil = HexColor("#4CAF50")
        cok_acik_yesil = HexColor("#E8F5E9")
        acik_gri = HexColor("#ECEFF1")
        koyu_gri = HexColor("#212121")
        orta_gri = HexColor("#757575")
        turuncu_acik = HexColor("#FFF3E0")
        turuncu = HexColor("#FF9800")
        beyaz = white

        c.setFillColor(acik_gri)
        c.rect(0, 0, genislik, yukseklik, fill=1, stroke=0)

        c.setFillColor(koyu_yesil)
        c.rect(0, yukseklik - 70, genislik, 70, fill=1, stroke=0)

        c.setFillColor(beyaz)
        c.setFont(fb, 22)
        c.drawCentredString(genislik/2, yukseklik - 35, "GÜBRE TAVSİYESİ SONUCUNUZ")
        c.setFont(fn, 12)
        c.drawCentredString(genislik/2, yukseklik - 55, "Bilgilendirme Karti")

        kart_x, kart_genislik = 40, genislik - 80
        kart_y, kart_yukseklik = yukseklik - 500, 420

        c.setFillColor(beyaz)
        c.setStrokeColor(HexColor("#E0E0E0"))
        c.roundRect(kart_x, kart_y, kart_genislik, kart_yukseklik, 10, fill=1, stroke=1)

        y = kart_y + kart_yukseklik - 40
        c.setFillColor(koyu_gri)
        c.setFont(fb, 16)
        c.drawString(kart_x + 20, y, bitki)
        c.setFont(fn, 10)
        c.setFillColor(orta_gri)
        c.drawString(kart_x + 20, y - 18, f"{bolge} Bölgesi  |  {tarim} Tarımı")
        c.setFont(fn, 9)
        c.drawRightString(kart_x + kart_genislik - 15, y, f"Tarih: {datetime.now().strftime('%d.%m.%Y')}")

        y -= 35
        c.setStrokeColor(HexColor("#E0E0E0"))
        c.line(kart_x + 20, y, kart_x + kart_genislik - 20, y)

        y -= 20
        badge_w = 120
        c.setFillColor(koyu_yesil)
        c.roundRect((genislik - badge_w)/2, y - 8, badge_w, 18, 9, fill=1, stroke=0)
        c.setFillColor(beyaz)
        c.setFont(fb, 8)
        c.drawCentredString(genislik/2, y - 4, "ÖNERİLEN GÜBRE")

        y -= 30
        gubre_kutu_y = y - 90
        c.setFillColor(cok_acik_yesil)
        c.roundRect(kart_x + 20, gubre_kutu_y, kart_genislik - 40, 90, 5, fill=1, stroke=0)
        c.setFillColor(koyu_yesil)
        c.roundRect(kart_x + 20, gubre_kutu_y, 8, 90, 4, fill=1, stroke=0)

        c.setFillColor(koyu_gri)
        c.setFont(fb, 14)
        c.drawString(kart_x + 40, gubre_kutu_y + 60, en_uygun_gubre)
        c.setFillColor(orta_gri)
        c.setFont(fn, 9)
        c.drawString(kart_x + 40, gubre_kutu_y + 42, "Kompoze Gübre" if n_saf > 0 and p_saf > 0 else "Gübre")

        npk_text = f"{n_saf}-{p_saf}-{k_saf}"
        c.setFillColor(koyu_yesil)
        c.setFont(fb, 26)
        c.drawRightString(kart_x + kart_genislik - 30, gubre_kutu_y + 65, npk_text)
        c.setFillColor(orta_gri)
        c.setFont(fn, 7)
        c.drawRightString(kart_x + kart_genislik - 30, gubre_kutu_y + 48, "kg/da  (N - P2O5 - K2O)")

        y = gubre_kutu_y - 15
        kart_en = (kart_genislik - 60) / 3
        for i, (etiket, deger) in enumerate([
            ("AZOT (N)", f"{n_saf} kg/da"), ("FOSFOR (P2O5)", f"{p_saf} kg/da"), ("POTASYUM (K2O)", f"{k_saf} kg/da")
        ]):
            kx = kart_x + 20 + i * (kart_en + 10)
            c.setFillColor(cok_acik_yesil)
            c.roundRect(kx, y - 50, kart_en, 55, 5, fill=1, stroke=0)
            c.setFillColor(koyu_yesil)
            c.setFont(fb, 7)
            c.drawCentredString(kx + kart_en/2, y - 28, etiket)
            c.setFillColor(koyu_gri)
            c.setFont(fb, 13)
            c.drawCentredString(kx + kart_en/2, y - 45, deger)

        y -= 75
        detay_y = y - 60
        c.setStrokeColor(koyu_yesil)
        c.setLineWidth(2)
        c.setFillColor(beyaz)
        c.roundRect(kart_x + 20, detay_y, kart_genislik - 40, 65, 5, fill=1, stroke=1)
        c.setFillColor(koyu_yesil)
        c.setFont(fb, 12)
        c.drawString(kart_x + 40, detay_y + 40, en_uygun_gubre)
        c.setFillColor(orta_gri)
        c.setFont(fn, 9)
        c.drawString(kart_x + 40, detay_y + 20, f"Önerilen: {en_uygun_miktar}")
        c.setFillColor(koyu_yesil)
        c.setFont(fb, 16)
        c.drawRightString(kart_x + kart_genislik - 30, detay_y + 40, en_uygun_miktar)

        y = detay_y - 15
        for i, (etiket, deger) in enumerate([
            ("UYGULAMA DOZU", en_uygun_miktar), ("UYGULAMA ZAMANI", "Ekimden Önce"), ("UYGULAMA ŞEKLİ", "Toprağa Serpme")
        ]):
            kx = kart_x + 20 + i * (kart_en + 10)
            c.setFillColor(cok_acik_yesil)
            c.roundRect(kx, y - 50, kart_en, 55, 5, fill=1, stroke=0)
            c.setFillColor(koyu_yesil)
            c.setFont(fb, 7)
            c.drawCentredString(kx + kart_en/2, y - 28, etiket)
            c.setFillColor(koyu_gri)
            c.setFont(fb, 10)
            c.drawCentredString(kx + kart_en/2, y - 45, deger)

        y -= 75
        uyari_y = y - 55
        c.setFillColor(turuncu_acik)
        c.setStrokeColor(turuncu)
        c.setLineWidth(1.5)
        c.roundRect(kart_x + 20, uyari_y, kart_genislik - 40, 60, 5, fill=1, stroke=1)
        c.setFillColor(koyu_yesil)
        c.circle(kart_x + 45, uyari_y + 30, 14, fill=1, stroke=0)
        c.setFillColor(beyaz)
        c.setFont(fb, 16)
        c.drawCentredString(kart_x + 45, uyari_y + 24, "i")
        c.setFillColor(koyu_gri)
        c.setFont(fb, 9)
        c.drawString(kart_x + 65, uyari_y + 38, "Daha verimli sonuçlar için toprağınızın düzenli analizini yaptırmayı")
        c.setFont(fn, 9)
        c.drawString(kart_x + 65, uyari_y + 22, "ve önerilen doza uymayı unutmayınız.")
        c.setFont(fn, 8)
        c.setFillColor(orta_gri)
        c.drawString(kart_x + 65, uyari_y + 8, "Her bitkinin ihtiyacı farklıdır. Komşu tarlalardaki dozajları kopyalamayınız.")

        # --- DIKKAT VE SORUMLULUK REDDI BEYANI (Buyuk ve dikkat cezbedici) ---
        uyaribeyan_y = kart_y - 20
        uyaribeyan_h = 160
        uyaribeyan_x = kart_x + 25
        uyaribeyan_w = kart_genislik - 50

        # Kirmizi kenarlikli dikkat kutusu
        c.setFillColor(HexColor("#FFF3E0"))
        c.setStrokeColor(HexColor("#D32F2F"))
        c.setLineWidth(3)
        c.roundRect(uyaribeyan_x, uyaribeyan_y - uyaribeyan_h, uyaribeyan_w, uyaribeyan_h, 8, fill=1, stroke=1)

        # Kirmizi ust bant
        c.setFillColor(HexColor("#D32F2F"))
        c.roundRect(uyaribeyan_x, uyaribeyan_y - 32, uyaribeyan_w, 32, 8, fill=1, stroke=0)
        c.rect(uyaribeyan_x, uyaribeyan_y - 32, uyaribeyan_w, 16, fill=1, stroke=0)

        # Dikkat basligi - buyuk ve net
        c.setFillColor(beyaz)
        c.setFont(fb, 14)
        c.drawCentredString(genislik/2, uyaribeyan_y - 24, "DİKKAT: BU BELGE GENEL BİLGİLENDİRME AMAÇLIDIR")

        # Uyari ikonu (buyuk kirmizi daire + !)
        uyari_ikon_x = uyaribeyan_x + 24
        uyari_ikon_y = uyaribeyan_y - 65
        c.setFillColor(HexColor("#D32F2F"))
        c.circle(uyari_ikon_x, uyari_ikon_y, 16, fill=1, stroke=0)
        c.setFillColor(beyaz)
        c.setFont(fb, 18)
        c.drawCentredString(uyari_ikon_x, uyari_ikon_y - 6, "!")

        # Uyari metni - buyuk ve okunakli (kisa ve net satirlar)
        c.setFillColor(HexColor("#212121"))
        c.setFont(fn, 11)
        metin_x = uyaribeyan_x + 50
        c.drawString(metin_x, uyaribeyan_y - 52,
                     "Kesin gübre dozajları için bölgenizdeki Tarım ve Orman")
        c.drawString(metin_x, uyaribeyan_y - 66,
                     "İl/İlçe Müdürlüğü'ne danışınız.")
        c.drawString(metin_x, uyaribeyan_y - 80,
                     "Toprak analiz sonuçlarına göre yerel ayarlama gerekebilir.")

        # Ayirici cizgi
        c.setStrokeColor(HexColor("#E0E0E0"))
        c.setLineWidth(0.8)
        c.line(uyaribeyan_x + 50, uyaribeyan_y - 92, uyaribeyan_x + uyaribeyan_w - 25, uyaribeyan_y - 92)

        # Sorumluluk reddi - buyuk yazi
        c.setFillColor(HexColor("#616161"))
        c.setFont(fn, 9)
        c.drawString(uyaribeyan_x + 50, uyaribeyan_y - 108,
                     "Sorumluluk Reddi: Bu uygulama ve çıktıkları genel rehber")
        c.drawString(uyaribeyan_x + 50, uyaribeyan_y - 122,
                     "niteliğindedir. Uygulamadan kaynaklanabilecek zarardan")
        c.drawString(uyaribeyan_x + 50, uyaribeyan_y - 136,
                     "program yazıcı, geliştirici ve dağıtıcıları sorumlu tutulamaz.")

        # Kucuk not
        c.setFillColor(HexColor("#9E9E9E"))
        c.setFont(fn, 7.5)
        c.drawString(uyaribeyan_x + 50, uyaribeyan_y - 152,
                     "Bu belge profesyonel tarımsal danışmanlık yerine geçmez.")

        # --- FOOTER ---
        footer_y = uyaribeyan_y - uyaribeyan_h - 25
        c.setStrokeColor(acik_yesil)
        c.setDash(3, 2)
        c.line(kart_x + 50, footer_y + 15, kart_x + kart_genislik - 50, footer_y + 15)
        c.setDash()
        c.setFillColor(koyu_yesil)
        c.setFont(fb, 10)
        c.drawCentredString(genislik/2, footer_y, cevir("verimli_hasat"))
        c.setFillColor(orta_gri)
        c.setFont(fn, 7)
        c.drawCentredString(genislik/2, footer_y - 15, f"TARIM-GUBRE v{SURUM}  |  {SURUM_TARIHI}  |  {YAZILIMCI}")

        c.save()

    def _hakkinda_goster(self):
        msg = QMessageBox(self)
        msg.setWindowTitle(cevir("hakkinda"))
        msg.setTextFormat(Qt.TextFormat.RichText)
        msg.setText(f"""
        <h2 style="color: #2E7D32;">TARIM-GUBRE v{SURUM}</h2>
        <p><b>{cevir('uygulama_hakkinda')}</b></p>
        <hr>
        <p><b>Kisayol Tuslari:</b></p>
        <ul>
            <li>Ctrl+H - Hesapla</li>
            <li>Ctrl+P - PDF Kaydet</li>
            <li>Ctrl+E - E-posta Gonder</li>
            <li>Ctrl+O - Dosya Yukle</li>
            <li>Ctrl+G - Gecmis</li>
            <li>Ctrl+L - Konum Bul</li>
            <li>Ctrl+T - Tema Degistir</li>
            <li>Ctrl+D - Dil Degistir</li>
        </ul>
        <hr>
        <p style="color: #D32F2F;">{cevir('sorumluluk_reddi')}</p>
        <p style="color: #2E7D32;">Programlayan: {YAZILIMCI}</p>
        <p style="color: #757575;">Copyright 2026 Her Haksi Saklidir!</p>
        """)
        msg.setIcon(QMessageBox.Icon.Information)
        msg.exec()


# =============================================================================
# 6. BASLATMA
# =============================================================================

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    icon_yolu = kaynak_yolu("app.ico")
    if os.path.exists(icon_yolu):
        app.setWindowIcon(QIcon(icon_yolu))

    if sys.platform == "win32":
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("com.tarim.gubre.tavsiye.v3")
        except Exception:
            pass

    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor("#F5F5F5"))
    palette.setColor(QPalette.ColorRole.WindowText, QColor("#212121"))
    palette.setColor(QPalette.ColorRole.Base, QColor("#FFFFFF"))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor("#F1F8E9"))
    palette.setColor(QPalette.ColorRole.Text, QColor("#212121"))
    palette.setColor(QPalette.ColorRole.Button, QColor("#E8E8E8"))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor("#212121"))
    palette.setColor(QPalette.ColorRole.Highlight, QColor("#2E7D32"))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#FFFFFF"))
    app.setPalette(palette)

    window = TarimApp()
    window.show()
    sys.exit(app.exec())
