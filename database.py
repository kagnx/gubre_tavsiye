"""SQLite veritabani islemleri - Gecmis kayitlari."""
import sqlite3
import os
from datetime import datetime
from config import veri_klasoru

DB_YOL = os.path.join(veri_klasoru(), "gecmis.db")


def baglanti_olustur():
    """Veritabani baglantisi olusturur."""
    conn = sqlite3.connect(DB_YOL)
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def tablolari_olustur():
    """Gerekli tablolari olusturur."""
    conn = baglanti_olustur()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS hesaplamalar (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tarih TEXT NOT NULL,
            bolge TEXT NOT NULL,
            bitki TEXT NOT NULL,
            tarim_sekli TEXT NOT NULL,
            om REAL,
            fosfor REAL,
            potasyum REAL,
            n_saf REAL,
            p_saf REAL,
            k_saf REAL,
            gubre_onerisi TEXT,
            toplam_miktar TEXT,
            notlar TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS karsilastirmalar (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tarih TEXT NOT NULL,
            baslik TEXT NOT NULL,
            senaryolar TEXT NOT NULL
        )
    """)
    conn.commit()
    conn.close()


def hesaplama_kaydet(bolge, bitki, tarim_sekli, om, fosfor, potasyum,
                      n_saf, p_saf, k_saf, gubre_onerisi, toplam_miktar, notlar=""):
    """Hesaplama sonucunu kaydeder."""
    conn = baglanti_olustur()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO hesaplamalar
        (tarih, bolge, bitki, tarim_sekli, om, fosfor, potasyum,
         n_saf, p_saf, k_saf, gubre_onerisi, toplam_miktar, notlar)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        bolge, bitki, tarim_sekli, om, fosfor, potasyum,
        n_saf, p_saf, k_saf, gubre_onerisi, toplam_miktar, notlar
    ))
    conn.commit()
    conn.close()


def gecmisi_listele(limit=100, filtre=None):
    """Gecmis kayitlarini listeler."""
    conn = baglanti_olustur()
    cursor = conn.cursor()
    if filtre:
        cursor.execute("""
            SELECT * FROM hesaplamalar
            WHERE bolge LIKE ? OR bitki LIKE ? OR tarih LIKE ?
            ORDER BY id DESC LIMIT ?
        """, (f"%{filtre}%", f"%{filtre}%", f"%{filtre}%", limit))
    else:
        cursor.execute("""
            SELECT * FROM hesaplamalar ORDER BY id DESC LIMIT ?
        """, (limit,))
    sonuclar = cursor.fetchall()
    conn.close()
    return sonuclar


def gecmis_sil(kayit_id):
    """Belirli bir gecmis kaydini siler."""
    conn = baglanti_olustur()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM hesaplamalar WHERE id = ?", (kayit_id,))
    conn.commit()
    conn.close()


def gecmisi_temizle():
    """Tum gecmis kayitlarini siler."""
    conn = baglanti_olustur()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM hesaplamalar")
    conn.commit()
    conn.close()


def karsilastirma_kaydet(baslik, senaryolar_json):
    """Karsilastirma sonucunu kaydeder."""
    conn = baglanti_olustur()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO karsilastirmalar (tarih, baslik, senaryolar)
        VALUES (?, ?, ?)
    """, (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), baslik, senaryolar_json))
    conn.commit()
    conn.close()


def karsilastirmalari_listele(limit=50):
    """Karsilastirma kayitlarini listeler."""
    conn = baglanti_olustur()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT * FROM karsilastirmalar ORDER BY id DESC LIMIT ?
    """, (limit,))
    sonuclar = cursor.fetchall()
    conn.close()
    return sonuclar


# Baslat
tablolari_olustur()
