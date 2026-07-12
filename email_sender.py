"""E-posta gonderme modulu - SMTP destegi."""
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os


def eposta_gonder(alici, konu, icerik, ek_dosya=None,
                  smtp_sunucu="smtp.gmail.com", smtp_port=587,
                 gonderen_email="", sifre=""):
    """E-posta gonderir.

    Args:
        alici: Alici e-posta adresi
        konu: E-posta konusu
        icerik: E-posta icerigi (HTML destekli)
        ek_dosya: Eklenecek dosya yolu (opsiyonel)
        smtp_sunucu: SMTP sunucu adresi
        smtp_port: SMTP port numarasi
        gonderen_email: Gonderen e-posta adresi
        sifre: E-posta sifresi

    Returns:
        (basarili, hata_mesaji)
    """
    if not gonderen_email or not sifre:
        return False, "E-posta ayarlari yapilmamis"

    try:
        msg = MIMEMultipart()
        msg["From"] = gonderen_email
        msg["To"] = alici
        msg["Subject"] = konu

        # HTML icerik
        html_icerik = f"""
        <html>
        <body style="font-family: Arial, sans-serif; color: #333;">
            <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
                <div style="background-color: #2E7D32; color: white; padding: 15px; text-align: center; border-radius: 8px 8px 0 0;">
                    <h2 style="margin: 0;">Gubre Tavsiyesi Raporu</h2>
                </div>
                <div style="background-color: #f9f9f9; padding: 20px; border: 1px solid #ddd;">
                    {icerik}
                </div>
                <div style="background-color: #2E7D32; color: white; padding: 10px; text-align: center; border-radius: 0 0 8px 8px; font-size: 12px;">
                    Gubre Tavsiyesi Uygulamasi v3.0
                </div>
            </div>
        </body>
        </html>
        """
        msg.attach(MIMEText(html_icerik, "html", "utf-8"))

        # Ekleme varsa
        if ek_dosya and os.path.exists(ek_dosya):
            with open(ek_dosya, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            dosya_adi = os.path.basename(ek_dosya)
            part.add_header("Content-Disposition", f"attachment; filename={dosya_adi}")
            msg.attach(part)

        # SMTP baglantisi
        with smtplib.SMTP(smtp_sunucu, smtp_port) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(gonderen_email, sifre)
            server.sendmail(gonderen_email, alici, msg.as_string())

        return True, None

    except smtplib.SMTPAuthenticationError:
        return False, "E-posta dogrulamasi basarisiz. Sifrenizi kontrol edin."
    except smtplib.SMTPConnectError:
        return False, "SMTP sunucusuna baglanamadi."
    except Exception as e:
        return False, f"E-posta gonderilemedi: {str(e)}"


def smtp_test_et(smtp_sunucu, smtp_port, kullanici, sifre):
    """SMTP baglantisini test eder."""
    try:
        with smtplib.SMTP(smtp_sunucu, smtp_port) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(kullanici, sifre)
        return True, "Baglanti basarili"
    except smtplib.SMTPAuthenticationError:
        return False, "Kullanici adi veya sifre hatali"
    except smtplib.SMTPConnectError:
        return False, "Sunucuya baglanamiyor"
    except Exception as e:
        return False, str(e)
