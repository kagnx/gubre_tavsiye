# TARIM-GÜBRE v3.0

**Türkiye Profesyonel Gübre Tavsiye Sistemi**

PyQt6 ile geliştirilmiş, profesyonel gübre tavsiyeleri sunan masaüstü uygulaması.

---

## 🌾 Özellikler

### Temel Özellikler
- **9 Bölge Desteği**: Trakya, Marmara, Karadeniz, Orta Anadolu, Akdeniz, Adalar, Güneydoğu Anadolu, Doğu Anadolu, Göller
- **55+ Bitki Türü**: Buğday, Arpa, Mısır, Ayçiçeği, Şeker Pancarı, Pamuk, Domates, Fındık, Zeytin ve daha fazlası
- **Kuru/Sulu Yetiştirme**: Her bitki için hem kuru hem sulu yetiştirme tekniği desteği
- **Detaylı Gübreleme Rehberi**: Her bitki için zamanlama, uygulama şekli, dozaj bilgileri

### Hesaplama Özellikleri
- **Kimyasal (Inorganik) Gübre Önerileri**: Azot, Fosfor, Potasyum ihtiyaç hesaplama
- **Organik Gübre Önerileri**: Çiftlik gübresi, tavuk gübresi, kompost önerileri
- **NPK Dağılım Grafikleri**: Pasta ve çubuk grafikleri ile görsel analiz
- **Karşılaştırmalı Analiz**: Farklı senaryoları karşılaştırma ve en iyi seçimi belirleme

### Uygulama Özellikleri
- **Tema Desteği**: Açık ve Karanlık tema
- **Dil Desteği**: Türkçe ve İngilizce
- **PDF Raporlama**: Profesyonel görünümlü PDF çıktıları
- **E-posta Gönderme**: SMTP ile rapor e-postası gönderme
- **Geçmiş Kayıtları**: SQLite veritabanında tüm hesaplama geçmişi
- **Otomatik Kaydetme**: Son seçimlerin otomatik kaydedilmesi
- **Kısayol Tuşları**: Hızlı erişim için kısayol tuşları
- **Konum Tespiti**: IP tabanlı otomatik bölge tespiti
- **Excel/CSV Yükleme**: Dışarıdan veri yükleme desteği

---

## 📁 Proje Yapısı

```
gubre_tavsiye/
├── gubre_tavsiye.py      # Ana uygulama
├── config.py             # Uygulama ayarları
├── database.py           # SQLite veritabanı işlemleri
├── translator.py         # Çeviri sistemi
├── data_loader.py        # Excel/CSV yükleme
├── email_sender.py       # E-posta gönderme
├── location.py           # Konum tespiti
├── charts.py             # Grafik çizimleri
├── translations/         # Dil dosyaları
│   ├── tr.json
│   └── en.json
├── app.ico               # Uygulama ikonu
├── app.png               # Uygulama görseli
└── dist/                 # Derlenmiş EXE
```

---

## 🚀 Kurulum

### Gereksinimler
- Python 3.10+
- pip (Python paket yöneticisi)

### Bağımlılıkların Yüklenmesi

```bash
pip install PyQt6 reportlab openpyxl matplotlib geopy
```

### Uygulamanın Çalıştırılması

```bash
python gubre_tavsiye.py
```

---

## 📦 EXE Oluşturma

PyInstaller ile EXE oluşturmak için:

```bash
pip install pyinstaller
pyinstaller --clean gubre_tavsiyesi.spec
```

Oluşturulan EXE: `dist/gubre_tavsiyesi.exe`

---

## ⌨️ Kısayol Tuşları

| Kısayol | İşlev |
|---------|-------|
| `Ctrl+H` | Hesapla |
| `Ctrl+P` | PDF Kaydet |
| `Ctrl+E` | E-posta Gönder |
| `Ctrl+O` | Dosya Yükle |
| `Ctrl+G` | Geçmiş |
| `Ctrl+L` | Konum Bul |
| `Ctrl+T` | Tema Değiştir |
| `Ctrl+D` | Dil Değiştir |

---

## 📊 Kullanım Kılavuzu

1. **Bölge Seçimi**: İklim/Bölge dropdown menüsünden bölgenizi seçin
2. **Bitki Seçimi**: Bitki/Ürün dropdown menüsünden bitkinizi seçin
3. **Yetiştirme Tekniği**: Kuru veya Sulu yetiştirme tekniğini seçin
4. **Toprak Analizi**: Organik madde, fosfor ve potasyum değerlerini girin
5. **Hesapla**: "GÜBRE TAVSİYESİ OLUŞTUR" butonuna tıklayın
6. **Sonuçları İnceleyin**: Kimyasal, organik gübre önerilerini ve grafikleri inceleyin
7. **PDF Kaydedin**: İsterseniz raporu PDF olarak kaydedin

---

## 🔧 Teknik Detaylar

- **GUI Framework**: PyQt6
- **PDF Oluşturma**: ReportLab
- **Grafikler**: Matplotlib
- **Veritabanı**: SQLite
- **Veri Analizi**: OpenPyXL
- **Konum Tespiti**: IP tabanlı (ip-api.com, ipinfo.io, ipwho.is)

---

## 📄 Lisans

Bu proje telif hakkı ile korunmaktadır. İzinsiz kopyalanması, dağıtılması veya değiştirilmesi yasaktır.

---

## 👨‍💻 Geliştirici

**Oğuz Kaan FIRAT**

---

## � Katkıda Bulunma

Bu proje kişisel bir projedir. Katkıda bulunmak için lütfen geliştirici ile iletişime geçin.

---

## 📞 İletişim

Sorularınız veya önerileriniz için geliştirici ile iletişime geçebilirsiniz.

---

*Son güncelleme: Temmuz 2026*
