# Medar Yaka Kart Otomasyonu (v3)

PDF (ön/arka) kart şablonlarını alıp A4’e **dupleks (Long Edge)** baskıya uygun şekilde Word çıktısı (`.docx`) üreten masaüstü araç.

## Özellikler
- PDF seçimi + ZIP/RAR içinden PDF çıkarma
- Dosya listesi: sıralama / çoklu silme
- Önizleme (ön/arka)
- Profil kaydetme/yükleme
- DPI & sayfa başı kart ayarları
- Tema (Açık/Koyu)
- İstatistik paneli

## Kurulum (Geliştirme)
Python 3.10+ önerilir.

```bash
pip install -r requirements.txt
python -m medar_yakakart
```

> Not: Sürükle-bırak için opsiyonel olarak `tkinterdnd2` kurulabilir.

## Kullanım
1. **PDF Seç** veya **ZIP/RAR Aç** ile kart dosyalarını ekleyin  
2. Ayarları seçin (profil / DPI / kart boyutu / kenar boşlukları)  
3. **Kimlikleri Oluştur** ile `.docx` çıktıyı alın  
4. Yazıcı ayarı: Dupleks = Açık, Flip = **Long Edge**, Ölçek = %100, Kağıt = A4

## Yapılandırma Dosyaları
Uygulama çalışırken aynı klasöre aşağıdaki dosyaları oluşturur:
- `config.json`
- `profiles.json`
- `stats.json`

Repo’da örnekleri mevcut:
- `config.example.json`
- `profiles.example.json`
- `stats.example.json`

## EXE (PyInstaller) - Önerilen
Detaylar: `docs/BUILD.md`
