# Denetim - Excel Çok Sayfalı Rapor Oluşturucu

Excel dosyalarındaki farklı sayfalardan veri çekerek kapsamlı raporlar oluşturan Python uygulaması.

## Özellikler

- ✅ Excel dosyasındaki tüm sayfaları otomatik olarak okuma
- ✅ Belirli sayfaları seçerek okuma
- ✅ Farklı sayfaları birleştirme (dikey veya yatay)
- ✅ Özet raporlar oluşturma
- ✅ İstatistiksel analiz raporları
- ✅ Çıktıları Excel formatında kaydetme

## Kurulum

1. Gerekli Python paketlerini yükleyin:

```bash
pip install -r requirements.txt
```

## Kullanım

### 1. Örnek Excel Dosyası Oluşturma

Test için örnek bir Excel dosyası oluşturun:

```bash
python ornek_veri_olustur.py
```

Bu komut `ornek_veri.xlsx` dosyasını oluşturur (3 farklı sayfa içerir).

### 2. Rapor Oluşturma

Ana programı çalıştırın:

```bash
# Tüm örnekleri çalıştır
python main.py ornek_veri.xlsx

# Sadece özet rapor oluştur (Örnek 1)
python main.py ornek_veri.xlsx 1

# Sayfaları birleştir (Örnek 2)
python main.py ornek_veri.xlsx 2

# İstatistiksel analiz yap (Örnek 3)
python main.py ornek_veri.xlsx 3
```

### 3. Programatik Kullanım

Kendi Python kodunuzda kullanmak için:

```python
from excel_reader import ExcelReader
from rapor_olusturucu import RaporOlusturucu

# Excel dosyasını oku
okuyucu = ExcelReader('verilerim.xlsx')

# Tüm sayfaları oku
tum_veriler = okuyucu.tum_sayfalari_oku()

# Rapor oluştur
rapor = RaporOlusturucu(tum_veriler)

# Özet rapor kaydet
rapor.rapor_kaydet('cikti_rapor.xlsx', rapor_turu='ozet')
```

## Modüller

### ExcelReader

Excel dosyalarından veri okuma modülü:

- `sayfa_oku(sayfa_adi)` - Tek bir sayfayı oku
- `tum_sayfalari_oku()` - Tüm sayfaları oku
- `belirli_sayfalari_oku(sayfa_isimleri)` - Belirtilen sayfaları oku
- `sayfa_bilgisi_al(sayfa_adi)` - Sayfa hakkında özet bilgi al

### RaporOlusturucu

Raporlama işlemleri modülü:

- `ozet_rapor_olustur()` - Tüm sayfalar için özet oluştur
- `sayfalari_birlestir()` - Sayfaları birleştir
- `istatistik_rapor_olustur()` - İstatistiksel analiz
- `rapor_kaydet()` - Raporu Excel olarak kaydet
- `konsol_rapor_goster()` - Konsola rapor yazdır

## Rapor Türleri

### Özet Rapor
Tüm sayfaların özetini ve detaylarını içeren rapor:
```python
rapor.rapor_kaydet('ozet.xlsx', rapor_turu='ozet')
```

### Birleşik Rapor
Sayfaları birleştirerek tek bir veri seti oluşturma:
```python
# Dikey birleştirme (alt alta)
rapor.rapor_kaydet('birlesik.xlsx', rapor_turu='birlesik', birlesim_turu='vertical')

# Yatay birleştirme (ortak sütun ile)
rapor.rapor_kaydet('birlesik.xlsx', rapor_turu='birlesik', 
                   birlesim_turu='horizontal', ortak_sutun='ID')
```

### Ayrı Sayfalar
Sadece seçili sayfaları içeren rapor:
```python
rapor.rapor_kaydet('secili.xlsx', rapor_turu='ayri_sayfalar', 
                   sayfa_isimleri=['Sayfa1', 'Sayfa2'])
```

## Çıktılar

Oluşturulan raporlar `output/` klasörüne kaydedilir:
- `output/ozet_rapor.xlsx` - Özet rapor
- `output/birlesik_rapor.xlsx` - Birleşik rapor

## Gereksinimler

- Python 3.7+
- pandas 2.1.4
- openpyxl 3.1.2

## Lisans

MIT
