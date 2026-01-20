# Excel Rapor Oluşturucu - Kullanım Kılavuzu

## Proje Özeti

Bu proje, Excel dosyalarındaki farklı sayfalardan veri çekerek kapsamlı raporlar oluşturan bir Python uygulamasıdır.

## Kurulum

```bash
# Bağımlılıkları yükle
pip install -r requirements.txt
```

## Hızlı Başlangıç

### 1. Örnek Veri Oluşturma

```bash
python ornek_veri_olustur.py
```

Bu komut `ornek_veri.xlsx` dosyasını oluşturur. Dosya 3 sayfa içerir:
- **Çalışanlar**: Çalışan bilgileri (ID, Ad, Departman, Maaş, Başlangıç Tarihi)
- **Satışlar**: Satış kayıtları (Çalışan ID, Ürün, Miktar, Fiyat, Tarih)
- **Departmanlar**: Departman özeti (Departman, Çalışan Sayısı, Ortalama Maaş, Bütçe)

### 2. Rapor Oluşturma

```bash
# Tüm örnekleri çalıştır
python main.py ornek_veri.xlsx

# Sadece özet rapor
python main.py ornek_veri.xlsx 1

# Sayfaları birleştir
python main.py ornek_veri.xlsx 2

# İstatistiksel analiz
python main.py ornek_veri.xlsx 3
```

## Modül Kullanımı

### Excel Okuyucu (ExcelReader)

```python
from excel_reader import ExcelReader

# Excel dosyasını aç
okuyucu = ExcelReader('dosya.xlsx')

# Tek bir sayfa oku
df = okuyucu.sayfa_oku('Sayfa1')

# Tüm sayfaları oku
tum_veriler = okuyucu.tum_sayfalari_oku()

# Belirli sayfaları oku
secili_veriler = okuyucu.belirli_sayfalari_oku(['Sayfa1', 'Sayfa2'])

# Sayfa bilgisi al
bilgi = okuyucu.sayfa_bilgisi_al('Sayfa1')
```

### Rapor Oluşturucu (RaporOlusturucu)

```python
from rapor_olusturucu import RaporOlusturucu

# Rapor oluşturucu başlat
rapor = RaporOlusturucu(tum_veriler)

# Özet rapor
ozet = rapor.ozet_rapor_olustur()

# Sayfaları birleştir (dikey)
birlesik = rapor.sayfalari_birlestir(birlesim_turu='vertical')

# Sayfaları birleştir (yatay - ortak sütun ile)
birlesik = rapor.sayfalari_birlestir(
    birlesim_turu='horizontal',
    ortak_sutun='ID'
)

# İstatistiksel rapor
istatistikler = rapor.istatistik_rapor_olustur('Sayfa1')

# Konsola yazdır
rapor.konsol_rapor_goster()

# Excel olarak kaydet
rapor.rapor_kaydet('cikti.xlsx', rapor_turu='ozet')
```

## Rapor Türleri

### 1. Özet Rapor
Tüm sayfaların özetini ve detaylarını içerir.

```python
rapor.rapor_kaydet('ozet.xlsx', rapor_turu='ozet')
```

Çıktı:
- İlk sayfa: Tüm sayfaların özeti (satır/sütun sayıları)
- Sonraki sayfalar: Her sayfanın tam verisi

### 2. Birleşik Rapor
Sayfaları tek bir veri seti olarak birleştirir.

```python
# Dikey birleştirme (alt alta)
rapor.rapor_kaydet(
    'birlesik.xlsx',
    rapor_turu='birlesik',
    birlesim_turu='vertical'
)

# Yatay birleştirme (ortak sütun ile)
rapor.rapor_kaydet(
    'birlesik.xlsx',
    rapor_turu='birlesik',
    birlesim_turu='horizontal',
    ortak_sutun='ID'
)
```

### 3. Ayrı Sayfalar
Sadece seçili sayfaları içerir.

```python
rapor.rapor_kaydet(
    'secili.xlsx',
    rapor_turu='ayri_sayfalar',
    sayfa_isimleri=['Sayfa1', 'Sayfa2']
)
```

## Gelişmiş Özellikler

### Özel Pandas Parametreleri

```python
# Belirli satırları atla
df = okuyucu.sayfa_oku('Sayfa1', skiprows=2)

# Belirli sütunları oku
df = okuyucu.sayfa_oku('Sayfa1', usecols=['Ad', 'Maaş'])

# Veri tiplerini belirt
df = okuyucu.sayfa_oku('Sayfa1', dtype={'ID': int, 'Maaş': float})
```

### İstatistiksel Analiz

```python
# Tüm sayısal sütunlar için
istatistikler = rapor.istatistik_rapor_olustur('Sayfa1')

# Belirli sütunlar için
istatistikler = rapor.istatistik_rapor_olustur(
    'Sayfa1',
    sayisal_sutunlar=['Maaş', 'Yaş']
)
```

## Çıktılar

Tüm çıktı dosyaları `output/` dizinine kaydedilir:
- `output/ozet_rapor.xlsx`
- `output/birlesik_rapor.xlsx`

## Hata Yönetimi

Program aşağıdaki durumlarda uygun hata mesajları verir:

- Excel dosyası bulunamadığında
- Belirtilen sayfa mevcut olmadığında
- Birleştirme için ortak sütun eksik olduğunda
- Veri okuma sırasında hata oluştuğunda

## Örnek Senaryolar

### Senaryo 1: Farklı Aylara Ait Satış Verilerini Birleştirme

```python
from excel_reader import ExcelReader
from rapor_olusturucu import RaporOlusturucu

# Excel'i oku (Ocak, Şubat, Mart sayfaları var)
okuyucu = ExcelReader('satislar.xlsx')
veriler = okuyucu.belirli_sayfalari_oku(['Ocak', 'Şubat', 'Mart'])

# Rapor oluştur ve birleştir
rapor = RaporOlusturucu(veriler)
yillik_rapor = rapor.sayfalari_birlestir(birlesim_turu='vertical')

# Kaydet
rapor.rapor_kaydet('yillik_satis.xlsx', rapor_turu='birlesik', birlesim_turu='vertical')
```

### Senaryo 2: Çalışan ve Departman Verilerini Eşleştirme

```python
# Çalışanlar ve Departmanlar sayfalarını oku
okuyucu = ExcelReader('firma.xlsx')
veriler = okuyucu.belirli_sayfalari_oku(['Çalışanlar', 'Departmanlar'])

# Departman sütunu üzerinden birleştir
rapor = RaporOlusturucu(veriler)
rapor.rapor_kaydet(
    'eslesmis_veri.xlsx',
    rapor_turu='birlesik',
    birlesim_turu='horizontal',
    ortak_sutun='Departman'
)
```

### Senaryo 3: Tüm Sayfalar İçin Özet İstatistik

```python
# Tüm sayfaları oku
okuyucu = ExcelReader('veriler.xlsx')
tum_veriler = okuyucu.tum_sayfalari_oku()

# Her sayfa için istatistik oluştur
rapor = RaporOlusturucu(tum_veriler)
for sayfa_adi in tum_veriler.keys():
    print(f"\n{sayfa_adi} İstatistikleri:")
    istatistikler = rapor.istatistik_rapor_olustur(sayfa_adi)
    print(istatistikler)
```

## Sık Sorulan Sorular

**S: Çok büyük Excel dosyaları için performans nasıl?**
C: Pandas optimize edilmiştir ancak çok büyük dosyalar için chunk okuma kullanabilirsiniz.

**S: Birleştirme sırasında sütun isimleri çakışırsa ne olur?**
C: Pandas otomatik olarak suffix ekler (_x, _y). Bunu kontrol etmek için merge parametrelerini kullanabilirsiniz.

**S: Excel'de formüller korunur mu?**
C: Hayır, sadece hesaplanmış değerler okunur. Formüller kaybolur.

**S: Farklı formatlardaki tarihler nasıl işlenir?**
C: Pandas otomatik tarih algılama yapar ancak parse_dates parametresi ile kontrol edebilirsiniz.

## Destek ve Katkı

Sorunlar veya öneriler için GitHub Issues kullanabilirsiniz.
