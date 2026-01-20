#!/usr/bin/env python3
"""
Excel Rapor Oluşturucu - Ana Program

Bu program Excel dosyalarından farklı sayfaları okuyup rapor oluşturur.
"""
import sys
import os
from excel_reader import ExcelReader
from rapor_olusturucu import RaporOlusturucu


def ornek_kullanim_1(excel_dosyasi: str):
    """
    Örnek 1: Tüm sayfaları okuyup özet rapor oluştur
    """
    print("\n" + "=" * 80)
    print("ÖRNEK 1: TÜM SAYFALARI OKU VE ÖZET RAPOR OLUŞTUR")
    print("=" * 80)
    
    # Excel dosyasını oku
    okuyucu = ExcelReader(excel_dosyasi)
    
    # Tüm sayfaları oku
    tum_veriler = okuyucu.tum_sayfalari_oku()
    
    # Rapor oluştur
    rapor = RaporOlusturucu(tum_veriler)
    
    # Konsola özet rapor yazdır
    rapor.konsol_rapor_goster()
    
    # Excel olarak kaydet
    cikti_dosyasi = "output/ozet_rapor.xlsx"
    rapor.rapor_kaydet(cikti_dosyasi, rapor_turu='ozet')
    
    return rapor


def ornek_kullanim_2(excel_dosyasi: str):
    """
    Örnek 2: Belirli sayfaları okuyup birleştir
    """
    print("\n" + "=" * 80)
    print("ÖRNEK 2: BELİRLİ SAYFALARI OKU VE BİRLEŞTİR")
    print("=" * 80)
    
    # Excel dosyasını oku
    okuyucu = ExcelReader(excel_dosyasi)
    
    # Mevcut sayfa isimlerini göster
    print(f"\nMevcut sayfalar: {okuyucu.sayfa_isimleri}")
    
    # İlk iki sayfayı oku (eğer varsa)
    if len(okuyucu.sayfa_isimleri) >= 2:
        secili_sayfalar = okuyucu.sayfa_isimleri[:2]
        veriler = okuyucu.belirli_sayfalari_oku(secili_sayfalar)
        
        # Rapor oluştur
        rapor = RaporOlusturucu(veriler)
        
        # Sayfaları dikey birleştir
        birlesik_veri = rapor.sayfalari_birlestir(birlesim_turu='vertical')
        print(f"\nBirleşik veri boyutu: {birlesik_veri.shape}")
        print(f"İlk 5 satır:\n{birlesik_veri.head()}")
        
        # Birleşik rapor kaydet
        cikti_dosyasi = "output/birlesik_rapor.xlsx"
        rapor.rapor_kaydet(cikti_dosyasi, rapor_turu='birlesik', birlesim_turu='vertical')
        
        return rapor
    else:
        print("Uyarı: En az 2 sayfa gerekli")
        return None


def ornek_kullanim_3(excel_dosyasi: str):
    """
    Örnek 3: Her sayfa için istatistiksel analiz yap
    """
    print("\n" + "=" * 80)
    print("ÖRNEK 3: İSTATİSTİKSEL ANALİZ RAPORU")
    print("=" * 80)
    
    # Excel dosyasını oku
    okuyucu = ExcelReader(excel_dosyasi)
    
    # Tüm sayfaları oku
    tum_veriler = okuyucu.tum_sayfalari_oku()
    
    # Rapor oluştur
    rapor = RaporOlusturucu(tum_veriler)
    
    # Her sayfa için istatistik göster
    for sayfa_adi in tum_veriler.keys():
        print(f"\n--- {sayfa_adi} İstatistikleri ---")
        istatistikler = rapor.istatistik_rapor_olustur(sayfa_adi)
        if not istatistikler.empty:
            print(istatistikler)
        else:
            print("Sayısal veri bulunamadı")
    
    return rapor


def main():
    """Ana fonksiyon"""
    print("Excel Çok Sayfalı Rapor Oluşturucu")
    print("=" * 80)
    
    # Komut satırı argümanlarını kontrol et
    if len(sys.argv) < 2:
        print("\nKullanım: python main.py <excel_dosyasi> [ornek_numarasi]")
        print("\nÖrnekler:")
        print("  python main.py ornek_veri.xlsx        # Tüm örnekleri çalıştır")
        print("  python main.py ornek_veri.xlsx 1      # Sadece örnek 1'i çalıştır")
        print("  python main.py ornek_veri.xlsx 2      # Sadece örnek 2'yi çalıştır")
        print("  python main.py ornek_veri.xlsx 3      # Sadece örnek 3'ü çalıştır")
        sys.exit(1)
    
    excel_dosyasi = sys.argv[1]
    
    # Dosya varlığını kontrol et
    if not os.path.exists(excel_dosyasi):
        print(f"\nHata: '{excel_dosyasi}' dosyası bulunamadı!")
        sys.exit(1)
    
    # Hangi örneği çalıştıracağımızı belirle
    ornek_no = int(sys.argv[2]) if len(sys.argv) > 2 else 0
    
    try:
        if ornek_no == 0:
            # Tüm örnekleri çalıştır
            ornek_kullanim_1(excel_dosyasi)
            ornek_kullanim_2(excel_dosyasi)
            ornek_kullanim_3(excel_dosyasi)
        elif ornek_no == 1:
            ornek_kullanim_1(excel_dosyasi)
        elif ornek_no == 2:
            ornek_kullanim_2(excel_dosyasi)
        elif ornek_no == 3:
            ornek_kullanim_3(excel_dosyasi)
        else:
            print(f"Hata: Geçersiz örnek numarası: {ornek_no}")
            sys.exit(1)
        
        print("\n" + "=" * 80)
        print("İşlem tamamlandı!")
        print("=" * 80)
        
    except Exception as e:
        print(f"\nHata oluştu: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
