#!/usr/bin/env python3
"""
Tam İş Akışı Demosu - Excel çok sayfalı rapor oluşturma
"""
from excel_reader import ExcelReader
from rapor_olusturucu import RaporOlusturucu
import pandas as pd
import os


def demo_workflow():
    """Tam bir rapor oluşturma iş akışını göster"""
    print("=" * 80)
    print("EXCEL ÇOK SAYFALI RAPOR OLUŞTURMA - TAM İŞ AKIŞI")
    print("=" * 80)
    
    # 1. Örnek veri dosyası oluştur
    print("\n1. Örnek Excel dosyası oluşturuluyor...")
    if not os.path.exists('ornek_veri.xlsx'):
        import ornek_veri_olustur
        ornek_veri_olustur.ornek_excel_olustur()
    else:
        print("   Mevcut örnek dosya kullanılıyor: ornek_veri.xlsx")
    
    # 2. Excel dosyasını oku
    print("\n2. Excel dosyası okunuyor...")
    okuyucu = ExcelReader('ornek_veri.xlsx')
    print(f"   Toplam {len(okuyucu.sayfa_isimleri)} sayfa bulundu")
    
    # 3. Tüm sayfaları oku
    print("\n3. Tüm sayfalar okunuyor...")
    tum_veriler = okuyucu.tum_sayfalari_oku()
    
    # 4. Rapor oluştur
    print("\n4. Rapor oluşturucu başlatılıyor...")
    rapor = RaporOlusturucu(tum_veriler)
    
    # 5. Özet raporu konsola yazdır
    print("\n5. Özet rapor:")
    rapor.konsol_rapor_goster()
    
    # 6. Farklı rapor türleri oluştur
    print("\n6. Farklı rapor türleri oluşturuluyor...")
    
    # Çıktı dizini oluştur
    if not os.path.exists('output'):
        os.makedirs('output')
    
    # Özet rapor
    rapor.rapor_kaydet('output/tam_ozet.xlsx', rapor_turu='ozet')
    print("   ✓ Özet rapor: output/tam_ozet.xlsx")
    
    # Birleşik rapor (dikey)
    rapor.rapor_kaydet('output/birlesik_dikey.xlsx', rapor_turu='birlesik', birlesim_turu='vertical')
    print("   ✓ Birleşik rapor (dikey): output/birlesik_dikey.xlsx")
    
    # Çalışan ve Departman birleştirme (yatay)
    calisanlar_dept = okuyucu.belirli_sayfalari_oku(['Çalışanlar', 'Departmanlar'])
    rapor_yatay = RaporOlusturucu(calisanlar_dept)
    rapor_yatay.rapor_kaydet(
        'output/birlesik_yatay.xlsx',
        rapor_turu='birlesik',
        birlesim_turu='horizontal',
        ortak_sutun='Departman'
    )
    print("   ✓ Birleşik rapor (yatay): output/birlesik_yatay.xlsx")
    
    # 7. İstatistikler
    print("\n7. İstatistiksel analizler:")
    for sayfa_adi in tum_veriler.keys():
        istatistikler = rapor.istatistik_rapor_olustur(sayfa_adi)
        if not istatistikler.empty:
            print(f"\n   {sayfa_adi}:")
            print(f"   - Sayısal sütun sayısı: {len(istatistikler.columns)}")
            print(f"   - Analiz edilen satır: {int(istatistikler.loc['count'].iloc[0])}")
    
    # 8. Başarı özeti
    print("\n" + "=" * 80)
    print("İŞ AKIŞI TAMAMLANDI!")
    print("=" * 80)
    print("\nOluşturulan dosyalar:")
    for file in ['output/tam_ozet.xlsx', 'output/birlesik_dikey.xlsx', 'output/birlesik_yatay.xlsx']:
        if os.path.exists(file):
            size = os.path.getsize(file)
            print(f"  ✓ {file} ({size:,} bytes)")
    
    print("\nKullanım örnekleri için:")
    print("  - README.md dosyasına bakın")
    print("  - KULLANIM_KILAVUZU.md dosyasına bakın")
    print("  - python main.py ornek_veri.xlsx komutunu çalıştırın")
    print("\n" + "=" * 80)


if __name__ == "__main__":
    demo_workflow()
