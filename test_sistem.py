#!/usr/bin/env python3
"""
Test scripti - Kurulumun ve temel işlevlerin doğrulaması
"""
import sys
import os


def test_imports():
    """Gerekli modüllerin yüklenip yüklenmediğini test et"""
    print("1. Modül yüklemeleri test ediliyor...")
    try:
        import pandas
        import openpyxl
        from excel_reader import ExcelReader
        from rapor_olusturucu import RaporOlusturucu
        print("   ✓ Tüm modüller başarıyla yüklendi")
        return True
    except ImportError as e:
        print(f"   ✗ Modül yükleme hatası: {e}")
        print("   Lütfen 'pip install -r requirements.txt' komutunu çalıştırın")
        return False


def test_excel_creation():
    """Örnek Excel dosyası oluşturmayı test et"""
    print("\n2. Örnek Excel dosyası oluşturma testi...")
    try:
        import pandas as pd
        
        # Basit bir test Excel dosyası oluştur
        test_file = 'test_temp.xlsx'
        df1 = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
        df2 = pd.DataFrame({'C': [7, 8, 9], 'D': [10, 11, 12]})
        
        with pd.ExcelWriter(test_file, engine='openpyxl') as writer:
            df1.to_excel(writer, sheet_name='Sayfa1', index=False)
            df2.to_excel(writer, sheet_name='Sayfa2', index=False)
        
        # Dosyanın oluştuğunu kontrol et
        if os.path.exists(test_file):
            print("   ✓ Excel dosyası başarıyla oluşturuldu")
            os.remove(test_file)
            return True
        else:
            print("   ✗ Excel dosyası oluşturulamadı")
            return False
            
    except Exception as e:
        print(f"   ✗ Hata: {e}")
        return False


def test_excel_reading():
    """Excel okuma işlevini test et"""
    print("\n3. Excel okuma testi...")
    try:
        from excel_reader import ExcelReader
        import pandas as pd
        
        # Test dosyası oluştur
        test_file = 'test_temp.xlsx'
        df1 = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
        df2 = pd.DataFrame({'C': [7, 8, 9], 'D': [10, 11, 12]})
        
        with pd.ExcelWriter(test_file, engine='openpyxl') as writer:
            df1.to_excel(writer, sheet_name='Sayfa1', index=False)
            df2.to_excel(writer, sheet_name='Sayfa2', index=False)
        
        # Oku
        okuyucu = ExcelReader(test_file)
        veriler = okuyucu.tum_sayfalari_oku()
        
        # Kontrol et
        if len(veriler) == 2 and 'Sayfa1' in veriler and 'Sayfa2' in veriler:
            print("   ✓ Excel dosyası başarıyla okundu")
            os.remove(test_file)
            return True
        else:
            print("   ✗ Excel okuma hatası")
            if os.path.exists(test_file):
                os.remove(test_file)
            return False
            
    except Exception as e:
        print(f"   ✗ Hata: {e}")
        if os.path.exists('test_temp.xlsx'):
            os.remove('test_temp.xlsx')
        return False


def test_report_generation():
    """Rapor oluşturma işlevini test et"""
    print("\n4. Rapor oluşturma testi...")
    try:
        from excel_reader import ExcelReader
        from rapor_olusturucu import RaporOlusturucu
        import pandas as pd
        
        # Test dosyası oluştur
        test_file = 'test_temp.xlsx'
        output_file = 'test_output.xlsx'
        
        df1 = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
        df2 = pd.DataFrame({'C': [7, 8, 9], 'D': [10, 11, 12]})
        
        with pd.ExcelWriter(test_file, engine='openpyxl') as writer:
            df1.to_excel(writer, sheet_name='Sayfa1', index=False)
            df2.to_excel(writer, sheet_name='Sayfa2', index=False)
        
        # Oku ve rapor oluştur
        okuyucu = ExcelReader(test_file)
        veriler = okuyucu.tum_sayfalari_oku()
        rapor = RaporOlusturucu(veriler)
        rapor.rapor_kaydet(output_file, rapor_turu='ozet')
        
        # Kontrol et
        if os.path.exists(output_file):
            print("   ✓ Rapor başarıyla oluşturuldu")
            os.remove(test_file)
            os.remove(output_file)
            return True
        else:
            print("   ✗ Rapor oluşturulamadı")
            if os.path.exists(test_file):
                os.remove(test_file)
            return False
            
    except Exception as e:
        print(f"   ✗ Hata: {e}")
        for f in ['test_temp.xlsx', 'test_output.xlsx']:
            if os.path.exists(f):
                os.remove(f)
        return False


def main():
    """Ana test fonksiyonu"""
    print("=" * 70)
    print("Excel Rapor Oluşturucu - Test Scripti")
    print("=" * 70)
    
    results = []
    
    # Testleri çalıştır
    results.append(("Modül Yükleme", test_imports()))
    results.append(("Excel Oluşturma", test_excel_creation()))
    results.append(("Excel Okuma", test_excel_reading()))
    results.append(("Rapor Oluşturma", test_report_generation()))
    
    # Sonuçları göster
    print("\n" + "=" * 70)
    print("TEST SONUÇLARI")
    print("=" * 70)
    
    all_passed = True
    for test_name, result in results:
        status = "✓ BAŞARILI" if result else "✗ BAŞARISIZ"
        print(f"{test_name:<20} : {status}")
        if not result:
            all_passed = False
    
    print("=" * 70)
    
    if all_passed:
        print("\n🎉 Tüm testler başarılı! Sistem kullanıma hazır.")
        return 0
    else:
        print("\n⚠️  Bazı testler başarısız oldu. Lütfen hataları kontrol edin.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
