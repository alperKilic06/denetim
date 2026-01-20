#!/usr/bin/env python3
"""
Örnek Excel dosyası oluşturucu
"""
import pandas as pd
import os


def ornek_excel_olustur():
    """Test için örnek Excel dosyası oluştur"""
    
    # Çıktı dizinini oluştur
    if not os.path.exists('output'):
        os.makedirs('output')
    
    # Sayfa 1: Çalışan Bilgileri
    calisanlar = pd.DataFrame({
        'ID': [1, 2, 3, 4, 5],
        'Ad': ['Ahmet Yılmaz', 'Ayşe Demir', 'Mehmet Kaya', 'Fatma Şahin', 'Ali Çelik'],
        'Departman': ['Satış', 'IT', 'Satış', 'İK', 'IT'],
        'Maaş': [5000, 7000, 5500, 6000, 7500],
        'Başlangıç Tarihi': ['2020-01-15', '2019-03-20', '2021-06-10', '2020-11-05', '2018-08-22']
    })
    
    # Sayfa 2: Satış Verileri
    satislar = pd.DataFrame({
        'Çalışan ID': [1, 1, 2, 3, 3, 3, 4, 5],
        'Ürün': ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Laptop', 'Mouse', 'Printer', 'Server'],
        'Miktar': [2, 5, 3, 1, 1, 10, 2, 1],
        'Birim Fiyat': [15000, 100, 500, 3000, 15000, 100, 2000, 50000],
        'Satış Tarihi': ['2024-01-10', '2024-01-11', '2024-01-12', '2024-01-13', 
                        '2024-01-14', '2024-01-15', '2024-01-16', '2024-01-17']
    })
    
    # Toplam tutarı hesapla
    satislar['Toplam Tutar'] = satislar['Miktar'] * satislar['Birim Fiyat']
    
    # Sayfa 3: Departman Özeti
    departmanlar = pd.DataFrame({
        'Departman': ['Satış', 'IT', 'İK'],
        'Çalışan Sayısı': [2, 2, 1],
        'Ortalama Maaş': [5250, 7250, 6000],
        'Bütçe': [100000, 150000, 80000]
    })
    
    # Excel dosyasını oluştur
    dosya_adi = 'ornek_veri.xlsx'
    
    with pd.ExcelWriter(dosya_adi, engine='openpyxl') as writer:
        calisanlar.to_excel(writer, sheet_name='Çalışanlar', index=False)
        satislar.to_excel(writer, sheet_name='Satışlar', index=False)
        departmanlar.to_excel(writer, sheet_name='Departmanlar', index=False)
    
    print(f"Örnek Excel dosyası oluşturuldu: {dosya_adi}")
    print(f"Sayfalar: Çalışanlar, Satışlar, Departmanlar")
    
    return dosya_adi


if __name__ == "__main__":
    ornek_excel_olustur()
