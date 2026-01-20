"""
Rapor oluşturucu modülü - Excel sayfalarından rapor üretme
"""
import pandas as pd
from typing import Dict, List, Optional
from datetime import datetime
import os


class RaporOlusturucu:
    """Farklı Excel sayfalarından rapor oluşturma sınıfı"""
    
    def __init__(self, veriler: Dict[str, pd.DataFrame]):
        """
        Rapor oluşturucuyu başlat
        
        Args:
            veriler: Sayfa adı -> DataFrame eşlemesi
        """
        self.veriler = veriler
        self.rapor_tarihi = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    def ozet_rapor_olustur(self) -> pd.DataFrame:
        """
        Tüm sayfalar için özet rapor oluştur
        
        Returns:
            Özet bilgiler içeren DataFrame
        """
        ozet_veriler = []
        
        for sayfa_adi, df in self.veriler.items():
            ozet_veriler.append({
                'Sayfa Adı': sayfa_adi,
                'Satır Sayısı': len(df),
                'Sütun Sayısı': len(df.columns),
                'Sütunlar': ', '.join(df.columns)
            })
        
        ozet_df = pd.DataFrame(ozet_veriler)
        return ozet_df
    
    def sayfalari_birlestir(self, 
                           sayfa_isimleri: Optional[List[str]] = None,
                           birlesim_turu: str = 'vertical',
                           ortak_sutun: Optional[str] = None) -> pd.DataFrame:
        """
        Belirtilen sayfaları birleştir
        
        Args:
            sayfa_isimleri: Birleştirilecek sayfa isimleri (None ise hepsi)
            birlesim_turu: 'vertical' (alt alta) veya 'horizontal' (yan yana)
            ortak_sutun: Horizontal birleşim için ortak sütun adı
            
        Returns:
            Birleştirilmiş DataFrame
        """
        if sayfa_isimleri is None:
            sayfa_isimleri = list(self.veriler.keys())
        
        dfs = [self.veriler[sayfa] for sayfa in sayfa_isimleri if sayfa in self.veriler]
        
        if not dfs:
            raise ValueError("Birleştirilecek veri bulunamadı")
        
        if birlesim_turu == 'vertical':
            # Alt alta birleştir
            birlesik_df = pd.concat(dfs, ignore_index=True)
            print(f"{len(dfs)} sayfa dikey olarak birleştirildi")
        elif birlesim_turu == 'horizontal':
            # Yan yana birleştir
            if ortak_sutun:
                birlesik_df = dfs[0]
                for df in dfs[1:]:
                    birlesik_df = pd.merge(birlesik_df, df, on=ortak_sutun, how='outer')
                print(f"{len(dfs)} sayfa '{ortak_sutun}' sütunu üzerinden birleştirildi")
            else:
                birlesik_df = pd.concat(dfs, axis=1)
                print(f"{len(dfs)} sayfa yatay olarak birleştirildi")
        else:
            raise ValueError("birlesim_turu 'vertical' veya 'horizontal' olmalıdır")
        
        return birlesik_df
    
    def rapor_kaydet(self, 
                     cikti_dosyasi: str,
                     rapor_turu: str = 'ozet',
                     sayfa_isimleri: Optional[List[str]] = None,
                     **kwargs):
        """
        Raporu Excel dosyası olarak kaydet
        
        Args:
            cikti_dosyasi: Çıktı dosya yolu
            rapor_turu: 'ozet', 'birlesik', veya 'ayri_sayfalar'
            sayfa_isimleri: İşlenecek sayfa isimleri (None ise hepsi)
            **kwargs: Ek parametreler (birlesim_turu, ortak_sutun vb.)
        """
        # Çıktı dizinini oluştur
        cikti_dizini = os.path.dirname(cikti_dosyasi)
        if cikti_dizini and not os.path.exists(cikti_dizini):
            os.makedirs(cikti_dizini)
        
        with pd.ExcelWriter(cikti_dosyasi, engine='openpyxl') as writer:
            if rapor_turu == 'ozet':
                # Özet rapor oluştur
                ozet_df = self.ozet_rapor_olustur()
                ozet_df.to_excel(writer, sheet_name='Özet', index=False)
                
                # Her sayfa için detayları ekle
                for sayfa_adi, df in self.veriler.items():
                    # Excel sayfa adı uzunluk sınırlaması (31 karakter)
                    kisa_ad = sayfa_adi[:31]
                    df.to_excel(writer, sheet_name=kisa_ad, index=False)
                
                print(f"Özet rapor kaydedildi: {cikti_dosyasi}")
                
            elif rapor_turu == 'birlesik':
                # Birleşik rapor oluştur
                birlesik_df = self.sayfalari_birlestir(sayfa_isimleri, **kwargs)
                birlesik_df.to_excel(writer, sheet_name='Birleşik Veri', index=False)
                print(f"Birleşik rapor kaydedildi: {cikti_dosyasi}")
                
            elif rapor_turu == 'ayri_sayfalar':
                # Sadece belirtilen sayfaları kaydet
                sayfalar = sayfa_isimleri if sayfa_isimleri else list(self.veriler.keys())
                for sayfa_adi in sayfalar:
                    if sayfa_adi in self.veriler:
                        kisa_ad = sayfa_adi[:31]
                        self.veriler[sayfa_adi].to_excel(writer, sheet_name=kisa_ad, index=False)
                print(f"Seçili sayfalar kaydedildi: {cikti_dosyasi}")
            else:
                raise ValueError("rapor_turu 'ozet', 'birlesik' veya 'ayri_sayfalar' olmalıdır")
    
    def istatistik_rapor_olustur(self, sayfa_adi: str, sayisal_sutunlar: Optional[List[str]] = None) -> pd.DataFrame:
        """
        Belirtilen sayfa için istatistiksel rapor oluştur
        
        Args:
            sayfa_adi: Analiz edilecek sayfa adı
            sayisal_sutunlar: Analiz edilecek sütunlar (None ise tüm sayısal sütunlar)
            
        Returns:
            İstatistiksel özet DataFrame
        """
        if sayfa_adi not in self.veriler:
            raise ValueError(f"'{sayfa_adi}' sayfası bulunamadı")
        
        df = self.veriler[sayfa_adi]
        
        if sayisal_sutunlar:
            df_analiz = df[sayisal_sutunlar]
        else:
            df_analiz = df.select_dtypes(include=['number'])
        
        if df_analiz.empty:
            print(f"Uyarı: '{sayfa_adi}' sayfasında sayısal veri bulunamadı")
            return pd.DataFrame()
        
        istatistikler = df_analiz.describe()
        return istatistikler
    
    def konsol_rapor_goster(self):
        """Konsola özet rapor yazdır"""
        print("\n" + "=" * 80)
        print(f"RAPOR TARİHİ: {self.rapor_tarihi}")
        print("=" * 80)
        
        ozet_df = self.ozet_rapor_olustur()
        print("\nSAYFA ÖZETLERİ:")
        print(ozet_df.to_string(index=False))
        
        print("\n" + "=" * 80)
