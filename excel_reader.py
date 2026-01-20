"""
Excel okuyucu modülü - Farklı sayfalardan veri çekme
"""
import pandas as pd
from typing import Dict, List, Optional
import openpyxl


class ExcelReader:
    """Excel dosyasından farklı sayfaları okuma sınıfı"""
    
    def __init__(self, dosya_yolu: str):
        """
        Excel dosyasını yükle
        
        Args:
            dosya_yolu: Excel dosyasının yolu
        """
        self.dosya_yolu = dosya_yolu
        self.excel_dosyasi = None
        self.sayfa_isimleri = []
        self._dosyayi_yukle()
    
    def _dosyayi_yukle(self):
        """Excel dosyasını yükle ve sayfa isimlerini al"""
        try:
            self.excel_dosyasi = pd.ExcelFile(self.dosya_yolu)
            self.sayfa_isimleri = self.excel_dosyasi.sheet_names
            print(f"Excel dosyası yüklendi: {self.dosya_yolu}")
            print(f"Bulunan sayfalar: {', '.join(self.sayfa_isimleri)}")
        except Exception as e:
            print(f"Hata: Excel dosyası yüklenemedi - {e}")
            raise
    
    def sayfa_oku(self, sayfa_adi: str, **kwargs) -> pd.DataFrame:
        """
        Belirtilen sayfayı oku
        
        Args:
            sayfa_adi: Okunacak sayfa adı
            **kwargs: pandas read_excel için ek parametreler
            
        Returns:
            Sayfa verilerini içeren DataFrame
        """
        if sayfa_adi not in self.sayfa_isimleri:
            raise ValueError(f"'{sayfa_adi}' sayfası bulunamadı. Mevcut sayfalar: {self.sayfa_isimleri}")
        
        try:
            df = pd.read_excel(self.excel_dosyasi, sheet_name=sayfa_adi, **kwargs)
            print(f"'{sayfa_adi}' sayfası okundu. Satır sayısı: {len(df)}")
            return df
        except Exception as e:
            print(f"Hata: '{sayfa_adi}' sayfası okunamadı - {e}")
            raise
    
    def tum_sayfalari_oku(self, **kwargs) -> Dict[str, pd.DataFrame]:
        """
        Tüm sayfaları oku
        
        Args:
            **kwargs: pandas read_excel için ek parametreler
            
        Returns:
            Sayfa adı -> DataFrame eşlemesi
        """
        tum_veriler = {}
        for sayfa_adi in self.sayfa_isimleri:
            tum_veriler[sayfa_adi] = self.sayfa_oku(sayfa_adi, **kwargs)
        return tum_veriler
    
    def belirli_sayfalari_oku(self, sayfa_isimleri: List[str], **kwargs) -> Dict[str, pd.DataFrame]:
        """
        Belirtilen sayfaları oku
        
        Args:
            sayfa_isimleri: Okunacak sayfa isimleri listesi
            **kwargs: pandas read_excel için ek parametreler
            
        Returns:
            Sayfa adı -> DataFrame eşlemesi
        """
        veriler = {}
        for sayfa_adi in sayfa_isimleri:
            veriler[sayfa_adi] = self.sayfa_oku(sayfa_adi, **kwargs)
        return veriler
    
    def sayfa_bilgisi_al(self, sayfa_adi: str) -> dict:
        """
        Sayfa hakkında özet bilgi al
        
        Args:
            sayfa_adi: Bilgisi alınacak sayfa adı
            
        Returns:
            Sayfa özet bilgileri
        """
        df = self.sayfa_oku(sayfa_adi)
        return {
            'sayfa_adi': sayfa_adi,
            'satir_sayisi': len(df),
            'sutun_sayisi': len(df.columns),
            'sutunlar': list(df.columns),
            'veri_tipleri': df.dtypes.to_dict()
        }
