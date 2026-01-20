# Denetim - Excel Rapor Uygulaması

Arvento ile denetim raporlama yapmak için web uygulaması.

## Özellikler

- 📊 Excel dosyalarını (.xlsx, .xls) yükleme
- 📑 Excel sayfalarından veri çekme
- 📈 Otomatik rapor oluşturma
- 🖨️ Raporu yazdırma ve indirme
- 🎨 Modern ve kullanıcı dostu arayüz

## Gereksinimler

- **Node.js:** v18.0.0 veya üzeri
- **npm:** v9.0.0 veya üzeri

Node.js versiyonunuzu kontrol etmek için:
```bash
node --version
npm --version
```

Eğer Node.js yüklü değilse veya eski versiyonunuz varsa, [nodejs.org](https://nodejs.org/) adresinden en son LTS versiyonunu indirin.

## Kurulum

### 1. Depoyu klonlayın
```bash
git clone https://github.com/alperKilic06/denetim.git
cd denetim
```

### 2. Bağımlılıkları yükleyin
```bash
npm install
```

**Hata alıyorsanız:**

- Node.js versiyonunuzun 18 veya üzeri olduğundan emin olun
- npm önbelleğini temizleyin:
  ```bash
  npm cache clean --force
  npm install
  ```
- `node_modules` ve `package-lock.json` dosyalarını silip tekrar deneyin:
  ```bash
  rm -rf node_modules package-lock.json
  npm install
  ```

### 3. Uygulamayı başlatın
```bash
npm start
```

### 4. Tarayıcınızda açın
```
http://localhost:3000
```

## Kullanım

1. Ana sayfada "Dosya Seç" butonuna tıklayın veya Excel dosyanızı sürükleyip bırakın
2. Dosyayı seçtikten sonra "Yükle ve İşle" butonuna tıklayın
3. Excel verileriniz işlenecek ve sayfa bilgileri gösterilecektir
4. "Rapor Oluştur" butonuna tıklayarak raporu oluşturun
5. Oluşturulan raporu yazdırabilir veya indirebilirsiniz

## Teknolojiler

- **Backend:** Node.js v20, Express
- **Excel İşleme:** ExcelJS
- **Dosya Yükleme:** Multer v2.0.2
- **Güvenlik:** express-rate-limit
- **Frontend:** HTML, CSS, JavaScript

## Sorun Giderme

### npm install çalışmıyor
- Node.js versiyonunuzu kontrol edin (`node --version`)
- En az Node.js v18 gereklidir
- npm önbelleğini temizleyin: `npm cache clean --force`

### Port 3000 kullanımda hatası
- Başka bir uygulama 3000 portunu kullanıyor olabilir
- `server.js` dosyasındaki `port` değişkenini değiştirin

### Excel dosyası yüklenmiyor
- Dosyanın .xlsx veya .xls uzantılı olduğundan emin olun
- Dosya boyutunun makul olduğundan emin olun
- Tarayıcı konsolunda hata mesajlarını kontrol edin

## Lisans

ISC
