# Universal Document Viewer (VS Code Extension)

VS Code üzerinde çalışan; Word (`.docx`, `.doc`), Excel (`.xlsx`, `.xls`, `.xlsm`) ve veri tablolarını (`.csv`, `.tsv`) harici bir ofis yazılımına ihtiyaç duymadan **salt okunur (read-only)** olarak açan, yüksek performanslı ve **Koyu Tema (Dark Mode)** uyumlu doküman görüntüleyici eklentisi.

---

## 🚀 Özellikler

- 📄 **Word Görüntüleyici (`.docx`):** `docx-preview` motoru ile Word stilleri, paragraflar, başlıklar, resimler ve tablolar aslına uygun biçimde çizilir.
- 📊 **Excel Görüntüleyici (`.xlsx`, `.xls`, `.xlsm`, `.csv`, `.tsv`):**
  - Çoklu çalışma sayfası (sheet) desteği.
  - Altta sekme çubuğu ve üst araç çubuğunda hızlı sayfa seçici.
  - Sabitlenmiş satır ve sütun başlıkları (`A, B, C...` ve `1, 2, 3...`).
  - Sayı, tarih ve metin hücre hizalama ve biçimlendirmeleri.
  - Büyük veri setleri için akıllı parçalı yükleme (Lazy Load / Pagination).
- 🌓 **Akıllı Koyu Tema (Adaptive Dark Mode):**
  - VS Code'un aktif teması ile otomatik senkronizasyon (`var(--vscode-*)`).
  - Word dokümanları için gömülü renkleri ve sayfaları koyu temaya adapte eden filtreleme.
  - İstenildiğinde tek tıkla Açık / Koyu / Sistem temaları arasında geçiş yapabilen araç çubuğu düğmesi.
- 🔍 **Doküman İçi Arama:**
  - Excel tablolarında eşleşen hücreleri canlı vurgulama ve eşleşme sayısı sayacı.
  - Word dokümanlarında yerel arama entegrasyonu.
- 🔎 **Zoom Kontrolleri:** Word dokümanlarında `%40` ile `%250` arasında ölçekleme (Yakınlaştır, Uzaklaştır, 1:1 Sıfırla).
- 🔒 **Güvenlik (CSP & Sanitization):** `DOMPurify` ve sıkı Content Security Policy (CSP) ile korunur. Makro veya harici betik çalıştırmaz.
- 📦 **Legacy Desteği (`.doc`):** Eski ikili Word 97-2003 formatından metin içeriğini ayrıştırarak kullanıcıya doğrudan sunar.

---

## 🛠️ Mimari ve Proje Yapısı

```
vcDocs/
├── .vscode/
│   ├── launch.json              # F5 Hata Ayıklama (Extension Host)
│   └── tasks.json               # Otomatik derleme görevi
├── src/
│   ├── extension.ts             # Eklenti aktivasyon ve komut tanımları
│   ├── DocViewerCustomEditorProvider.ts # vscode.CustomReadonlyEditorProvider
│   └── webview/
│       ├── main.ts              # Webview kontrolcü ve yönlendirici (Router)
│       ├── styles.css           # Koyu tema ve arayüz stilleri
│       ├── themeManager.ts      # Tema senkronizasyon motoru
│       └── renderers/
│           ├── docxRenderer.ts  # Word OpenXML render motoru
│           ├── excelRenderer.ts # SheetJS tablo ve sekme motoru
│           └── legacyDocRenderer.ts # .doc ikili metin çıkarıcı
├── test/
│   ├── testParsers.js           # Otomatik test paketi
│   ├── createDocx.js            # Test DOCX üretici
│   └── samples/                 # Test ofis dosyaları (.docx, .xlsx, .csv)
├── esbuild.js                   # Çift hedefli esbuild derleyici
├── package.json
└── tsconfig.json
```

---

## 🧪 Canlı Test ve Geliştirme

### 1. Test Paketini Çalıştırma
Tüm ayrıştırıcıları, şema ve bundle bütünlüğünü doğrulamak için:
```bash
npm test
```

### 2. VS Code İçinde Canlı Çalıştırma
1. Bu projeyi VS Code'da açın.
2. Klavyenizden **`F5`** tuşuna basın.
3. Açılan yeni **Extension Development Host** penceresinde `test/samples/` altındaki:
   - `ornek_dokuman.docx`
   - `test_butce.xlsx`
   - `test_veriler.csv`
   dosyalarına tıklayın. Eklenti otomatik olarak özel editör sekmesinde açılacaktır.
4. VS Code temasını `Ctrl+K Ctrl+T` ile Koyu ve Açık temalar arasında değiştirerek adaptasyonu anlık izleyebilirsiniz.

