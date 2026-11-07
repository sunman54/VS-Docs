# Universal Document Viewer (VS Code Extension)

VS Code üzerinde çalışan; **Word (`.docx`, `.doc`)**, **Excel (`.xlsx`, `.xls`, `.xlsm`, `.csv`, `.tsv`, `.ods`)**, **PDF (`.pdf`)**, **OpenDocument (`.odt`)** ve **Zengin Metin (`.rtf`)** dosyalarını harici bir ofis yazılımına ihtiyaç duymadan **salt okunur (read-only)** olarak açan, yüksek performanslı ve **Koyu Tema (Dark Mode)** uyumlu kapsamlı doküman görüntüleyici eklentisi.

---

## 🚀 Öne Çıkan Özellikler

### 📄 1. Word & Metin Dokümanları (`.docx`, `.doc`, `.odt`, `.rtf`)
- **OpenXML Doğruluğu:** `docx-preview` motoru ile paragraflar, tablolar, biçimlendirmeler, başlıklar ve gömülü görseller aslına uygun çizilir.
- **📑 İçindekiler / Başlık Ağacı (Outline):** Dokümandaki `H1, H2, H3` başlıklarını otomatik tarayarak sol tarafta gezinti paneli oluşturur. Tıklanan başlığa yumuşak animasyonla kaydırır.
- **🖼️ Gömülü Görselleri Dışa Aktarma:** Word dokümanı içerisindeki tüm çizelge, diyagram ve resimleri tek tıkla `[dokuman]_gorseller/` klasörüne kaydeder.
- **Eski Format & Açık Standartlar:** Word 97-2003 (`.doc`), OpenDocument Text (`.odt`) ve Rich Text (`.rtf`) dosyalarını anında açar.

### 📊 2. Excel & Veri Tabloları (`.xlsx`, `.xls`, `.xlsm`, `.csv`, `.tsv`, `.ods`)
- **Çoklu Çalışma Sayfaları:** Altta sekme çubuğu ve üstte hızlı açılır sayfa menüsü.
- **📐 Sütun Genişliği Boyutlandırma:** Sütun sınırlarını fareyle sürükleyerek genişliği dilediğiniz gibi ayarlama.
- **🔍 Sütun Bazlı Filtreleme:** Her sütun başlığındaki filtreleme menüsüyle istenilen değerleri seçip anında filtreleme.
- **🔢 Çoklu Hücre Seçimi ve Anlık İstatistik:** Fareyle aralık seçildiğinde durum çubuğunda canlı `Seçim | Toplam | Ortalama | Min | Max` hesaplaması.
- **Akıllı CSV Ayrıştırıcı:** Virgül (`,`), noktalı virgül (`;`), tab (`\t`) veya dikey çizgi (`|`) ayraçlarını otomatik algılama ve tek tıkla değiştirebilme.

### 📋 3. Geliştirici Süper Güçleri
- **Markdown Olarak Kopyala:** Tabloyu veya seçili hücreleri tek tıkla GitHub Flavored Markdown formatında panoya kopyalar.
- **JSON Olarak Dışa Aktar:** Tablo satırlarını doğrudan JSON nesne dizisi olarak panoya kopyalar.
- **Yazdır / PDF Desteği:** `@media print` stilleri ile üst çubuk ve araçlar gizlenerek temiz kenar boşluklarıyla yazdırma / PDF kaydetme.

### 📕 4. PDF Görüntüleyici (`.pdf`)
- Mozilla **`pdfjs-dist`** motoru ile sıfır harici bağımlılıkla yüksek kaliteli PDF sayfalama.
- **Koyu Mod PDF Filtresi:** Gece çalışma konforu için PDF sayfalarını koyu temayla uyumlu hale getiren akıllı renk dönüştürücü.

### 🌓 5. Akıllı Koyu Tema (Adaptive Dark Mode)
- VS Code'un aktif renk temasıyla (`var(--vscode-*)`) otomatik senkronize olur.
- Word sayfalarında koyu zemin üzerinde koyu renkli metinlerin kaybolmasını önleyen akıllı renk uyarlama motoru barındırır.
- Üst araç çubuğundaki tema butonu ile istenildiğinde **Otomatik**, **Koyu** veya **Açık** moda zorlanabilir.

---

## 🛠️ Mimari ve Proje Yapısı

```
vcDocs/
├── dist/                        # Derlenmiş üretim paketleri
│   ├── extension.js             # Extension Host mantığı
│   ├── webview.js               # Webview istemci kodu (PDF.js, SheetJS, docx-preview, DOMPurify)
│   └── webview.css              # Koyu/Açık tema ve responsive grid stilleri
├── src/
│   ├── extension.ts             # Eklenti aktivasyon ve komutları
│   ├── DocViewerCustomEditorProvider.ts # vscode.CustomReadonlyEditorProvider & IPC
│   └── webview/
│       ├── main.ts              # Webview router & denetleyici
│       ├── themeManager.ts      # Dinamik tema takipçisi
│       ├── outlineManager.ts    # Word içindekiler gezgini
│       ├── imageExporter.ts     # Word resim ayıklayıcı
│       └── renderers/
│           ├── docxRenderer.ts  # Word OpenXML motoru
│           ├── excelRenderer.ts # SheetJS, grid, boyutlandırma ve filtreleme
│           ├── pdfRenderer.ts   # Mozilla PDF.js canvas motoru
│           ├── odtRenderer.ts   # OpenDocument & RTF motoru
│           └── legacyDocRenderer.ts # .doc ikili metin çıkarıcı
├── test/
│   ├── testParsers.js           # 6 kategorili otomatik test paketi
│   └── samples/                 # Örnek test dosyaları (.docx, .xlsx, .csv, .pdf, .rtf)
└── vcdocs-viewer-1.0.0.vsix     # Dağıtıma hazır VSIX paketi
```

---

## 🧪 Test ve Kurulum

### Otomatik Testleri Çalıştırma
```bash
npm test
```

### Derleme ve Paketleme
```bash
# TypeScript kontrolü ve esbuild derlemesi
npm run build

# Dağıtım için VSIX paketi üretme
npm run package
```

### VS Code'a VSIX Yükleme
Üretilen `vcdocs-viewer-1.0.0.vsix` dosyasını VS Code'da çalıştırmak için:
1. VS Code Extensions sekmesini açın (`Ctrl+Shift+X`).
2. Sağ üstteki `...` (Daha Fazla Eylem) menüsüne tıklayın.
3. **"Install from VSIX..."** seçeneğini seçip `vcdocs-viewer-1.0.0.vsix` dosyasını gösterin.
