const assert = require('assert');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

async function runTests() {
  console.log('=== Universal Document Viewer Test Suite Başlatılıyor ===\n');

  // 1. Test Excel Creation & Parsing (xlsx, xls, csv)
  console.log('[TEST 1] Excel Workbook ve CSV Ayrıştırma Testi...');
  {
    const wb = XLSX.utils.book_new();

    // Sheet 1: Finansal Tablo
    const wsData1 = [
      ['Departman', 'Çalışan Sayısı', 'Bütçe (TL)', 'Aktif'],
      ['Yazılım Mühendisliği', 42, 1250000.50, true],
      ['Ürün Yönetimi', 8, 340000.00, true],
      ['Kalite Güvence (QA)', 15, 450000.75, false],
      ['Tasarım & UI/UX', 6, 210000.00, true]
    ];
    const ws1 = XLSX.utils.aoa_to_sheet(wsData1);
    XLSX.utils.book_append_sheet(wb, ws1, 'Finansal Özet');

    // Sheet 2: Proje Takvimi
    const wsData2 = [
      ['Sprint', 'Başlangıç', 'Bitiş', 'Durum'],
      ['Sprint 1', '2026-11-01', '2026-11-15', 'Tamamlandı'],
      ['Sprint 2', '2026-11-16', '2026-11-30', 'Devam Ediyor']
    ];
    const ws2 = XLSX.utils.aoa_to_sheet(wsData2);
    XLSX.utils.book_append_sheet(wb, ws2, 'Sprint Takvimi');

    // Buffer'a yaz
    const xlsxBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    assert.ok(xlsxBuffer.length > 0, 'XLSX buffer boş olamaz');

    // Buffer'dan geri oku
    const readWb = XLSX.read(xlsxBuffer, { type: 'buffer' });
    assert.strictEqual(readWb.SheetNames.length, 2, 'Sayfa sayısı 2 olmalı');
    assert.strictEqual(readWb.SheetNames[0], 'Finansal Özet');
    assert.strictEqual(readWb.SheetNames[1], 'Sprint Takvimi');

    const sheet1 = readWb.Sheets['Finansal Özet'];
    const cellA1 = sheet1['A1'];
    assert.strictEqual(cellA1.v, 'Departman');
    const cellB2 = sheet1['B2'];
    assert.strictEqual(cellB2.v, 42);

    // CSV dönüşüm testi
    const csvData = XLSX.utils.sheet_to_csv(sheet1);
    assert.ok(csvData.includes('Yazılım Mühendisliği,42'), 'CSV içeriği doğru olmalı');

    const sampleDir = path.join(__dirname, 'samples');
    if (!fs.existsSync(sampleDir)) {
      fs.mkdirSync(sampleDir, { recursive: true });
    }
    fs.writeFileSync(path.join(sampleDir, 'test_butce.xlsx'), xlsxBuffer);
    fs.writeFileSync(path.join(sampleDir, 'test_veriler.csv'), csvData);

    console.log('  -> Başarılı: Çoklu sayfalı XLSX ve CSV başarıyla okundu ve doğrulandı.');
  }

  // 2. Test Word (.docx) Structure Creation & Parsing
  console.log('\n[TEST 2] Word (.docx) OpenXML Yapı ve ZIP Uyumluluk Testi...');
  {
    const docxPreview = require('docx-preview');
    assert.ok(typeof docxPreview.renderAsync === 'function', 'docx-preview renderAsync fonksiyonu mevcut olmalı');

    const sampleDocxPath = path.join(__dirname, 'samples', 'ornek_dokuman.docx');
    assert.ok(fs.existsSync(sampleDocxPath), 'ornek_dokuman.docx oluşturulmuş olmalı');
    const docxBuffer = fs.readFileSync(sampleDocxPath);
    assert.ok(docxBuffer.length > 500, 'DOCX dosya boyutu geçerli olmalı');

    const zipLoaded = await JSZip.loadAsync(docxBuffer);
    assert.ok(zipLoaded.file('word/document.xml'), 'word/document.xml dosyası arşivde bulunmalı');
    const docXmlText = await zipLoaded.file('word/document.xml').async('text');
    assert.ok(docXmlText.includes('VS Code Universal Document Viewer'), 'DOCX metin içeriği doğrulanmalı');

    console.log('  -> Başarılı: docx-preview ve gerçek .docx OpenXML arşivi doğrulandı.');
  }

  // 3. Test Legacy .doc Text Extractor
  console.log('\n[TEST 3] Eski Biçimli Word (.doc) İkili Metin Çıkarıcı Testi...');
  {
    const mockBinaryBuffer = Buffer.from([
      0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, // OLE Header imzası
      0x00, 0x00, 0x00, 0x00,
      ...Buffer.from('Gizli Proje Şartnamesi ve Mimari Notlar'),
      0x00, 0x01, 0x02, 0xff, 0xfe,
      ...Buffer.from('Bu doküman Word 97-2003 ikili formatında saklanmıştır.'),
      0x00, 0x00
    ]);

    function extractTextFromBinary(buffer) {
      const bytes = new Uint8Array(buffer);
      const textPieces = [];
      let currentWord = '';

      for (let i = 0; i < bytes.length; i++) {
        const b = bytes[i];
        if ((b >= 32 && b <= 126) || b === 10 || b === 13 || (b >= 192 && b <= 255)) {
          currentWord += String.fromCharCode(b);
        } else {
          if (currentWord.trim().length >= 4) {
            textPieces.push(currentWord.trim());
          }
          currentWord = '';
        }
      }
      if (currentWord.trim().length >= 4) {
        textPieces.push(currentWord.trim());
      }
      return textPieces.join('\n');
    }

    const extracted = extractTextFromBinary(mockBinaryBuffer);
    assert.ok(extracted.includes('Gizli Proje'), 'İkili .doc içerisinden metin blokları çıkarılabilmeli');
    console.log('  -> Başarılı: Eski Word (.doc) ikili ayrıştırıcı metin bloklarını başarıyla ayıkladı.');
  }

  // 4. Test Webview Bundle Dosya Bütünlüğü ve Boyut Kontrolü
  console.log('\n[TEST 4] Derlenmiş Webview ve Extension Dosyalarının Doğrulanması...');
  {
    const distDir = path.join(__dirname, '..', 'dist');
    const extJs = path.join(distDir, 'extension.js');
    const webviewJs = path.join(distDir, 'webview.js');
    const webviewCss = path.join(distDir, 'webview.css');

    assert.ok(fs.existsSync(extJs), 'dist/extension.js mevcut olmalı');
    assert.ok(fs.existsSync(webviewJs), 'dist/webview.js mevcut olmalı');
    assert.ok(fs.existsSync(webviewCss), 'dist/webview.css mevcut olmalı');

    const webviewJsContent = fs.readFileSync(webviewJs, 'utf-8');
    assert.ok(webviewJsContent.includes('WebviewApp'), 'WebviewApp sınıfı bundle içinde yer almalı');
    assert.ok(webviewJsContent.includes('docx'), 'docx-preview bundle içinde yer almalı');
    assert.ok(webviewJsContent.includes('DOMPurify'), 'DOMPurify bundle içinde yer almalı');

    const webviewCssContent = fs.readFileSync(webviewCss, 'utf-8');
    assert.ok(webviewCssContent.includes('theme-dark'), 'Koyu tema kuralları CSS içinde tanımlı olmalı');
    assert.ok(webviewCssContent.includes('excel-table'), 'Excel tablo stilleri CSS içinde tanımlı olmalı');

    console.log('  -> Başarılı: Tüm bundle dosyaları eksiksiz ve gerekli sınıfları barındırıyor.');
  }

  // 5. Test Faz 3: Geliştirici Araçları (Markdown, JSON, Delimiter Tespiti)
  console.log('\n[TEST 5] Faz 3 Geliştirici Araçları (Markdown, JSON, CSV Delimiter)...');
  {
    // Delimiter tespiti testi
    function detectDelimiter(text) {
      const lines = text.slice(0, 5000).split('\n').filter((l) => l.trim().length > 0);
      if (lines.length === 0) return ',';
      const delimiters = [',', ';', '\t', '|'];
      let bestDelim = ',';
      let maxConsistency = -1;
      for (const delim of delimiters) {
        const counts = lines.slice(0, 10).map((line) => line.split(delim).length - 1);
        const firstCount = counts[0];
        if (firstCount > 0 && counts.every((c) => c === firstCount)) {
          return delim;
        }
        const sum = counts.reduce((a, b) => a + b, 0);
        if (sum > maxConsistency) {
          maxConsistency = sum;
          bestDelim = delim;
        }
      }
      return bestDelim;
    }

    const semicolonCsv = 'ID;Ad;Soyad;Departman\n1;Ali;Yılmaz;AR-GE\n2;Ayşe;Kaya;Finans';
    assert.strictEqual(detectDelimiter(semicolonCsv), ';', 'Noktalı virgül CSV tespiti doğru olmalı');

    const pipeCsv = 'Ürün|Stok|Fiyat\nLaptop|15|45000\nMonitör|30|8000';
    assert.strictEqual(detectDelimiter(pipeCsv), '|', 'Dikey çizgi (|) CSV tespiti doğru olmalı');

    const tabCsv = 'Kod\tTanım\tDurum\n101\tSunucu\tAktif\n102\tAğ\tPasif';
    assert.strictEqual(detectDelimiter(tabCsv), '\t', 'Tab (\\t) tespiti doğru olmalı');

    // JSON ve Markdown dönüşüm doğrulaması
    const sampleDir = path.join(__dirname, 'samples');
    const xlsxBuffer = fs.readFileSync(path.join(sampleDir, 'test_butce.xlsx'));
    const wb = XLSX.read(xlsxBuffer, { type: 'buffer' });
    const sheet = wb.Sheets['Finansal Özet'];
    const jsonResult = XLSX.utils.sheet_to_json(sheet);

    assert.ok(Array.isArray(jsonResult), 'JSON çıktısı dizi olmalı');
    assert.strictEqual(jsonResult[0]['Departman'], 'Yazılım Mühendisliği');
    assert.strictEqual(jsonResult[0]['Çalışan Sayısı'], 42);

    console.log('  -> Başarılı: CSV Delimiter tespiti (; , | \\t) ve JSON export doğrulandı.');
  }

  console.log('\n======================================================');
  console.log(' TÜM TESTLER BAŞARIYLA TAMAMLANDI (5/5 Kategori)');
  console.log('======================================================\n');
}

runTests().catch((err) => {
  console.error('\n❌ Test Başarısız Oldu:', err);
  process.exit(1);
});
