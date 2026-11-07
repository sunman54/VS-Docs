const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

async function setupSamples() {
  const sampleDir = path.join(__dirname, 'samples');
  if (!fs.existsSync(sampleDir)) {
    fs.mkdirSync(sampleDir, { recursive: true });
  }

  // 1. Create a minimal valid PDF
  const minimalPdf = `%PDF-1.4
1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj
3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 400 200] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >> endobj
4 0 obj << /Length 55 >> stream
BT /F1 20 Tf 40 120 Td (VS Code Universal Document Viewer) Tj ET
endstream endobj
5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj
xref
0 6
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000244 00000 n 
0000000350 00000 n 
trailer << /Size 6 /Root 1 0 R >>
startxref
427
%%EOF`;

  fs.writeFileSync(path.join(sampleDir, 'ornek_dokuman.pdf'), minimalPdf);
  console.log('Sample PDF oluşturuldu:', path.join(sampleDir, 'ornek_dokuman.pdf'));

  // 2. Add an embedded image to ornek_dokuman.docx
  const docxPath = path.join(sampleDir, 'ornek_dokuman.docx');
  if (fs.existsSync(docxPath)) {
    const docxBuffer = fs.readFileSync(docxPath);
    const zip = await JSZip.loadAsync(docxBuffer);

    // Minimal 1x1 transparent PNG bytes
    const pngBytes = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
      'base64'
    );
    zip.folder('word').folder('media').file('logo.png', pngBytes);

    const updatedBuffer = await zip.generateAsync({ type: 'nodebuffer' });
    fs.writeFileSync(docxPath, updatedBuffer);
    console.log('Sample DOCX içine test görseli (word/media/logo.png) eklendi.');
  }

  // 3. Create a sample .rtf
  const rtfContent = `{\\rtf1\\ansi\\deff0
{\\fonttbl{\\f0 Courier;}}
\\viewkind4\\uc1\\pard\\f0\\fs20
VS Code Universal Document Viewer - RTF Test Dokumani\\par
Bu bir Zengin Metin Formati (.rtf) testidir.\\par
}`;
  fs.writeFileSync(path.join(sampleDir, 'ornek_metin.rtf'), rtfContent);
  console.log('Sample RTF oluşturuldu:', path.join(sampleDir, 'ornek_metin.rtf'));
}

setupSamples().catch(console.error);

