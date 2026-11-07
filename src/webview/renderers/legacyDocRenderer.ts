import DOMPurify from 'dompurify';

export function renderLegacyDoc(
  buffer: ArrayBuffer,
  container: HTMLElement,
  fileName: string
): void {
  container.innerHTML = '';

  const wrapper = document.createElement('div');
  wrapper.style.display = 'flex';
  wrapper.style.flexDirection = 'column';
  wrapper.style.alignItems = 'center';
  wrapper.style.width = '100%';
  wrapper.style.gap = '20px';

  // Notice Card
  const card = document.createElement('div');
  card.className = 'legacy-notice-card';
  card.innerHTML = `
    <h3>📄 Eski Biçimli Word Dokümanı (.doc)</h3>
    <p>
      <strong>${DOMPurify.sanitize(fileName)}</strong> dosyası eski ikili Word 97-2003 biçimindedir.<br>
      Tüm sayfa düzeni, grafik ve gelişmiş stilleri eksiksiz görüntülemek için dokümanı modern <code>.docx</code> formatında kaydetmeniz önerilir.
    </p>
    <div style="font-size: 11px; opacity: 0.8; margin-top: 8px;">Aşağıda dokümandan çıkarılan metin içeriği listelenmektedir:</div>
  `;
  wrapper.appendChild(card);

  // Extract text from binary buffer
  const extractedText = extractTextFromBinary(buffer);
  const textContainer = document.createElement('pre');
  textContainer.className = 'legacy-extracted-text';
  textContainer.textContent = extractedText || '(Dokümandan okunabilir düz metin çıkarılamadı)';

  wrapper.appendChild(textContainer);
  container.appendChild(wrapper);
}

function extractTextFromBinary(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  const textPieces: string[] = [];
  let currentWord = '';

  // Scan for printable ASCII strings (length >= 4)
  for (let i = 0; i < bytes.length; i++) {
    const b = bytes[i];
    // Printable ASCII or Turkish Latin-1 extensions
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

  // Join and clean up excessive blank lines
  return textPieces.join('\n').replace(/\n{3,}/g, '\n\n');
}

