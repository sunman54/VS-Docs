import JSZip from 'jszip';
import DOMPurify from 'dompurify';

export async function renderOdt(buffer: ArrayBuffer, container: HTMLElement): Promise<void> {
  container.innerHTML = '';

  const zip = await JSZip.loadAsync(buffer);
  const contentFile = zip.file('content.xml');
  if (!contentFile) {
    throw new Error('Geçersiz .odt dosyası: content.xml bulunamadı.');
  }

  const xmlText = await contentFile.async('text');
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlText, 'application/xml');

  const wrapper = document.createElement('div');
  wrapper.className = 'docx-wrapper';

  const section = document.createElement('section');
  section.className = 'docx';
  section.style.padding = '40px';
  section.style.maxWidth = '800px';
  section.style.width = '100%';

  // Extract paragraphs and headings
  const body = xmlDoc.getElementsByTagName('office:text')[0];
  if (body) {
    for (let i = 0; i < body.children.length; i++) {
      const node = body.children[i];
      const tag = node.tagName.toLowerCase();

      if (tag.endsWith(':h')) {
        const level = node.getAttribute('text:outline-level') || '1';
        const h = document.createElement(`h${Math.min(parseInt(level, 10), 6)}`);
        h.textContent = node.textContent || '';
        section.appendChild(h);
      } else if (tag.endsWith(':p')) {
        const p = document.createElement('p');
        p.textContent = node.textContent || '';
        p.style.marginBottom = '12px';
        section.appendChild(p);
      } else if (tag.endsWith(':table')) {
        const table = document.createElement('table');
        table.className = 'excel-table';
        const rows = node.getElementsByTagName('table:table-row');
        for (let r = 0; r < rows.length; r++) {
          const tr = document.createElement('tr');
          const cells = rows[r].getElementsByTagName('table:table-cell');
          for (let c = 0; c < cells.length; c++) {
            const td = document.createElement('td');
            td.textContent = cells[c].textContent || '';
            tr.appendChild(td);
          }
          table.appendChild(tr);
        }
        section.appendChild(table);
      }
    }
  }

  wrapper.appendChild(section);
  container.appendChild(wrapper);
}

export function renderRtf(buffer: ArrayBuffer, container: HTMLElement): void {
  container.innerHTML = '';

  const decoder = new TextDecoder('utf-8');
  const rawText = decoder.decode(buffer);

  // Simple RTF text stripper
  let cleanText = rawText
    .replace(/\\par[d]?/g, '\n')
    .replace(/\\tab/g, '\t')
    .replace(/\\[a-z0-9\-]+/gi, '')
    .replace(/[{}]/g, '')
    .replace(/\r\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  cleanText = DOMPurify.sanitize(cleanText);

  const wrapper = document.createElement('div');
  wrapper.style.maxWidth = '800px';
  wrapper.style.width = '100%';
  wrapper.style.padding = '24px';

  const pre = document.createElement('pre');
  pre.className = 'legacy-extracted-text';
  pre.textContent = cleanText || '(RTF metin içeriği boş)';

  wrapper.appendChild(pre);
  container.appendChild(wrapper);
}

