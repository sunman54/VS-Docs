import { ThemeManager, VsCodeThemeKind } from './themeManager';
import { renderDocx, adjustDocxColorsForTheme } from './renderers/docxRenderer';
import { ExcelViewer } from './renderers/excelRenderer';
import { renderLegacyDoc } from './renderers/legacyDocRenderer';
import { PdfViewer } from './renderers/pdfRenderer';
import { renderOdt, renderRtf } from './renderers/odtRenderer';
import { OutlineManager } from './outlineManager';
import { extractImagesFromDocx } from './imageExporter';

declare function acquireVsCodeApi(): {
  postMessage: (msg: any) => void;
  getState: () => any;
  setState: (state: any) => void;
};

class WebviewApp {
  private vscode = acquireVsCodeApi();
  private themeManager: ThemeManager;
  private excelViewer: ExcelViewer | null = null;
  private pdfViewer: PdfViewer | null = null;
  private outlineManager: OutlineManager | null = null;

  // DOM Elements
  private docTitleEl = document.getElementById('doc-title');
  private docBadgeEl = document.getElementById('doc-badge');
  private loadingOverlay = document.getElementById('loading-overlay');
  private errorBanner = document.getElementById('error-banner');
  private errorMessage = document.getElementById('error-message');
  private btnCloseError = document.getElementById('btn-close-error');
  private viewerContainer = document.getElementById('viewer-container');
  private documentContent = document.getElementById('document-content');

  // Status Bar
  private statusSelectionEl = document.getElementById('status-selection');
  private statusStatsEl = document.getElementById('status-stats');

  // Zoom elements
  private zoomControls = document.getElementById('zoom-controls');
  private btnZoomIn = document.getElementById('btn-zoom-in');
  private btnZoomOut = document.getElementById('btn-zoom-out');
  private btnZoomReset = document.getElementById('btn-zoom-reset');
  private zoomLevelEl = document.getElementById('zoom-level');
  private currentZoom = 1.0;

  // Search elements
  private searchInput = document.getElementById('search-input') as HTMLInputElement | null;
  private searchCountEl = document.getElementById('search-count');

  // CSV Delimiter elements
  private delimiterContainer = document.getElementById('delimiter-selector-container');
  private delimiterSelect = document.getElementById('delimiter-select') as HTMLSelectElement | null;

  // Export & Action elements
  private btnCopyMd = document.getElementById('btn-copy-md');
  private btnCopyJson = document.getElementById('btn-copy-json');
  private btnExportImages = document.getElementById('btn-export-images');
  private btnPrint = document.getElementById('btn-print');
  private btnCloseOutline = document.getElementById('btn-close-outline');

  // Current doc state
  private currentExt = '';
  private currentBuffer: ArrayBuffer | null = null;
  private currentFileName = '';

  constructor() {
    this.themeManager = new ThemeManager();

    if (this.documentContent) {
      this.excelViewer = new ExcelViewer(this.documentContent, {
        onCellSelected: (addr, val) => {
          if (this.statusSelectionEl) {
            this.statusSelectionEl.textContent = `Hücre ${addr}: ${val}`;
          }
        },
        onStatsUpdated: (info) => {
          if (this.statusStatsEl) {
            this.statusStatsEl.textContent = info;
          }
        },
        onSelectionStats: (stats) => {
          if (this.statusSelectionEl) {
            this.statusSelectionEl.textContent = stats;
          }
        },
      });

      this.pdfViewer = new PdfViewer(this.documentContent, {
        onPageCount: (total) => {
          if (this.statusStatsEl) {
            this.statusStatsEl.textContent = `Toplam ${total} sayfa`;
          }
        },
      });

      this.outlineManager = new OutlineManager(this.documentContent);
    }

    this.initEventListeners();
    this.initMessageListener();

    // Signal extension host that webview is ready
    this.vscode.postMessage({ command: 'webviewReady' });
  }

  private initEventListeners() {
    // Error banner close
    this.btnCloseError?.addEventListener('click', () => {
      this.errorBanner?.classList.add('hidden');
    });

    // Outline close
    this.btnCloseOutline?.addEventListener('click', () => {
      this.outlineManager?.close();
    });

    // Zoom controls
    this.btnZoomIn?.addEventListener('click', () => this.adjustZoom(0.1));
    this.btnZoomOut?.addEventListener('click', () => this.adjustZoom(-0.1));
    this.btnZoomReset?.addEventListener('click', () => this.resetZoom());

    // Print button
    this.btnPrint?.addEventListener('click', () => {
      window.print();
    });

    // Copy as Markdown
    this.btnCopyMd?.addEventListener('click', () => {
      this.handleCopyMarkdown();
    });

    // Copy as JSON
    this.btnCopyJson?.addEventListener('click', () => {
      this.handleCopyJson();
    });

    // Export Images from Word
    this.btnExportImages?.addEventListener('click', async () => {
      await this.handleExportImages();
    });

    // CSV Delimiter change
    this.delimiterSelect?.addEventListener('change', (e) => {
      const val = (e.target as HTMLSelectElement).value;
      if (this.excelViewer && ['csv', 'tsv'].includes(this.currentExt)) {
        this.excelViewer.reparseCsvWithDelimiter(val);
      }
    });

    // Search input
    let searchDebounceTimer: any = null;
    this.searchInput?.addEventListener('input', () => {
      clearTimeout(searchDebounceTimer);
      searchDebounceTimer = setTimeout(() => {
        this.performSearch(this.searchInput?.value || '');
      }, 250);
    });
  }

  private initMessageListener() {
    window.addEventListener('message', async (event) => {
      const message = event.data;
      switch (message.command) {
        case 'renderDocument':
          this.currentBuffer = message.buffer;
          this.currentFileName = message.fileName;
          await this.handleRenderDocument(
            message.fileName,
            message.fileExt,
            message.buffer,
            message.theme
          );
          break;
        case 'themeChanged':
          this.themeManager.setVsCodeTheme(message.theme as VsCodeThemeKind);
          const isDark = this.themeManager.getEffectiveTheme() !== 'light';

          if (this.currentExt === 'docx' && this.documentContent) {
            adjustDocxColorsForTheme(this.documentContent, isDark);
          } else if (this.currentExt === 'pdf' && this.pdfViewer) {
            this.pdfViewer.setDarkMode(isDark);
          }
          break;
        case 'showError':
          this.showError(message.message);
          this.hideLoading();
          break;
      }
    });
  }

  private async handleRenderDocument(
    fileName: string,
    fileExt: string,
    buffer: ArrayBuffer,
    themeKind: VsCodeThemeKind
  ) {
    this.showLoading('Doküman açılıyor...');
    this.currentExt = fileExt.toLowerCase();

    // Update Header Info
    if (this.docTitleEl) {
      this.docTitleEl.textContent = fileName;
      this.docTitleEl.title = fileName;
    }

    if (this.docBadgeEl) {
      this.docBadgeEl.textContent = this.currentExt.toUpperCase();
      this.docBadgeEl.className = 'doc-badge';
      if (['docx', 'doc'].includes(this.currentExt)) {
        this.docBadgeEl.classList.add(this.currentExt === 'doc' ? 'legacy' : 'word');
      } else if (['xlsx', 'xls', 'xlsm', 'csv', 'tsv', 'ods'].includes(this.currentExt)) {
        this.docBadgeEl.classList.add('excel');
      } else if (this.currentExt === 'pdf') {
        this.docBadgeEl.classList.add('pdf');
      } else if (['odt', 'rtf'].includes(this.currentExt)) {
        this.docBadgeEl.classList.add('odt');
      }
    }

    // Apply Theme
    this.themeManager.setVsCodeTheme(themeKind);
    const isDark = this.themeManager.getEffectiveTheme() !== 'light';

    try {
      if (!this.documentContent) {
        throw new Error('İçerik konteyneri bulunamadı.');
      }

      this.resetZoom();
      this.btnExportImages?.classList.add('hidden');
      this.outlineManager?.close();

      // Configure Delimiter Dropdown Visibility
      if (['csv', 'tsv'].includes(this.currentExt)) {
        this.delimiterContainer?.classList.remove('hidden');
      } else {
        this.delimiterContainer?.classList.add('hidden');
      }

      // Route by file extension
      if (this.currentExt === 'docx') {
        this.zoomControls?.classList.remove('hidden');
        if (this.statusStatsEl) this.statusStatsEl.textContent = 'Word Dokümanı';
        if (this.statusSelectionEl) this.statusSelectionEl.textContent = 'Hazır';

        await renderDocx(buffer, this.documentContent, isDark);

        // Update Outline Headings & Image button
        const headingsCount = this.outlineManager?.updateHeadings() || 0;
        if (headingsCount > 0 && this.statusStatsEl) {
          this.statusStatsEl.textContent = `Word Dokümanı (${headingsCount} başlık)`;
        }

        // Check if docx has images
        extractImagesFromDocx(buffer).then((images) => {
          if (images.length > 0) {
            this.btnExportImages?.classList.remove('hidden');
            if (this.btnExportImages) {
              this.btnExportImages.title = `Dokümandaki ${images.length} görseli dışa aktar`;
            }
          }
        }).catch(() => {});

      } else if (this.currentExt === 'pdf') {
        this.zoomControls?.classList.remove('hidden');
        if (this.pdfViewer) {
          await this.pdfViewer.render(buffer, isDark);
        }
      } else if (['xlsx', 'xls', 'xlsm', 'csv', 'tsv', 'ods'].includes(this.currentExt)) {
        this.zoomControls?.classList.add('hidden');
        if (this.excelViewer) {
          const isCsv = ['csv', 'tsv'].includes(this.currentExt);
          this.excelViewer.render(buffer, isCsv);
        }
      } else if (this.currentExt === 'odt') {
        this.zoomControls?.classList.remove('hidden');
        await renderOdt(buffer, this.documentContent);
        this.outlineManager?.updateHeadings();
      } else if (this.currentExt === 'rtf') {
        this.zoomControls?.classList.add('hidden');
        renderRtf(buffer, this.documentContent);
      } else if (this.currentExt === 'doc') {
        this.zoomControls?.classList.add('hidden');
        if (this.statusStatsEl) this.statusStatsEl.textContent = 'Word 97-2003 (.doc)';
        if (this.statusSelectionEl) this.statusSelectionEl.textContent = 'Salt Okunur Metin Modu';
        renderLegacyDoc(buffer, this.documentContent, fileName);
      } else {
        throw new Error(`Desteklenmeyen dosya uzantısı: .${this.currentExt}`);
      }

      this.hideLoading();
    } catch (err: any) {
      console.error('Render error:', err);
      this.showError(`Görüntüleme hatası: ${err?.message || String(err)}`);
      this.hideLoading();
    }
  }

  private async handleExportImages() {
    if (!this.currentBuffer || this.currentExt !== 'docx') return;
    try {
      this.showLoading('Görseller ayıklanıyor...');
      const images = await extractImagesFromDocx(this.currentBuffer);
      this.hideLoading();

      if (images.length === 0) {
        this.vscode.postMessage({
          command: 'showInfoMessage',
          text: 'Bu dokümanda dışa aktarılacak görsel bulunamadı.',
        });
        return;
      }

      this.vscode.postMessage({
        command: 'exportImages',
        images: images,
      });
    } catch (err: any) {
      this.hideLoading();
      this.showError(`Görsel ayıklama hatası: ${err?.message || String(err)}`);
    }
  }

  private handleCopyMarkdown() {
    if (this.excelViewer && ['xlsx', 'xls', 'xlsm', 'csv', 'tsv', 'ods'].includes(this.currentExt)) {
      const md = this.excelViewer.exportAsMarkdown();
      if (md) {
        this.vscode.postMessage({
          command: 'copyToClipboard',
          text: md,
          infoMessage: 'Excel tablosu panoya Markdown olarak kopyalandı! 📋',
        });
      }
    } else if (this.currentExt === 'docx' && this.documentContent) {
      const table = this.documentContent.querySelector('table');
      if (table) {
        const md = this.convertHtmlTableToMarkdown(table);
        this.vscode.postMessage({
          command: 'copyToClipboard',
          text: md,
          infoMessage: 'Word tablosu panoya Markdown olarak kopyalandı! 📋',
        });
      } else {
        this.vscode.postMessage({
          command: 'showInfoMessage',
          text: 'Dokümanda kopyalanacak bir tablo bulunamadı.',
        });
      }
    }
  }

  private handleCopyJson() {
    if (this.excelViewer && ['xlsx', 'xls', 'xlsm', 'csv', 'tsv', 'ods'].includes(this.currentExt)) {
      const json = this.excelViewer.exportAsJson();
      if (json && json !== '[]') {
        this.vscode.postMessage({
          command: 'copyToClipboard',
          text: json,
          infoMessage: 'Veriler JSON dizisi olarak panoya kopyalandı! 📋',
        });
      }
    }
  }

  private convertHtmlTableToMarkdown(table: HTMLTableElement): string {
    const rows = Array.from(table.querySelectorAll('tr'));
    if (rows.length === 0) return '';

    const lines: string[] = [];
    const headerCells = Array.from(rows[0].querySelectorAll('th, td'));
    const headers = headerCells.map((c) => (c.textContent || '').replace(/\|/g, '\\|').trim());
    lines.push(`| ${headers.join(' | ')} |`);
    lines.push(`| ${headers.map(() => '---').join(' | ')} |`);

    for (let i = 1; i < rows.length; i++) {
      const cells = Array.from(rows[i].querySelectorAll('td'));
      const values = cells.map((c) => (c.textContent || '').replace(/\|/g, '\\|').trim());
      lines.push(`| ${values.join(' | ')} |`);
    }

    return lines.join('\n');
  }

  private adjustZoom(delta: number) {
    this.currentZoom = Math.min(2.5, Math.max(0.4, Math.round((this.currentZoom + delta) * 10) / 10));
    this.applyZoom();
  }

  private resetZoom() {
    this.currentZoom = 1.0;
    this.applyZoom();
  }

  private applyZoom() {
    if (this.documentContent) {
      this.documentContent.style.transform = `scale(${this.currentZoom})`;
    }
    if (this.zoomLevelEl) {
      this.zoomLevelEl.textContent = `${Math.round(this.currentZoom * 100)}%`;
    }
  }

  private performSearch(query: string) {
    if (!query.trim()) {
      if (this.searchCountEl) this.searchCountEl.textContent = '';
      if (this.excelViewer && ['xlsx', 'xls', 'xlsm', 'csv', 'tsv', 'ods'].includes(this.currentExt)) {
        this.excelViewer.filterCells('');
      }
      return;
    }

    if (this.excelViewer && ['xlsx', 'xls', 'xlsm', 'csv', 'tsv', 'ods'].includes(this.currentExt)) {
      const count = this.excelViewer.filterCells(query);
      if (this.searchCountEl) {
        this.searchCountEl.textContent = `${count} eşleşme`;
      }
    } else {
      const win = window as any;
      if (typeof win.find === 'function') {
        const found = win.find(query, false, false, true, false, true, false);
        if (this.searchCountEl) {
          this.searchCountEl.textContent = found ? 'Bulundu' : 'Bulunamadı';
        }
      }
    }
  }

  private showLoading(text?: string) {
    if (this.loadingOverlay) {
      this.loadingOverlay.classList.remove('hidden');
      const textEl = document.getElementById('loading-text');
      if (textEl && text) textEl.textContent = text;
    }
  }

  private hideLoading() {
    if (this.loadingOverlay) {
      this.loadingOverlay.classList.add('hidden');
    }
    document.body.classList.remove('loading');
  }

  private showError(msg: string) {
    if (this.errorBanner && this.errorMessage) {
      this.errorMessage.textContent = msg;
      this.errorBanner.classList.remove('hidden');
    }
  }
}

// Bootstrap
window.addEventListener('DOMContentLoaded', () => {
  new WebviewApp();
});
