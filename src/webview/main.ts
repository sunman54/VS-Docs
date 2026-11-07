import { ThemeManager, VsCodeThemeKind } from './themeManager';
import { renderDocx } from './renderers/docxRenderer';
import { ExcelViewer } from './renderers/excelRenderer';
import { renderLegacyDoc } from './renderers/legacyDocRenderer';

declare function acquireVsCodeApi(): {
  postMessage: (msg: any) => void;
  getState: () => any;
  setState: (state: any) => void;
};

// Main Webview Controller
class WebviewApp {
  private vscode = acquireVsCodeApi();
  private themeManager: ThemeManager;
  private excelViewer: ExcelViewer | null = null;

  // DOM Elements
  private docTitleEl = document.getElementById('doc-title');
  private docBadgeEl = document.getElementById('doc-badge');
  private loadingOverlay = document.getElementById('loading-overlay');
  private errorBanner = document.getElementById('error-banner');
  private errorMessage = document.getElementById('error-message');
  private btnCloseError = document.getElementById('btn-close-error');
  private viewerContainer = document.getElementById('viewer-container');
  private documentContent = document.getElementById('document-content');

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

  // Current doc state
  private currentExt = '';

  constructor() {
    this.themeManager = new ThemeManager();

    if (this.documentContent) {
      this.excelViewer = new ExcelViewer(this.documentContent);
    }

    this.initEventListeners();
    this.initMessageListener();

    // Notify extension host that webview is ready to receive data
    this.vscode.postMessage({ command: 'webviewReady' });
  }

  private initEventListeners() {
    // Error banner close
    this.btnCloseError?.addEventListener('click', () => {
      this.errorBanner?.classList.add('hidden');
    });

    // Zoom controls
    this.btnZoomIn?.addEventListener('click', () => this.adjustZoom(0.1));
    this.btnZoomOut?.addEventListener('click', () => this.adjustZoom(-0.1));
    this.btnZoomReset?.addEventListener('click', () => this.resetZoom());

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
          await this.handleRenderDocument(
            message.fileName,
            message.fileExt,
            message.buffer,
            message.theme
          );
          break;
        case 'themeChanged':
          this.themeManager.setVsCodeTheme(message.theme as VsCodeThemeKind);
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
      } else if (['xlsx', 'xls', 'xlsm', 'csv', 'tsv'].includes(this.currentExt)) {
        this.docBadgeEl.classList.add('excel');
      }
    }

    // Apply Theme
    this.themeManager.setVsCodeTheme(themeKind);

    try {
      if (!this.documentContent) {
        throw new Error('İçerik konteyneri bulunamadı.');
      }

      this.resetZoom();

      // Route by file extension
      if (this.currentExt === 'docx') {
        this.zoomControls?.classList.remove('hidden');
        await renderDocx(buffer, this.documentContent);
      } else if (['xlsx', 'xls', 'xlsm', 'csv', 'tsv'].includes(this.currentExt)) {
        // Zoom is not needed for virtual tables (table has its own scrolling)
        this.zoomControls?.classList.add('hidden');
        if (this.excelViewer) {
          this.excelViewer.render(buffer);
        }
      } else if (this.currentExt === 'doc') {
        this.zoomControls?.classList.add('hidden');
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
      if (this.excelViewer && ['xlsx', 'xls', 'xlsm', 'csv', 'tsv'].includes(this.currentExt)) {
        this.excelViewer.filterCells('');
      }
      return;
    }

    if (this.excelViewer && ['xlsx', 'xls', 'xlsm', 'csv', 'tsv'].includes(this.currentExt)) {
      const count = this.excelViewer.filterCells(query);
      if (this.searchCountEl) {
        this.searchCountEl.textContent = `${count} eşleşme`;
      }
    } else {
      // In Word/Docx, window.find can be used or mark matches
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
