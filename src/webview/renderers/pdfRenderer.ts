import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf';

// Configure pdfjs worker
(pdfjsLib as any).GlobalWorkerOptions.workerSrc = '';

export interface PdfRenderCallbacks {
  onPageCount?: (total: number) => void;
  onPageVisible?: (pageNum: number, total: number) => void;
}

export class PdfViewer {
  private container: HTMLElement;
  private pdfDoc: any = null;
  private currentScale: number = 1.2;
  private isDarkMode: boolean = false;
  private callbacks: PdfRenderCallbacks;

  constructor(container: HTMLElement, callbacks: PdfRenderCallbacks = {}) {
    this.container = container;
    this.callbacks = callbacks;
  }

  public async render(buffer: ArrayBuffer, isDark: boolean = false): Promise<void> {
    this.container.innerHTML = '';
    this.isDarkMode = isDark;

    const loadingTask = pdfjsLib.getDocument({
      data: new Uint8Array(buffer),
      isEvalSupported: false,
      useWorkerFetch: false,
    });

    this.pdfDoc = await loadingTask.promise;
    const numPages = this.pdfDoc.numPages;
    this.callbacks.onPageCount?.(numPages);

    const pdfWrapper = document.createElement('div');
    pdfWrapper.className = `pdf-document-wrapper ${this.isDarkMode ? 'dark-filter' : ''}`;
    pdfWrapper.id = 'pdf-wrapper';
    this.container.appendChild(pdfWrapper);

    for (let pageNum = 1; pageNum <= numPages; pageNum++) {
      const page = await this.pdfDoc.getPage(pageNum);
      const viewport = page.getViewport({ scale: this.currentScale });

      const pageContainer = document.createElement('div');
      pageContainer.className = 'pdf-page-container';
      pageContainer.dataset.pageNumber = String(pageNum);

      const canvas = document.createElement('canvas');
      canvas.className = 'pdf-canvas';
      const context = canvas.getContext('2d');

      canvas.height = viewport.height;
      canvas.width = viewport.width;

      pageContainer.appendChild(canvas);
      pdfWrapper.appendChild(pageContainer);

      if (context) {
        await page.render({
          canvasContext: context,
          viewport: viewport,
        }).promise;
      }
    }
  }

  public setScale(scale: number) {
    this.currentScale = scale;
    const wrapper = document.getElementById('pdf-wrapper');
    if (wrapper) {
      wrapper.style.transform = `scale(${this.currentScale / 1.2})`;
      wrapper.style.transformOrigin = 'top center';
    }
  }

  public setDarkMode(isDark: boolean) {
    this.isDarkMode = isDark;
    const wrapper = document.getElementById('pdf-wrapper');
    if (wrapper) {
      if (isDark) {
        wrapper.classList.add('dark-filter');
      } else {
        wrapper.classList.remove('dark-filter');
      }
    }
  }
}
