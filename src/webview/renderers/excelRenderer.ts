import * as XLSX from 'xlsx';
import DOMPurify from 'dompurify';

export interface ExcelRenderCallbacks {
  onCellSelected?: (cellAddress: string, value: string) => void;
  onStatsUpdated?: (info: string) => void;
}

export class ExcelViewer {
  private workbook: XLSX.WorkBook | null = null;
  private currentSheetName: string = '';
  private container: HTMLElement;
  private tableContainer: HTMLElement | null = null;
  private tabsBar: HTMLElement | null = null;
  private sheetSelect: HTMLSelectElement | null = null;
  private maxInitialRows = 1000;
  private sortColumnIndex: number | null = null;
  private sortAscending: boolean = true;
  private callbacks: ExcelRenderCallbacks;

  constructor(container: HTMLElement, callbacks: ExcelRenderCallbacks = {}) {
    this.container = container;
    this.callbacks = callbacks;
    this.sheetSelect = document.getElementById('sheet-select') as HTMLSelectElement | null;
  }

  public render(buffer: ArrayBuffer) {
    this.workbook = XLSX.read(new Uint8Array(buffer), {
      type: 'array',
      cellDates: true,
      cellStyles: true,
      dense: false,
    });

    if (!this.workbook.SheetNames || this.workbook.SheetNames.length === 0) {
      throw new Error('Çalışma kitabında görüntülenecek sayfa bulunamadı.');
    }

    this.container.innerHTML = '';

    // Create wrapper layout
    const wrapper = document.createElement('div');
    wrapper.className = 'excel-view-container';

    this.tableContainer = document.createElement('div');
    this.tableContainer.className = 'excel-grid-wrapper';
    wrapper.appendChild(this.tableContainer);

    this.tabsBar = document.createElement('div');
    this.tabsBar.className = 'sheet-tabs-bar';
    wrapper.appendChild(this.tabsBar);

    this.container.appendChild(wrapper);

    // Setup navigation
    this.setupSheetNavigation();

    // Render first sheet by default
    this.switchSheet(this.workbook.SheetNames[0]);
  }

  private setupSheetNavigation() {
    if (!this.workbook || !this.tabsBar) return;

    if (this.sheetSelect) {
      this.sheetSelect.innerHTML = '';
      const sheetContainer = document.getElementById('sheet-selector-container');
      if (sheetContainer) {
        sheetContainer.classList.remove('hidden');
      }
    }

    this.tabsBar.innerHTML = '';

    this.workbook.SheetNames.forEach((sheetName) => {
      // 1. Populate top toolbar dropdown
      if (this.sheetSelect) {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        this.sheetSelect.appendChild(option);
      }

      // 2. Populate bottom tabs bar
      const tab = document.createElement('button');
      tab.className = 'sheet-tab';
      tab.textContent = sheetName;
      tab.dataset.sheetName = sheetName;
      tab.addEventListener('click', () => {
        this.switchSheet(sheetName);
      });
      this.tabsBar!.appendChild(tab);
    });

    if (this.sheetSelect) {
      this.sheetSelect.onchange = (e) => {
        const target = e.target as HTMLSelectElement;
        this.switchSheet(target.value);
      };
    }
  }

  public switchSheet(sheetName: string) {
    if (!this.workbook || !this.tableContainer) return;
    this.currentSheetName = sheetName;
    this.sortColumnIndex = null;
    this.sortAscending = true;

    // Update dropdown
    if (this.sheetSelect && this.sheetSelect.value !== sheetName) {
      this.sheetSelect.value = sheetName;
    }

    // Update bottom tabs active state
    if (this.tabsBar) {
      const tabs = this.tabsBar.querySelectorAll('.sheet-tab');
      tabs.forEach((tab) => {
        const el = tab as HTMLElement;
        if (el.dataset.sheetName === sheetName) {
          el.classList.add('active');
          el.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'nearest' });
        } else {
          el.classList.remove('active');
        }
      });
    }

    // Render sheet table
    const worksheet = this.workbook.Sheets[sheetName];
    this.renderWorksheetTable(worksheet);
  }

  private renderWorksheetTable(sheet: XLSX.WorkSheet) {
    if (!this.tableContainer) return;
    this.tableContainer.innerHTML = '';

    if (!sheet || !sheet['!ref']) {
      this.tableContainer.innerHTML = '<div style="padding: 20px; text-align: center; color: var(--vscode-descriptionForeground);">Bu sayfa boş.</div>';
      this.callbacks.onStatsUpdated?.('0 satır, 0 sütun');
      return;
    }

    const range = XLSX.utils.decode_range(sheet['!ref']);
    const startRow = range.s.r;
    const endRow = range.e.r;
    const startCol = range.s.c;
    const endCol = range.e.c;
    const rowCount = endRow - startRow + 1;
    const colCount = endCol - startCol + 1;

    this.callbacks.onStatsUpdated?.(`${rowCount} satır, ${colCount} sütun | Sayfa: ${this.currentSheetName}`);

    const table = document.createElement('table');
    table.className = 'excel-table';

    // Build Header Row
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');

    const cornerTh = document.createElement('th');
    cornerTh.className = 'corner-header';
    cornerTh.textContent = '';
    headerRow.appendChild(cornerTh);

    for (let c = startCol; c <= endCol; c++) {
      const th = document.createElement('th');
      th.className = 'col-header';
      const colLetter = XLSX.utils.encode_col(c);
      th.textContent = colLetter;
      th.title = `Sütun ${colLetter} (Sıralamak için tıklayın)`;
      th.addEventListener('click', () => {
        this.sortTableByColumn(c, sheet);
      });
      headerRow.appendChild(th);
    }
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // Build Rows
    const tbody = document.createElement('tbody');
    const rowsToRender = Math.min(rowCount, this.maxInitialRows);

    for (let r = startRow; r < startRow + rowsToRender; r++) {
      const tr = document.createElement('tr');

      const rowHeader = document.createElement('td');
      rowHeader.className = 'row-header';
      rowHeader.textContent = String(r + 1);
      tr.appendChild(rowHeader);

      for (let c = startCol; c <= endCol; c++) {
        const cellAddress = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[cellAddress];
        const td = document.createElement('td');
        td.className = 'cell-content';
        td.dataset.addr = cellAddress;

        if (cell) {
          let val = cell.w !== undefined ? String(cell.w) : (cell.v !== undefined ? String(cell.v) : '');
          val = DOMPurify.sanitize(val);
          td.textContent = val;

          if (cell.t === 'n') {
            td.classList.add('type-number');
          }
        } else {
          td.textContent = '';
        }

        // Cell selection listener
        td.addEventListener('click', () => {
          this.selectCell(td, cellAddress, td.textContent || '');
        });

        tr.appendChild(td);
      }

      tbody.appendChild(tr);
    }

    table.appendChild(tbody);
    this.tableContainer.appendChild(table);

    // Load More indicator if sheet has more than 1000 rows
    if (rowCount > this.maxInitialRows) {
      const loadMoreContainer = document.createElement('div');
      loadMoreContainer.style.padding = '12px';
      loadMoreContainer.style.textAlign = 'center';

      const loadMoreBtn = document.createElement('button');
      loadMoreBtn.className = 'toolbar-btn text-btn';
      loadMoreBtn.textContent = `Kalan ${rowCount - this.maxInitialRows} satırı yükle (Toplam ${rowCount} satır)`;
      loadMoreBtn.onclick = () => {
        this.loadRemainingRows(sheet, tbody, startRow + this.maxInitialRows, endRow, startCol, endCol);
        loadMoreContainer.remove();
      };

      loadMoreContainer.appendChild(loadMoreBtn);
      this.tableContainer.appendChild(loadMoreContainer);
    }
  }

  private selectCell(cellEl: HTMLElement, addr: string, value: string) {
    if (!this.tableContainer) return;
    const prev = this.tableContainer.querySelectorAll('.cell-content.selected');
    prev.forEach((el) => el.classList.remove('selected'));

    cellEl.classList.add('selected');
    this.callbacks.onCellSelected?.(addr, value);
  }

  private sortTableByColumn(colIndex: number, sheet: XLSX.WorkSheet) {
    if (!this.tableContainer) return;

    if (this.sortColumnIndex === colIndex) {
      this.sortAscending = !this.sortAscending;
    } else {
      this.sortColumnIndex = colIndex;
      this.sortAscending = true;
    }

    const tbody = this.tableContainer.querySelector('tbody');
    if (!tbody) return;

    const rows = Array.from(tbody.querySelectorAll('tr'));
    rows.sort((a, b) => {
      // index 0 is row-header, so col index is colIndex + 1
      const cellA = a.children[colIndex + 1];
      const cellB = b.children[colIndex + 1];
      const valA = cellA?.textContent?.trim() || '';
      const valB = cellB?.textContent?.trim() || '';

      const numA = parseFloat(valA.replace(/,/g, ''));
      const numB = parseFloat(valB.replace(/,/g, ''));

      let cmp = 0;
      if (!isNaN(numA) && !isNaN(numB)) {
        cmp = numA - numB;
      } else {
        cmp = valA.localeCompare(valB, 'tr');
      }

      return this.sortAscending ? cmp : -cmp;
    });

    tbody.innerHTML = '';
    rows.forEach((r) => tbody.appendChild(r));

    const direction = this.sortAscending ? '▲' : '▼';
    this.callbacks.onStatsUpdated?.(
      `Sıralandı: Sütun ${XLSX.utils.encode_col(colIndex)} (${direction}) | Sayfa: ${this.currentSheetName}`
    );
  }

  private loadRemainingRows(
    sheet: XLSX.WorkSheet,
    tbody: HTMLTableSectionElement,
    fromRow: number,
    toRow: number,
    startCol: number,
    endCol: number
  ) {
    for (let r = fromRow; r <= toRow; r++) {
      const tr = document.createElement('tr');
      const rowHeader = document.createElement('td');
      rowHeader.className = 'row-header';
      rowHeader.textContent = String(r + 1);
      tr.appendChild(rowHeader);

      for (let c = startCol; c <= endCol; c++) {
        const cellAddress = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[cellAddress];
        const td = document.createElement('td');
        td.className = 'cell-content';
        td.dataset.addr = cellAddress;

        if (cell) {
          let val = cell.w !== undefined ? String(cell.w) : (cell.v !== undefined ? String(cell.v) : '');
          val = DOMPurify.sanitize(val);
          td.textContent = val;
          if (cell.t === 'n') {
            td.classList.add('type-number');
          }
        } else {
          td.textContent = '';
        }

        td.addEventListener('click', () => {
          this.selectCell(td, cellAddress, td.textContent || '');
        });

        tr.appendChild(td);
      }

      tbody.appendChild(tr);
    }
  }

  public filterCells(query: string): number {
    if (!this.tableContainer) return 0;
    const cells = this.tableContainer.querySelectorAll('.cell-content');
    let matchCount = 0;

    const lowerQuery = query.trim().toLowerCase();

    cells.forEach((cellEl) => {
      const el = cellEl as HTMLElement;
      const text = el.textContent || '';
      if (lowerQuery.length === 0) {
        el.style.backgroundColor = '';
        return;
      }

      if (text.toLowerCase().includes(lowerQuery)) {
        el.style.backgroundColor = 'var(--highlight-bg)';
        matchCount++;
      } else {
        el.style.backgroundColor = '';
      }
    });

    return matchCount;
  }
}
