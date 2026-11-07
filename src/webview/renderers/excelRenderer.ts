import * as XLSX from 'xlsx';
import DOMPurify from 'dompurify';

export interface ExcelRenderCallbacks {
  onCellSelected?: (cellAddress: string, value: string) => void;
  onStatsUpdated?: (info: string) => void;
  onSelectionStats?: (stats: string) => void;
}

export class ExcelViewer {
  private workbook: XLSX.WorkBook | null = null;
  private currentSheetName: string = '';
  private rawBuffer: ArrayBuffer | null = null;
  private rawCsvText: string = '';
  private isCsv: boolean = false;
  private container: HTMLElement;
  private tableContainer: HTMLElement | null = null;
  private tabsBar: HTMLElement | null = null;
  private sheetSelect: HTMLSelectElement | null = null;
  private maxInitialRows = 1000;
  private sortColumnIndex: number | null = null;
  private sortAscending: boolean = true;
  private callbacks: ExcelRenderCallbacks;

  // Multi-cell selection state
  private isSelecting: boolean = false;
  private selectStart: { r: number; c: number } | null = null;
  private selectedCells: HTMLElement[] = [];

  // Active filters by column index: colIndex -> Set of allowed values
  private activeFilters: Map<number, Set<string>> = new Map();
  private filterPopup: HTMLElement | null = null;

  constructor(container: HTMLElement, callbacks: ExcelRenderCallbacks = {}) {
    this.container = container;
    this.callbacks = callbacks;
    this.sheetSelect = document.getElementById('sheet-select') as HTMLSelectElement | null;

    // Global listeners
    window.addEventListener('mouseup', () => {
      this.isSelecting = false;
    });

    document.addEventListener('click', (e) => {
      if (this.filterPopup && !this.filterPopup.contains(e.target as Node)) {
        this.closeFilterPopup();
      }
    });
  }

  public render(buffer: ArrayBuffer, isCsv: boolean = false, customDelimiter?: string) {
    this.rawBuffer = buffer;
    this.isCsv = isCsv;
    this.activeFilters.clear();

    if (isCsv) {
      const decoder = new TextDecoder('utf-8');
      this.rawCsvText = decoder.decode(buffer);
      const delimiter = customDelimiter || this.detectDelimiter(this.rawCsvText);
      
      this.workbook = XLSX.read(this.rawCsvText, {
        type: 'string',
        FS: delimiter,
        cellDates: true,
      });
    } else {
      this.workbook = XLSX.read(new Uint8Array(buffer), {
        type: 'array',
        cellDates: true,
        cellStyles: true,
        dense: false,
      });
    }

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

  public reparseCsvWithDelimiter(delimiter: string) {
    if (!this.isCsv || !this.rawCsvText) return;
    this.activeFilters.clear();
    const finalDelim = delimiter === 'auto' ? this.detectDelimiter(this.rawCsvText) : delimiter;
    this.workbook = XLSX.read(this.rawCsvText, {
      type: 'string',
      FS: finalDelim,
      cellDates: true,
    });
    if (this.workbook.SheetNames.length > 0) {
      this.switchSheet(this.workbook.SheetNames[0]);
    }
  }

  public detectDelimiter(text: string): string {
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

  private setupSheetNavigation() {
    if (!this.workbook || !this.tabsBar) return;

    if (this.sheetSelect) {
      this.sheetSelect.innerHTML = '';
      const sheetContainer = document.getElementById('sheet-selector-container');
      if (sheetContainer) {
        if (this.workbook.SheetNames.length > 1) {
          sheetContainer.classList.remove('hidden');
        } else {
          sheetContainer.classList.add('hidden');
        }
      }
    }

    this.tabsBar.innerHTML = '';
    if (this.workbook.SheetNames.length <= 1) {
      this.tabsBar.style.display = 'none';
    } else {
      this.tabsBar.style.display = 'flex';
    }

    this.workbook.SheetNames.forEach((sheetName) => {
      if (this.sheetSelect) {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        this.sheetSelect.appendChild(option);
      }

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
    this.selectedCells = [];
    this.activeFilters.clear();

    if (this.sheetSelect && this.sheetSelect.value !== sheetName) {
      this.sheetSelect.value = sheetName;
    }

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

    // Header Row
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');

    const cornerTh = document.createElement('th');
    cornerTh.className = 'corner-header';
    cornerTh.textContent = '';
    headerRow.appendChild(cornerTh);

    for (let c = startCol; c <= endCol; c++) {
      const th = document.createElement('th');
      th.className = 'col-header';
      th.dataset.c = String(c);
      const colLetter = XLSX.utils.encode_col(c);

      const content = document.createElement('div');
      content.className = 'col-header-content';

      const label = document.createElement('span');
      label.textContent = colLetter;
      label.title = `Sütun ${colLetter} (Sıralamak için tıklayın)`;
      label.addEventListener('click', (e) => {
        e.stopPropagation();
        this.sortTableByColumn(c);
      });
      content.appendChild(label);

      // Filter Button
      const filterBtn = document.createElement('span');
      filterBtn.className = 'col-filter-btn';
      filterBtn.textContent = '▼';
      filterBtn.title = `${colLetter} Sütununu Filtrele`;
      filterBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        this.openFilterPopup(c, filterBtn);
      });
      content.appendChild(filterBtn);

      th.appendChild(content);

      // Resizer Handle
      const resizer = document.createElement('div');
      resizer.className = 'col-resizer';
      this.initColumnResizer(resizer, th, c);
      th.appendChild(resizer);

      headerRow.appendChild(th);
    }
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // Body
    const tbody = document.createElement('tbody');
    const rowsToRender = Math.min(rowCount, this.maxInitialRows);

    for (let r = startRow; r < startRow + rowsToRender; r++) {
      const tr = document.createElement('tr');
      tr.dataset.r = String(r);

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
        td.dataset.r = String(r);
        td.dataset.c = String(c);

        if (cell) {
          let val = cell.w !== undefined ? String(cell.w) : (cell.v !== undefined ? String(cell.v) : '');
          val = DOMPurify.sanitize(val);
          td.textContent = val;
          td.dataset.rawVal = String(cell.v !== undefined ? cell.v : '');

          if (cell.t === 'n') {
            td.classList.add('type-number');
          }
        } else {
          td.textContent = '';
        }

        // Multi-cell selection
        td.addEventListener('mousedown', (e) => {
          if (e.button !== 0) return;
          this.isSelecting = true;
          this.selectStart = { r, c };
          this.selectRange(r, c, r, c);
        });

        td.addEventListener('mouseover', () => {
          if (this.isSelecting && this.selectStart) {
            const minR = Math.min(this.selectStart.r, r);
            const maxR = Math.max(this.selectStart.r, r);
            const minC = Math.min(this.selectStart.c, c);
            const maxC = Math.max(this.selectStart.c, c);
            this.selectRange(minR, minC, maxR, maxC);
          }
        });

        tr.appendChild(td);
      }

      tbody.appendChild(tr);
    }

    table.appendChild(tbody);
    this.tableContainer.appendChild(table);

    // Load More indicator
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

  private initColumnResizer(resizer: HTMLElement, th: HTMLElement, colIndex: number) {
    let startX = 0;
    let startWidth = 0;

    const onMouseMove = (e: MouseEvent) => {
      const diff = e.clientX - startX;
      const newWidth = Math.max(45, startWidth + diff);
      th.style.width = `${newWidth}px`;
      th.style.minWidth = `${newWidth}px`;
      th.style.maxWidth = `${newWidth}px`;

      // Update all cells in this column
      const colCells = this.tableContainer?.querySelectorAll(`td[data-c="${colIndex}"]`);
      colCells?.forEach((cell) => {
        const el = cell as HTMLElement;
        el.style.width = `${newWidth}px`;
        el.style.minWidth = `${newWidth}px`;
        el.style.maxWidth = `${newWidth}px`;
      });
    };

    const onMouseUp = () => {
      resizer.classList.remove('resizing');
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
    };

    resizer.addEventListener('mousedown', (e) => {
      e.stopPropagation();
      e.preventDefault();
      startX = e.clientX;
      startWidth = th.offsetWidth;
      resizer.classList.add('resizing');
      document.addEventListener('mousemove', onMouseMove);
      document.addEventListener('mouseup', onMouseUp);
    });
  }

  private openFilterPopup(colIndex: number, btnEl: HTMLElement) {
    this.closeFilterPopup();

    const rect = btnEl.getBoundingClientRect();
    const popup = document.createElement('div');
    popup.className = 'column-filter-popup';
    popup.style.top = `${rect.bottom + 4}px`;
    popup.style.left = `${Math.min(window.innerWidth - 240, rect.left)}px`;

    // Collect unique values in column
    const cells = this.tableContainer?.querySelectorAll(`td[data-c="${colIndex}"]`);
    const valuesSet = new Set<string>();
    cells?.forEach((c) => {
      valuesSet.add(c.textContent?.trim() || '(Boş)');
    });
    const uniqueValues = Array.from(valuesSet).sort();

    // Search inside filter
    const searchInput = document.createElement('input');
    searchInput.className = 'filter-search-input';
    searchInput.placeholder = 'Değerlerde ara...';
    popup.appendChild(searchInput);

    // List container
    const list = document.createElement('div');
    list.className = 'filter-values-list';

    const selectedSet = this.activeFilters.get(colIndex) || new Set(uniqueValues);
    const checkboxes: Map<string, HTMLInputElement> = new Map();

    const renderListItems = (filterText: string = '') => {
      list.innerHTML = '';
      uniqueValues.forEach((val) => {
        if (filterText && !val.toLowerCase().includes(filterText.toLowerCase())) return;

        const item = document.createElement('label');
        item.className = 'filter-value-item';

        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = selectedSet.has(val);
        checkboxes.set(val, cb);

        const span = document.createElement('span');
        span.textContent = val;

        item.appendChild(cb);
        item.appendChild(span);
        list.appendChild(item);
      });
    };

    renderListItems();
    searchInput.addEventListener('input', () => {
      renderListItems(searchInput.value);
    });

    popup.appendChild(list);

    // Action buttons
    const actions = document.createElement('div');
    actions.className = 'filter-actions';

    const btnClear = document.createElement('button');
    btnClear.className = 'toolbar-btn text-btn';
    btnClear.textContent = 'Sıfırla';
    btnClear.onclick = () => {
      this.activeFilters.delete(colIndex);
      btnEl.classList.remove('active');
      this.applyActiveFilters();
      this.closeFilterPopup();
    };

    const btnApply = document.createElement('button');
    btnApply.className = 'toolbar-btn text-btn';
    btnApply.textContent = 'Uygula';
    btnApply.onclick = () => {
      const newAllowed = new Set<string>();
      checkboxes.forEach((cb, val) => {
        if (cb.checked) newAllowed.add(val);
      });
      if (newAllowed.size < uniqueValues.length) {
        this.activeFilters.set(colIndex, newAllowed);
        btnEl.classList.add('active');
      } else {
        this.activeFilters.delete(colIndex);
        btnEl.classList.remove('active');
      }
      this.applyActiveFilters();
      this.closeFilterPopup();
    };

    actions.appendChild(btnClear);
    actions.appendChild(btnApply);
    popup.appendChild(actions);

    document.body.appendChild(popup);
    this.filterPopup = popup;
  }

  private closeFilterPopup() {
    if (this.filterPopup) {
      this.filterPopup.remove();
      this.filterPopup = null;
    }
  }

  private applyActiveFilters() {
    const tbody = this.tableContainer?.querySelector('tbody');
    if (!tbody) return;

    const rows = Array.from(tbody.querySelectorAll('tr'));
    let visibleCount = 0;

    rows.forEach((row) => {
      let isVisible = true;

      this.activeFilters.forEach((allowedVals, colIdx) => {
        const cell = row.querySelector(`td[data-c="${colIdx}"]`);
        const cellVal = cell?.textContent?.trim() || '(Boş)';
        if (!allowedVals.has(cellVal)) {
          isVisible = false;
        }
      });

      row.style.display = isVisible ? '' : 'none';
      if (isVisible) visibleCount++;
    });

    this.callbacks.onStatsUpdated?.(
      `Filtrelendi: ${visibleCount} / ${rows.length} satır | Sayfa: ${this.currentSheetName}`
    );
  }

  private selectRange(minR: number, minC: number, maxR: number, maxC: number) {
    if (!this.tableContainer) return;

    this.selectedCells.forEach((el) => el.classList.remove('selected'));
    this.selectedCells = [];

    const cells = this.tableContainer.querySelectorAll('.cell-content');
    const numbers: number[] = [];

    cells.forEach((cellEl) => {
      const el = cellEl as HTMLElement;
      const r = parseInt(el.dataset.r || '-1', 10);
      const c = parseInt(el.dataset.c || '-1', 10);

      if (r >= minR && r <= maxR && c >= minC && c <= maxC) {
        el.classList.add('selected');
        this.selectedCells.push(el);

        const raw = el.dataset.rawVal || el.textContent || '';
        const parsed = parseFloat(raw.replace(/[$,₺€£\s]/g, '').replace(/,/g, ''));
        if (!isNaN(parsed) && raw.trim() !== '') {
          numbers.push(parsed);
        }
      }
    });

    const totalCount = this.selectedCells.length;
    if (totalCount === 1) {
      const single = this.selectedCells[0];
      const addr = single.dataset.addr || '';
      const val = single.textContent || '';
      this.callbacks.onCellSelected?.(addr, val);
    }

    if (numbers.length > 0) {
      const sum = numbers.reduce((a, b) => a + b, 0);
      const avg = sum / numbers.length;
      const min = Math.min(...numbers);
      const max = Math.max(...numbers);

      const formatNum = (n: number) =>
        new Intl.NumberFormat('tr-TR', { maximumFractionDigits: 2 }).format(n);

      this.callbacks.onSelectionStats?.(
        `Seçim: ${totalCount} hücre | Toplam: ${formatNum(sum)} | Ort: ${formatNum(avg)} | Min: ${formatNum(min)} | Max: ${formatNum(max)}`
      );
    } else {
      this.callbacks.onSelectionStats?.(`Seçim: ${totalCount} hücre`);
    }
  }

  private sortTableByColumn(colIndex: number) {
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
      tr.dataset.r = String(r);

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
        td.dataset.r = String(r);
        td.dataset.c = String(c);

        if (cell) {
          let val = cell.w !== undefined ? String(cell.w) : (cell.v !== undefined ? String(cell.v) : '');
          val = DOMPurify.sanitize(val);
          td.textContent = val;
          td.dataset.rawVal = String(cell.v !== undefined ? cell.v : '');
          if (cell.t === 'n') {
            td.classList.add('type-number');
          }
        } else {
          td.textContent = '';
        }

        td.addEventListener('mousedown', (e) => {
          if (e.button !== 0) return;
          this.isSelecting = true;
          this.selectStart = { r, c };
          this.selectRange(r, c, r, c);
        });

        td.addEventListener('mouseover', () => {
          if (this.isSelecting && this.selectStart) {
            const minR = Math.min(this.selectStart.r, r);
            const maxR = Math.max(this.selectStart.r, r);
            const minC = Math.min(this.selectStart.c, c);
            const maxC = Math.max(this.selectStart.c, c);
            this.selectRange(minR, minC, maxR, maxC);
          }
        });

        tr.appendChild(td);
      }

      tbody.appendChild(tr);
    }
  }

  public exportAsMarkdown(): string {
    if (!this.workbook || !this.currentSheetName) return '';
    const sheet = this.workbook.Sheets[this.currentSheetName];
    if (!sheet || !sheet['!ref']) return '';

    let startRow: number, endRow: number, startCol: number, endCol: number;

    if (this.selectedCells.length > 1) {
      const rows = this.selectedCells.map((el) => parseInt(el.dataset.r || '0', 10));
      const cols = this.selectedCells.map((el) => parseInt(el.dataset.c || '0', 10));
      startRow = Math.min(...rows);
      endRow = Math.max(...rows);
      startCol = Math.min(...cols);
      endCol = Math.max(...cols);
    } else {
      const range = XLSX.utils.decode_range(sheet['!ref']);
      startRow = range.s.r;
      endRow = Math.min(range.e.r, startRow + 500);
      startCol = range.s.c;
      endCol = range.e.c;
    }

    const rowsLines: string[] = [];
    const headers: string[] = [];
    const separators: string[] = [];

    for (let c = startCol; c <= endCol; c++) {
      const cellAddress = XLSX.utils.encode_cell({ r: startRow, c });
      const cell = sheet[cellAddress];
      const headerText = cell ? String(cell.v).replace(/\|/g, '\\|').trim() : `Kolon ${XLSX.utils.encode_col(c)}`;
      headers.push(headerText || `Kolon ${XLSX.utils.encode_col(c)}`);
      separators.push('---');
    }
    rowsLines.push(`| ${headers.join(' | ')} |`);
    rowsLines.push(`| ${separators.join(' | ')} |`);

    for (let r = startRow + 1; r <= endRow; r++) {
      const rowCells: string[] = [];
      for (let c = startCol; c <= endCol; c++) {
        const cellAddress = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[cellAddress];
        const val = cell ? String(cell.v).replace(/\|/g, '\\|').trim() : '';
        rowCells.push(val);
      }
      rowsLines.push(`| ${rowCells.join(' | ')} |`);
    }

    return rowsLines.join('\n');
  }

  public exportAsJson(): string {
    if (!this.workbook || !this.currentSheetName) return '[]';
    const sheet = this.workbook.Sheets[this.currentSheetName];
    if (!sheet || !sheet['!ref']) return '[]';

    const range = XLSX.utils.decode_range(sheet['!ref']);
    const jsonData = XLSX.utils.sheet_to_json(sheet, {
      range: range,
      defval: null,
    });

    return JSON.stringify(jsonData, null, 2);
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
