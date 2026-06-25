import { invoke, isTauri } from '@tauri-apps/api/core';

export interface SpreadsheetSheetSummary {
  name: string;
  total_rows: number;
  total_cols: number;
}

export interface SpreadsheetLoadResult {
  file_name: string;
  active_sheet_index: number;
  sheet_name: string;
  total_rows: number;
  total_cols: number;
  sheets: SpreadsheetSheetSummary[];
}

export interface SpreadsheetChunkResponse {
  start_row: number;
  row_limit: number;
  start_col: number;
  col_limit: number;
  total_rows: number;
  total_cols: number;
  rows: string[][];
}

export interface SpreadsheetSearchMatch {
  row: number;
  col: number;
  value: string;
}

export interface SpreadsheetVerticalWindow {
  visibleStartRow: number;
  visibleEndRow: number;
  requestStartRow: number;
  requestLimit: number;
  topSpacerHeight: number;
  bottomSpacerHeight: number;
  visibleRowCount: number;
}

export interface SpreadsheetHorizontalWindow {
  visibleStartCol: number;
  visibleEndCol: number;
  requestStartCol: number;
  requestLimit: number;
  leftSpacerWidth: number;
  rightSpacerWidth: number;
  visibleColCount: number;
}

function clamp(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), max);
}

function excelColumnLabel(index: number) {
  let value = index;
  let label = '';

  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }

  return label;
}

function canInvokeTauriCommand() {
  if (typeof window === 'undefined') {
    return false;
  }

  const tauriWindow = window as Window & {
    __TAURI_INTERNALS__?: {
      invoke?: (cmd: string, args?: unknown, options?: unknown) => Promise<unknown>;
    };
  };

  return isTauri() && typeof tauriWindow.__TAURI_INTERNALS__?.invoke === 'function';
}

async function invokeTauriCommand<T>(command: string, args?: Record<string, unknown>) {
  if (!canInvokeTauriCommand()) {
    throw new Error(
      'Excel loading requires the Tauri desktop runtime. Open the app with `python run.py` or a Tauri bundle.'
    );
  }

  return invoke<T>(command, args);
}

function computeVerticalWindow(
  scrollTop: number,
  viewportHeight: number,
  totalRows: number,
  rowHeight: number,
  overscanRows: number,
  rowChunkSize: number
): SpreadsheetVerticalWindow {
  if (totalRows <= 0 || rowHeight <= 0) {
    return {
      visibleStartRow: 0,
      visibleEndRow: 0,
      requestStartRow: 0,
      requestLimit: 0,
      topSpacerHeight: 0,
      bottomSpacerHeight: 0,
      visibleRowCount: 0
    };
  }

  const visibleRowCount = Math.max(1, Math.ceil(viewportHeight / rowHeight));
  const firstVisibleRow = clamp(Math.floor(scrollTop / rowHeight), 0, Math.max(totalRows - 1, 0));
  const visibleEndRow = clamp(firstVisibleRow + visibleRowCount, 0, totalRows);
  const bufferedStartRow = clamp(firstVisibleRow - overscanRows, 0, totalRows);
  const bufferedEndRow = clamp(visibleEndRow + overscanRows, bufferedStartRow, totalRows);
  const requestStartRow = Math.floor(bufferedStartRow / rowChunkSize) * rowChunkSize;
  const requestEndRow = Math.ceil(bufferedEndRow / rowChunkSize) * rowChunkSize;
  const requestLimit = clamp(requestEndRow - requestStartRow, 0, totalRows - requestStartRow);

  return {
    visibleStartRow: firstVisibleRow,
    visibleEndRow,
    requestStartRow,
    requestLimit,
    topSpacerHeight: requestStartRow * rowHeight,
    bottomSpacerHeight: Math.max(0, (totalRows - (requestStartRow + requestLimit)) * rowHeight),
    visibleRowCount
  };
}

function computeHorizontalWindow(
  scrollLeft: number,
  viewportWidth: number,
  rowHeaderWidth: number,
  totalCols: number,
  columnWidth: number,
  overscanCols: number,
  columnChunkSize: number
): SpreadsheetHorizontalWindow {
  if (totalCols <= 0 || columnWidth <= 0) {
    return {
      visibleStartCol: 0,
      visibleEndCol: 0,
      requestStartCol: 0,
      requestLimit: 0,
      leftSpacerWidth: 0,
      rightSpacerWidth: 0,
      visibleColCount: 0
    };
  }

  const usableViewportWidth = Math.max(columnWidth, viewportWidth - rowHeaderWidth);
  const visibleColCount = Math.max(1, Math.ceil(usableViewportWidth / columnWidth));
  const firstVisibleCol = clamp(Math.floor(scrollLeft / columnWidth), 0, Math.max(totalCols - 1, 0));
  const visibleEndCol = clamp(firstVisibleCol + visibleColCount, 0, totalCols);
  const bufferedStartCol = clamp(firstVisibleCol - overscanCols, 0, totalCols);
  const bufferedEndCol = clamp(visibleEndCol + overscanCols, bufferedStartCol, totalCols);
  const requestStartCol = Math.floor(bufferedStartCol / columnChunkSize) * columnChunkSize;
  const requestEndCol = Math.ceil(bufferedEndCol / columnChunkSize) * columnChunkSize;
  const requestLimit = clamp(requestEndCol - requestStartCol, 0, totalCols - requestStartCol);

  return {
    visibleStartCol: firstVisibleCol,
    visibleEndCol,
    requestStartCol,
    requestLimit,
    leftSpacerWidth: requestStartCol * columnWidth,
    rightSpacerWidth: Math.max(0, (totalCols - (requestStartCol + requestLimit)) * columnWidth),
    visibleColCount
  };
}

class SpreadsheetState {
  filePath = $state('');
  fileName = $state('');
  sheetName = $state('');
  sheets = $state<SpreadsheetSheetSummary[]>([]);
  activeSheetIndex = $state(0);
  totalRows = $state(0);
  totalCols = $state(0);
  scrollTop = $state(0);
  scrollLeft = $state(0);
  viewportHeight = $state(0);
  viewportWidth = $state(0);
  rowHeight = $state(28);
  columnWidth = $state(144);
  rowHeaderWidth = $state(72);
  rowChunkSize = $state(128);
  columnChunkSize = $state(8);
  overscanRows = $state(24);
  overscanCols = $state(3);
  isLoading = $state(false);
  isSearching = $state(false);
  loadError = $state<string | null>(null);
  selectedRow = $state<number | null>(null);
  selectedCol = $state<number | null>(null);
  searchQuery = $state('');
  searchResults = $state<SpreadsheetSearchMatch[]>([]);
  activePacket = $state<SpreadsheetChunkResponse | null>(null);
  lastLoadedKey = $state('');

  private syncScheduled = false;
  private requestToken = 0;
  private readonly chunkCache = new Map<string, SpreadsheetChunkResponse>();
  private readonly maxCacheEntries = 8;

  activeRows = $derived(this.activePacket?.rows ?? []);
  totalContentHeight = $derived(this.totalRows * this.rowHeight);
  totalContentWidth = $derived(this.totalCols * this.columnWidth);
  verticalWindow = $derived.by(() =>
    computeVerticalWindow(
      this.scrollTop,
      this.viewportHeight,
      this.totalRows,
      this.rowHeight,
      this.overscanRows,
      this.rowChunkSize
    )
  );
  horizontalWindow = $derived.by(() =>
    computeHorizontalWindow(
      this.scrollLeft,
      this.viewportWidth,
      this.rowHeaderWidth,
      this.totalCols,
      this.columnWidth,
      this.overscanCols,
      this.columnChunkSize
    )
  );
  hasWorkbook = $derived(this.fileName.length > 0 && this.sheets.length > 0);
  visibleColumnLabels = $derived.by(() =>
    Array.from({ length: this.activePacket?.col_limit ?? 0 }, (_, offset) =>
      excelColumnLabel((this.activePacket?.start_col ?? 0) + offset + 1)
    )
  );
  statusLabel = $derived(
    this.isLoading ? 'Loading' : this.loadError ? 'Error' : this.hasWorkbook ? 'Ready' : 'Idle'
  );

  setViewportMetrics(viewportHeight: number, viewportWidth: number) {
    this.viewportHeight = Math.max(0, Math.floor(viewportHeight));
    this.viewportWidth = Math.max(0, Math.floor(viewportWidth));
    void this.scheduleViewportSync();
  }

  setScrollOffsets(scrollTop: number, scrollLeft: number) {
    const nextTop = Math.max(0, Math.floor(scrollTop));
    const nextLeft = Math.max(0, Math.floor(scrollLeft));

    if (nextTop === this.scrollTop && nextLeft === this.scrollLeft) {
      return;
    }

    this.scrollTop = nextTop;
    this.scrollLeft = nextLeft;
    void this.scheduleViewportSync();
  }

  setSelectedCell(rowIndex: number | null, colIndex: number | null) {
    this.selectedRow = rowIndex;
    this.selectedCol = colIndex;
  }

  resetViewport() {
    this.scrollTop = 0;
    this.scrollLeft = 0;
    this.selectedRow = null;
    this.selectedCol = null;
    void this.scheduleViewportSync();
  }

  async loadExcelFile(path: string) {
    const normalizedPath = path.trim();

    if (!normalizedPath) {
      this.loadError = 'Enter a valid Excel file path.';
      return null;
    }

    this.isLoading = true;
    this.loadError = null;

    try {
      const result = await invokeTauriCommand<SpreadsheetLoadResult>('load_excel_file', {
        path: normalizedPath
      });

      this.filePath = normalizedPath;
      this.applyWorkbookResult(result);
      this.resetWorkbookViewportState();
      await this.syncViewportChunk(true);
      return result;
    } catch (error) {
      this.clearWorkbookState(normalizedPath);
      this.loadError = error instanceof Error ? error.message : 'Excel file could not be loaded.';
      return null;
    } finally {
      this.isLoading = false;
    }
  }

  async activateSheet(sheetIndex: number) {
    if (!this.hasWorkbook || sheetIndex === this.activeSheetIndex) {
      return null;
    }

    this.isLoading = true;
    this.loadError = null;

    try {
      const result = await invokeTauriCommand<SpreadsheetLoadResult>('set_active_sheet', {
        sheetIndex
      });

      this.applyWorkbookResult(result);
      this.resetWorkbookViewportState();
      await this.syncViewportChunk(true);
      return result;
    } catch (error) {
      this.loadError = error instanceof Error ? error.message : 'Worksheet could not be activated.';
      return null;
    } finally {
      this.isLoading = false;
    }
  }

  async reloadVisibleChunk() {
    await this.syncViewportChunk(true);
  }

  async searchWorkbook(query: string) {
    this.searchQuery = query;
    const normalizedQuery = query.trim();

    if (!this.hasWorkbook || normalizedQuery.length === 0) {
      this.searchResults = [];
      return [];
    }

    this.isSearching = true;
    this.loadError = null;

    try {
      const results = await invokeTauriCommand<SpreadsheetSearchMatch[]>('find_in_active_sheet', {
        query: normalizedQuery,
        limit: 40
      });

      this.searchResults = results;
      return results;
    } catch (error) {
      this.loadError = error instanceof Error ? error.message : 'Workbook search could not be completed.';
      return [];
    } finally {
      this.isSearching = false;
    }
  }

  focusCell(rowIndex: number, colIndex: number) {
    const clampedRow = clamp(rowIndex, 0, Math.max(this.totalRows - 1, 0));
    const clampedCol = clamp(colIndex, 0, Math.max(this.totalCols - 1, 0));

    this.selectedRow = clampedRow;
    this.selectedCol = clampedCol;
    this.scrollTop = clampedRow * this.rowHeight;
    this.scrollLeft = clampedCol * this.columnWidth;
    void this.scheduleViewportSync();
  }

  goToRow(rowNumber: number) {
    const targetRow = clamp(Math.floor(rowNumber) - 1, 0, Math.max(this.totalRows - 1, 0));
    this.selectedRow = targetRow;
    this.scrollTop = targetRow * this.rowHeight;
    void this.scheduleViewportSync();
    return targetRow;
  }

  private applyWorkbookResult(result: SpreadsheetLoadResult) {
    this.fileName = result.file_name;
    this.sheetName = result.sheet_name;
    this.sheets = result.sheets;
    this.activeSheetIndex = result.active_sheet_index;
    this.totalRows = result.total_rows;
    this.totalCols = result.total_cols;
  }

  private resetWorkbookViewportState() {
    this.scrollTop = 0;
    this.scrollLeft = 0;
    this.selectedRow = null;
    this.selectedCol = null;
    this.searchResults = [];
    this.activePacket = null;
    this.lastLoadedKey = '';
    this.chunkCache.clear();
  }

  private clearWorkbookState(filePath = '') {
    this.filePath = filePath;
    this.fileName = '';
    this.sheetName = '';
    this.sheets = [];
    this.activeSheetIndex = 0;
    this.totalRows = 0;
    this.totalCols = 0;
    this.scrollTop = 0;
    this.scrollLeft = 0;
    this.selectedRow = null;
    this.selectedCol = null;
    this.searchResults = [];
    this.activePacket = null;
    this.lastLoadedKey = '';
    this.chunkCache.clear();
  }

  private async scheduleViewportSync() {
    if (typeof window === 'undefined') {
      await this.syncViewportChunk(false);
      return;
    }

    if (this.syncScheduled) {
      return;
    }

    this.syncScheduled = true;

    window.requestAnimationFrame(() => {
      this.syncScheduled = false;
      void this.syncViewportChunk(false);
    });
  }

  private buildCacheKey(startRow: number, rowLimit: number, startCol: number, colLimit: number) {
    return `${startRow}:${rowLimit}:${startCol}:${colLimit}`;
  }

  private rememberChunk(chunk: SpreadsheetChunkResponse) {
    const key = this.buildCacheKey(
      chunk.start_row,
      chunk.row_limit,
      chunk.start_col,
      chunk.col_limit
    );

    this.chunkCache.set(key, chunk);
    this.lastLoadedKey = key;

    if (this.chunkCache.size <= this.maxCacheEntries) {
      return;
    }

    const oldestKey = this.chunkCache.keys().next().value as string | undefined;
    if (oldestKey) {
      this.chunkCache.delete(oldestKey);
    }
  }

  private getCachedChunk(startRow: number, rowLimit: number, startCol: number, colLimit: number) {
    const key = this.buildCacheKey(startRow, rowLimit, startCol, colLimit);
    const cached = this.chunkCache.get(key);

    if (!cached) {
      return null;
    }

    this.lastLoadedKey = key;
    return cached;
  }

  private async syncViewportChunk(forceFetch: boolean) {
    if (!this.hasWorkbook || this.totalRows <= 0 || this.totalCols <= 0) {
      this.activePacket = null;
      return;
    }

    const verticalWindow = this.verticalWindow;
    const horizontalWindow = this.horizontalWindow;

    if (verticalWindow.requestLimit <= 0 || horizontalWindow.requestLimit <= 0) {
      this.activePacket = null;
      return;
    }

    const cachedChunk = this.getCachedChunk(
      verticalWindow.requestStartRow,
      verticalWindow.requestLimit,
      horizontalWindow.requestStartCol,
      horizontalWindow.requestLimit
    );

    if (cachedChunk && !forceFetch) {
      this.activePacket = cachedChunk;
      this.loadError = null;
      return;
    }

    const token = ++this.requestToken;
    this.isLoading = true;

    try {
      const chunk = await invokeTauriCommand<SpreadsheetChunkResponse>('get_data_chunk', {
        startRow: verticalWindow.requestStartRow,
        rowLimit: verticalWindow.requestLimit,
        startCol: horizontalWindow.requestStartCol,
        colLimit: horizontalWindow.requestLimit
      });

      if (token !== this.requestToken) {
        return;
      }

      this.rememberChunk(chunk);
      this.activePacket = chunk;
      this.totalRows = chunk.total_rows;
      this.totalCols = chunk.total_cols;
      this.loadError = null;
    } catch (error) {
      if (token !== this.requestToken) {
        return;
      }

      this.loadError = error instanceof Error ? error.message : 'Viewport data could not be loaded.';
    } finally {
      if (token === this.requestToken) {
        this.isLoading = false;
      }
    }
  }
}

export const spreadsheetState = new SpreadsheetState();
