import { invoke, isTauri } from '@tauri-apps/api/core';

export type SpreadsheetSortDirection = 'asc' | 'desc';

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
  source_total_rows: number;
  total_cols: number;
  header_row_enabled: boolean;
  filter_query: string;
  sort_col: number | null;
  sort_direction: SpreadsheetSortDirection | null;
  sheets: SpreadsheetSheetSummary[];
}

export interface SpreadsheetChunkResponse {
  start_row: number;
  row_limit: number;
  start_col: number;
  col_limit: number;
  total_rows: number;
  source_total_rows: number;
  total_cols: number;
  header_row_enabled: boolean;
  headers: string[];
  source_rows: number[];
  rows: string[][];
}

export interface SpreadsheetSearchMatch {
  display_row: number;
  source_row: number;
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
  sourceTotalRows = $state(0);
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
  isApplyingView = $state(false);
  loadError = $state<string | null>(null);
  selectedRow = $state<number | null>(null);
  selectedSourceRow = $state<number | null>(null);
  selectedCol = $state<number | null>(null);
  searchQuery = $state('');
  filterQuery = $state('');
  headerRowEnabled = $state(false);
  sortCol = $state<number | null>(null);
  sortDirection = $state<SpreadsheetSortDirection | null>(null);
  searchResults = $state<SpreadsheetSearchMatch[]>([]);
  activePacket = $state<SpreadsheetChunkResponse | null>(null);
  lastLoadedKey = $state('');

  private syncScheduled = false;
  private requestToken = 0;
  private prefetchTimer = 0;
  private readonly chunkCache = new Map<string, SpreadsheetChunkResponse>();
  private readonly inflightChunkRequests = new Map<string, Promise<SpreadsheetChunkResponse>>();
  private readonly maxCacheEntries = 8;

  activeRows = $derived(this.activePacket?.rows ?? []);
  activeSourceRows = $derived(this.activePacket?.source_rows ?? []);
  selectedCellValue = $derived.by(() => {
    if (
      !this.activePacket ||
      this.selectedRow === null ||
      this.selectedCol === null ||
      this.selectedRow < this.activePacket.start_row ||
      this.selectedRow >= this.activePacket.start_row + this.activePacket.row_limit ||
      this.selectedCol < this.activePacket.start_col ||
      this.selectedCol >= this.activePacket.start_col + this.activePacket.col_limit
    ) {
      return '';
    }

    const rowOffset = this.selectedRow - this.activePacket.start_row;
    const colOffset = this.selectedCol - this.activePacket.start_col;
    return this.activePacket.rows[rowOffset]?.[colOffset] ?? '';
  });
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
    Array.from({ length: this.activePacket?.col_limit ?? 0 }, (_, offset) => {
      const fallbackLabel = excelColumnLabel((this.activePacket?.start_col ?? 0) + offset + 1);
      const headerValue = this.activePacket?.headers?.[offset]?.trim() ?? '';
      return headerValue.length > 0 ? headerValue : fallbackLabel;
    })
  );
  activeFilter = $derived(this.filterQuery.trim());
  statusLabel = $derived.by(() => {
    if (this.isLoading) {
      return 'Loading';
    }

    if (this.isApplyingView) {
      return 'Applying View';
    }

    if (this.loadError) {
      return 'Error';
    }

    return this.hasWorkbook ? 'Ready' : 'Idle';
  });

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

  setSelectedCell(
    displayRow: number | null,
    colIndex: number | null,
    sourceRow: number | null = null
  ) {
    this.selectedRow = displayRow;
    this.selectedCol = colIndex;
    this.selectedSourceRow = sourceRow;
  }

  resetViewport() {
    this.scrollTop = 0;
    this.scrollLeft = 0;
    this.selectedRow = null;
    this.selectedSourceRow = null;
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

  async applyFilter(query: string) {
    return this.configureView({ filterQuery: query });
  }

  async setHeaderRowEnabled(enabled: boolean) {
    return this.configureView({ headerRowEnabled: enabled });
  }

  async cycleColumnSort(colIndex: number) {
    if (this.sortCol !== colIndex) {
      return this.configureView({
        sortCol: colIndex,
        sortDirection: 'asc'
      });
    }

    if (this.sortDirection === 'asc') {
      return this.configureView({
        sortCol: colIndex,
        sortDirection: 'desc'
      });
    }

    if (this.sortDirection === 'desc') {
      return this.configureView({
        sortCol: null,
        sortDirection: null
      });
    }

    return this.configureView({
      sortCol: colIndex,
      sortDirection: 'asc'
    });
  }

  async setColumnSort(colIndex: number, direction: SpreadsheetSortDirection) {
    return this.configureView({
      sortCol: colIndex,
      sortDirection: direction
    });
  }

  async clearColumnSort() {
    return this.configureView({
      sortCol: null,
      sortDirection: null
    });
  }

  async clearViewOptions() {
    return this.configureView({
      filterQuery: '',
      sortCol: null,
      sortDirection: null
    });
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
      this.loadError =
        error instanceof Error ? error.message : 'Workbook search could not be completed.';
      return [];
    } finally {
      this.isSearching = false;
    }
  }

  focusCell(displayRow: number, colIndex: number, sourceRow: number | null = null) {
    const clampedRow = clamp(displayRow, 0, Math.max(this.totalRows - 1, 0));
    const clampedCol = clamp(colIndex, 0, Math.max(this.totalCols - 1, 0));

    this.selectedRow = clampedRow;
    this.selectedSourceRow = sourceRow;
    this.selectedCol = clampedCol;
    this.scrollTop = clampedRow * this.rowHeight;
    this.scrollLeft = clampedCol * this.columnWidth;
    void this.scheduleViewportSync();
  }

  moveSelection(rowDelta: number, colDelta: number) {
    if (!this.hasWorkbook) {
      return;
    }

    const baseRow = this.selectedRow ?? 0;
    const baseCol = this.selectedCol ?? 0;
    const nextRow = clamp(baseRow + rowDelta, 0, Math.max(this.totalRows - 1, 0));
    const nextCol = clamp(baseCol + colDelta, 0, Math.max(this.totalCols - 1, 0));

    this.selectedRow = nextRow;
    this.selectedCol = nextCol;
    this.selectedSourceRow = this.resolveSourceRow(nextRow);
    this.ensureSelectionVisible();
  }

  moveSelectionToEdge(kind: 'row-start' | 'row-end' | 'grid-start' | 'grid-end') {
    if (!this.hasWorkbook) {
      return;
    }

    const currentRow = this.selectedRow ?? 0;
    const currentCol = this.selectedCol ?? 0;
    let nextRow = currentRow;
    let nextCol = currentCol;

    if (kind === 'row-start') {
      nextCol = 0;
    } else if (kind === 'row-end') {
      nextCol = Math.max(this.totalCols - 1, 0);
    } else if (kind === 'grid-start') {
      nextRow = 0;
      nextCol = 0;
    } else {
      nextRow = Math.max(this.totalRows - 1, 0);
      nextCol = Math.max(this.totalCols - 1, 0);
    }

    this.selectedRow = nextRow;
    this.selectedCol = nextCol;
    this.selectedSourceRow = this.resolveSourceRow(nextRow);
    this.ensureSelectionVisible();
  }

  pageSelection(direction: 'up' | 'down') {
    if (!this.hasWorkbook) {
      return;
    }

    const pageSize = Math.max(1, this.verticalWindow.visibleRowCount - 1);
    this.moveSelection(direction === 'down' ? pageSize : -pageSize, 0);
  }

  goToRow(rowNumber: number) {
    const targetDisplayRow = clamp(Math.floor(rowNumber) - 1, 0, Math.max(this.totalRows - 1, 0));
    this.selectedRow = targetDisplayRow;
    this.selectedSourceRow = this.resolveSourceRow(targetDisplayRow);
    this.scrollTop = targetDisplayRow * this.rowHeight;
    void this.scheduleViewportSync();
    return targetDisplayRow;
  }

  private ensureSelectionVisible() {
    if (this.selectedRow === null || this.selectedCol === null) {
      return;
    }

    const rowTop = this.selectedRow * this.rowHeight;
    const rowBottom = rowTop + this.rowHeight;
    const colLeft = this.selectedCol * this.columnWidth;
    const colRight = colLeft + this.columnWidth;
    const viewportLeft = this.scrollLeft;
    const viewportRight = this.scrollLeft + Math.max(0, this.viewportWidth - this.rowHeaderWidth);
    const viewportTop = this.scrollTop;
    const viewportBottom = this.scrollTop + this.viewportHeight;

    if (rowTop < viewportTop) {
      this.scrollTop = rowTop;
    } else if (rowBottom > viewportBottom) {
      this.scrollTop = Math.max(0, rowBottom - this.viewportHeight);
    }

    if (colLeft < viewportLeft) {
      this.scrollLeft = colLeft;
    } else if (colRight > viewportRight) {
      this.scrollLeft = Math.max(0, colRight - Math.max(0, this.viewportWidth - this.rowHeaderWidth));
    }

    void this.scheduleViewportSync();
  }

  resolveSourceRow(displayRow: number) {
    const activePacket = this.activePacket;

    if (
      activePacket &&
      displayRow >= activePacket.start_row &&
      displayRow < activePacket.start_row + activePacket.row_limit
    ) {
      return activePacket.source_rows[displayRow - activePacket.start_row] ?? null;
    }

    if (this.sortCol === null && this.activeFilter.length === 0) {
      const sourceRow = displayRow + (this.headerRowEnabled && this.sourceTotalRows > 0 ? 1 : 0);
      return sourceRow < this.sourceTotalRows ? sourceRow : null;
    }

    return null;
  }

  private applyWorkbookResult(result: SpreadsheetLoadResult) {
    this.fileName = result.file_name;
    this.sheetName = result.sheet_name;
    this.sheets = result.sheets;
    this.activeSheetIndex = result.active_sheet_index;
    this.totalRows = result.total_rows;
    this.sourceTotalRows = result.source_total_rows;
    this.totalCols = result.total_cols;
    this.headerRowEnabled = result.header_row_enabled;
    this.filterQuery = result.filter_query;
    this.sortCol = result.sort_col;
    this.sortDirection = result.sort_direction;
  }

  private resetWorkbookViewportState() {
    this.scrollTop = 0;
    this.scrollLeft = 0;
    this.selectedRow = null;
    this.selectedSourceRow = null;
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
    this.sourceTotalRows = 0;
    this.totalCols = 0;
    this.scrollTop = 0;
    this.scrollLeft = 0;
    this.selectedRow = null;
    this.selectedSourceRow = null;
    this.selectedCol = null;
    this.searchQuery = '';
    this.filterQuery = '';
    this.headerRowEnabled = false;
    this.sortCol = null;
    this.sortDirection = null;
    this.searchResults = [];
    this.activePacket = null;
    this.lastLoadedKey = '';
    this.chunkCache.clear();
  }

  private async configureView(config: {
    headerRowEnabled?: boolean;
    filterQuery?: string;
    sortCol?: number | null;
    sortDirection?: SpreadsheetSortDirection | null;
  }) {
    if (!this.hasWorkbook) {
      return null;
    }

    this.isApplyingView = true;
    this.loadError = null;

    const nextHeaderRowEnabled = config.headerRowEnabled ?? this.headerRowEnabled;
    const nextFilterQuery = config.filterQuery ?? this.filterQuery;
    const nextSortCol =
      config.sortCol === undefined ? this.sortCol : (config.sortCol ?? null);
    const nextSortDirection =
      config.sortDirection === undefined ? this.sortDirection : (config.sortDirection ?? null);

    try {
      const result = await invokeTauriCommand<SpreadsheetLoadResult>('configure_active_sheet_view', {
        headerRowEnabled: nextHeaderRowEnabled,
        filterQuery: nextFilterQuery.trim().length > 0 ? nextFilterQuery : null,
        sortCol: nextSortCol,
        sortDirection: nextSortCol === null ? null : nextSortDirection
      });

      this.applyWorkbookResult(result);
      this.resetWorkbookViewportState();
      await this.syncViewportChunk(true);
      return result;
    } catch (error) {
      this.loadError =
        error instanceof Error ? error.message : 'Worksheet view could not be updated.';
      return null;
    } finally {
      this.isApplyingView = false;
    }
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
    return `${this.activeSheetIndex}:${this.headerRowEnabled ? 1 : 0}:${this.filterQuery}:${this.sortCol ?? 'none'}:${
      this.sortDirection ?? 'none'
    }:${startRow}:${rowLimit}:${startCol}:${colLimit}`;
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

  private queueChunkPrefetch(
    startRow: number,
    rowLimit: number,
    startCol: number,
    colLimit: number
  ) {
    if (
      !this.hasWorkbook ||
      rowLimit <= 0 ||
      colLimit <= 0 ||
      startRow < 0 ||
      startCol < 0 ||
      startRow >= this.totalRows ||
      startCol >= this.totalCols
    ) {
      return;
    }

    const boundedRowLimit = Math.min(rowLimit, this.totalRows - startRow);
    const boundedColLimit = Math.min(colLimit, this.totalCols - startCol);
    if (boundedRowLimit <= 0 || boundedColLimit <= 0) {
      return;
    }

    const key = this.buildCacheKey(startRow, boundedRowLimit, startCol, boundedColLimit);
    if (this.chunkCache.has(key) || this.inflightChunkRequests.has(key)) {
      return;
    }

    const request = invokeTauriCommand<SpreadsheetChunkResponse>('get_data_chunk', {
      startRow,
      rowLimit: boundedRowLimit,
      startCol,
      colLimit: boundedColLimit
    })
      .then((chunk) => {
        this.rememberChunk(chunk);
        return chunk;
      })
      .finally(() => {
        this.inflightChunkRequests.delete(key);
      });

    this.inflightChunkRequests.set(key, request);
  }

  private scheduleChunkPrefetch(chunk: SpreadsheetChunkResponse) {
    if (typeof window === 'undefined') {
      return;
    }

    if (this.prefetchTimer) {
      window.clearTimeout(this.prefetchTimer);
    }

    this.prefetchTimer = window.setTimeout(() => {
      this.prefetchTimer = 0;
      const nextRowStart = chunk.start_row + chunk.row_limit;
      const prevRowStart = Math.max(0, chunk.start_row - chunk.row_limit);
      const nextColStart = chunk.start_col + chunk.col_limit;
      const prevColStart = Math.max(0, chunk.start_col - chunk.col_limit);

      this.queueChunkPrefetch(nextRowStart, chunk.row_limit, chunk.start_col, chunk.col_limit);
      this.queueChunkPrefetch(prevRowStart, chunk.row_limit, chunk.start_col, chunk.col_limit);
      this.queueChunkPrefetch(chunk.start_row, chunk.row_limit, nextColStart, chunk.col_limit);
      this.queueChunkPrefetch(chunk.start_row, chunk.row_limit, prevColStart, chunk.col_limit);
    }, 40);
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
      this.sourceTotalRows = chunk.source_total_rows;
      this.totalCols = chunk.total_cols;
      this.headerRowEnabled = chunk.header_row_enabled;
      this.loadError = null;
      this.scheduleChunkPrefetch(chunk);
    } catch (error) {
      if (token !== this.requestToken) {
        return;
      }

      this.loadError =
        error instanceof Error ? error.message : 'Viewport data could not be loaded.';
    } finally {
      if (token === this.requestToken) {
        this.isLoading = false;
      }
    }
  }
}

export const spreadsheetState = new SpreadsheetState();
