import { invoke, isTauri } from '@tauri-apps/api/core';

export interface SpreadsheetLoadResult {
  file_name: string;
  sheet_name: string;
  total_rows: number;
  total_cols: number;
}

export interface SpreadsheetChunkResponse {
  start_row: number;
  limit: number;
  total_rows: number;
  total_cols: number;
  rows: string[][];
}

export interface SpreadsheetVirtualWindow {
  visibleStartRow: number;
  visibleEndRow: number;
  requestStartRow: number;
  requestLimit: number;
  topSpacerHeight: number;
  bottomSpacerHeight: number;
  visibleRowCount: number;
  overscanRows: number;
}

function clamp(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), max);
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
    throw new Error('Excel loading requires the Tauri desktop runtime. Open the app with `python run.py` or a Tauri bundle.');
  }

  return invoke<T>(command, args);
}

function computeVirtualWindow(
  scrollTop: number,
  viewportHeight: number,
  totalRows: number,
  rowHeight: number,
  overscanRows: number,
  chunkSize: number
): SpreadsheetVirtualWindow {
  if (totalRows <= 0 || rowHeight <= 0) {
    return {
      visibleStartRow: 0,
      visibleEndRow: 0,
      requestStartRow: 0,
      requestLimit: 0,
      topSpacerHeight: 0,
      bottomSpacerHeight: 0,
      visibleRowCount: 0,
      overscanRows
    };
  }

  const visibleRowCount = Math.max(1, Math.ceil(viewportHeight / rowHeight));
  const rawFirstVisibleRow = Math.floor(scrollTop / rowHeight);
  const visibleStartRow = clamp(rawFirstVisibleRow, 0, Math.max(totalRows - 1, 0));
  const visibleEndRow = clamp(visibleStartRow + visibleRowCount, 0, totalRows);
  const bufferedStartRow = clamp(visibleStartRow - overscanRows, 0, totalRows);
  const bufferedEndRow = clamp(visibleEndRow + overscanRows, bufferedStartRow, totalRows);
  const requestStartRow = Math.floor(bufferedStartRow / chunkSize) * chunkSize;
  const requestEndRow = Math.ceil(bufferedEndRow / chunkSize) * chunkSize;
  const requestLimit = clamp(requestEndRow - requestStartRow, 0, totalRows - requestStartRow);

  return {
    visibleStartRow,
    visibleEndRow,
    requestStartRow,
    requestLimit,
    topSpacerHeight: requestStartRow * rowHeight,
    bottomSpacerHeight: Math.max(0, (totalRows - (requestStartRow + requestLimit)) * rowHeight),
    visibleRowCount,
    overscanRows
  };
}

class SpreadsheetState {
  filePath = $state('');
  fileName = $state('');
  sheetName = $state('');
  totalRows = $state(0);
  totalCols = $state(0);
  scrollTop = $state(0);
  viewportHeight = $state(0);
  viewportWidth = $state(0);
  rowHeight = $state(28);
  chunkSize = $state(128);
  overscanRows = $state(24);
  isLoading = $state(false);
  loadError = $state<string | null>(null);
  selectedRow = $state<number | null>(null);
  activePacket = $state<SpreadsheetChunkResponse | null>(null);
  lastLoadedKey = $state('');

  private syncScheduled = false;
  private requestToken = 0;
  private readonly chunkCache = new Map<string, SpreadsheetChunkResponse>();
  private readonly maxCacheEntries = 6;

  activeRows = $derived(this.activePacket?.rows ?? []);
  totalContentHeight = $derived(this.totalRows * this.rowHeight);
  visibleRowCount = $derived(
    Math.max(1, this.viewportHeight > 0 ? Math.ceil(this.viewportHeight / this.rowHeight) : 1)
  );
  virtualWindow = $derived.by(() =>
    computeVirtualWindow(
      this.scrollTop,
      this.viewportHeight,
      this.totalRows,
      this.rowHeight,
      this.overscanRows,
      this.chunkSize
    )
  );
  hasWorkbook = $derived(this.fileName.length > 0);
  statusLabel = $derived(
    this.isLoading ? 'Loading' : this.loadError ? 'Error' : this.hasWorkbook ? 'Ready' : 'Idle'
  );

  setViewportMetrics(viewportHeight: number, viewportWidth: number) {
    this.viewportHeight = Math.max(0, Math.floor(viewportHeight));
    this.viewportWidth = Math.max(0, Math.floor(viewportWidth));
    void this.scheduleViewportSync();
  }

  setScrollTop(scrollTop: number) {
    const normalizedScrollTop = Math.max(0, Math.floor(scrollTop));

    if (normalizedScrollTop === this.scrollTop) {
      return;
    }

    this.scrollTop = normalizedScrollTop;
    void this.scheduleViewportSync();
  }

  setSelectedRow(rowIndex: number | null) {
    this.selectedRow = rowIndex;
  }

  resetViewport() {
    this.scrollTop = 0;
    this.selectedRow = null;
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
      this.fileName = result.file_name;
      this.sheetName = result.sheet_name;
      this.totalRows = result.total_rows;
      this.totalCols = result.total_cols;
      this.scrollTop = 0;
      this.selectedRow = null;
      this.activePacket = null;
      this.lastLoadedKey = '';
      this.chunkCache.clear();
      await this.syncViewportChunk(true);

      return result;
    } catch (error) {
      this.activePacket = null;
      this.filePath = normalizedPath;
      this.fileName = '';
      this.sheetName = '';
      this.totalRows = 0;
      this.totalCols = 0;
      this.lastLoadedKey = '';
      this.loadError = error instanceof Error ? error.message : 'Excel file could not be loaded.';
      return null;
    } finally {
      this.isLoading = false;
    }
  }

  async reloadVisibleChunk() {
    await this.syncViewportChunk(true);
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

  private buildCacheKey(startRow: number, limit: number) {
    return `${startRow}:${limit}`;
  }

  private rememberChunk(chunk: SpreadsheetChunkResponse) {
    const key = this.buildCacheKey(chunk.start_row, chunk.limit);
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

  private getCachedChunk(startRow: number, limit: number) {
    const key = this.buildCacheKey(startRow, limit);
    const cached = this.chunkCache.get(key);

    if (!cached) {
      return null;
    }

    this.lastLoadedKey = key;
    return cached;
  }

  private async syncViewportChunk(forceFetch: boolean) {
    if (!this.hasWorkbook || this.totalRows <= 0) {
      this.activePacket = null;
      return;
    }

    const window = this.virtualWindow;
    if (window.requestLimit <= 0) {
      this.activePacket = null;
      return;
    }

    const cachedChunk = this.getCachedChunk(window.requestStartRow, window.requestLimit);
    if (cachedChunk && !forceFetch) {
      this.activePacket = cachedChunk;
      this.loadError = null;
      return;
    }

    const cacheKey = this.buildCacheKey(window.requestStartRow, window.requestLimit);
    const token = ++this.requestToken;
    this.isLoading = true;

    try {
      const chunk = await invokeTauriCommand<SpreadsheetChunkResponse>('get_data_chunk', {
        startRow: window.requestStartRow,
        limit: window.requestLimit
      });

      if (token !== this.requestToken) {
        return;
      }

      this.rememberChunk(chunk);
      this.activePacket = chunk;
      this.totalRows = chunk.total_rows;
      this.totalCols = chunk.total_cols;
      this.loadError = null;
      this.lastLoadedKey = cacheKey;
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
