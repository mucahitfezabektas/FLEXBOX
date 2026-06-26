<script lang="ts">
  import { isTauri } from '@tauri-apps/api/core';
  import { Menu } from '@tauri-apps/api/menu';
  import { open } from '@tauri-apps/plugin-dialog';
  import {
    spreadsheetState,
    type SpreadsheetColumnProfile,
    type SpreadsheetDensityMode,
    type SpreadsheetSearchMatch,
    type SpreadsheetSortDirection
  } from '$lib/state/spreadsheetState.svelte';

  type NativeMenuItem = {
    id?: string;
    text: string;
    enabled?: boolean;
    accelerator?: string;
    action?: (id: string) => void;
  };

  type RecentSelection = {
    key: string;
    address: string;
    sheetName: string;
    displayRow: number;
    sourceRow: number | null;
    colIndex: number;
    value: string;
  };

  const isDesktopRuntime = isTauri();
  let filePathInput = $state('');
  let searchInput = $state('');
  let filterInput = $state('');
  let jumpToRowInput = $state('');
  let jumpToColumnInput = $state('');
  let viewportElement = $state<HTMLDivElement | null>(null);
  let filePathInputElement = $state<HTMLInputElement | null>(null);
  let searchInputElement = $state<HTMLInputElement | null>(null);
  let filterInputElement = $state<HTMLInputElement | null>(null);
  let jumpToRowInputElement = $state<HTMLInputElement | null>(null);
  let viewportHeight = $state(0);
  let viewportWidth = $state(0);
  let showInspector = $state(true);
  let activeSearchIndex = $state(0);
  let recentSelections = $state<RecentSelection[]>([]);

  const activeRows = $derived(spreadsheetState.activeRows);
  const activeSourceRows = $derived(spreadsheetState.activeSourceRows);
  const activeStartRow = $derived(spreadsheetState.activePacket?.start_row ?? 0);
  const activeRowLimit = $derived(spreadsheetState.activePacket?.row_limit ?? 0);
  const activeStartCol = $derived(spreadsheetState.activePacket?.start_col ?? 0);
  const activeColLimit = $derived(spreadsheetState.activePacket?.col_limit ?? 0);
  const activeColumnLabels = $derived(spreadsheetState.visibleColumnLabels);
  const selectedRow = $derived(spreadsheetState.selectedRow);
  const selectedSourceRow = $derived(spreadsheetState.selectedSourceRow);
  const selectedCol = $derived(spreadsheetState.selectedCol);
  const selectedAddress = $derived.by(() =>
    selectedCol !== null
      ? `${excelColumnLabel(selectedCol + 1)}${
          (selectedSourceRow ?? selectedRow ?? 0) + 1
        }`
      : 'A1'
  );
  const selectedValue = $derived(spreadsheetState.selectedCellValue);
  const activeSheetSummary = $derived(
    spreadsheetState.sheets[spreadsheetState.activeSheetIndex] ?? null
  );
  const searchResultCount = $derived(spreadsheetState.searchResults.length);
  const activeSearchMatch = $derived(
    spreadsheetState.searchResults[activeSearchIndex] ?? null
  );
  const isWorkbookBusy = $derived(
    spreadsheetState.isLoading || spreadsheetState.isApplyingView || spreadsheetState.isSearching
  );
  const selectedDisplayRowLabel = $derived(
    selectedRow !== null ? (selectedRow + 1).toLocaleString() : '-'
  );
  const selectedSourceRowLabel = $derived(
    selectedSourceRow !== null ? (selectedSourceRow + 1).toLocaleString() : '-'
  );
  const selectedColumnLabel = $derived(
    selectedCol !== null ? excelColumnLabel(selectedCol + 1) : '-'
  );
  const selectedColumnHeader = $derived.by(() => {
    if (selectedCol === null) {
      return '-';
    }

    return activeColumnLabels[selectedCol - activeStartCol] ?? excelColumnLabel(selectedCol + 1);
  });
  const visibleChunkStats = $derived.by(() => {
    let rowCount = activeRows.length;
    let cellCount = 0;
    let nonEmptyCells = 0;
    let numericCells = 0;
    let activeColumnFilled = 0;

    for (const row of activeRows) {
      cellCount += row.length;

      row.forEach((cell, index) => {
        const normalized = (cell ?? '').trim();
        if (normalized.length === 0) {
          return;
        }

        nonEmptyCells += 1;
        if (isNumericCell(cell)) {
          numericCells += 1;
        }

        if (selectedCol !== null && activeStartCol + index === selectedCol) {
          activeColumnFilled += 1;
        }
      });
    }

    return {
      rowCount,
      cellCount,
      nonEmptyCells,
      numericCells,
      activeColumnFilled
    };
  });
  const activeColumnNumericStats = $derived.by(() => {
    if (selectedCol === null) {
      return null;
    }

    const columnOffset = selectedCol - activeStartCol;
    if (columnOffset < 0 || columnOffset >= activeColLimit) {
      return null;
    }

    const numericValues = activeRows
      .map((row) => row[columnOffset] ?? '')
      .map((value) => {
        const normalized = value.trim();
        if (!isNumericCell(normalized)) {
          return null;
        }

        const parsed = Number.parseFloat(normalized.replace(',', '.'));
        return Number.isFinite(parsed) ? parsed : null;
      })
      .filter((value): value is number => value !== null);

    if (numericValues.length === 0) {
      return null;
    }

    const sum = numericValues.reduce((accumulator, value) => accumulator + value, 0);
    const min = Math.min(...numericValues);
    const max = Math.max(...numericValues);

    return {
      count: numericValues.length,
      sum,
      min,
      max,
      avg: sum / numericValues.length
    };
  });
  const selectedRowSnapshot = $derived.by(() => {
    if (
      !spreadsheetState.activePacket ||
      selectedRow === null ||
      selectedRow < spreadsheetState.activePacket.start_row ||
      selectedRow >= spreadsheetState.activePacket.start_row + spreadsheetState.activePacket.row_limit
    ) {
      return [];
    }

    const rowOffset = selectedRow - spreadsheetState.activePacket.start_row;
    const row = spreadsheetState.activePacket.rows[rowOffset] ?? [];
    return row
      .map((value, index) => ({
        label:
          spreadsheetState.activePacket?.headers[index]?.trim() ||
          excelColumnLabel((spreadsheetState.activePacket?.start_col ?? 0) + index + 1),
        value
      }))
      .filter((entry) => entry.value.trim().length > 0)
      .slice(0, 8);
  });
  const visibleRowRange = $derived.by(() => {
    if (!spreadsheetState.hasWorkbook || spreadsheetState.totalRows === 0) {
      return '0 - 0';
    }

    const start = spreadsheetState.verticalWindow.visibleStartRow + 1;
    const end = Math.max(
      spreadsheetState.verticalWindow.visibleEndRow,
      spreadsheetState.verticalWindow.visibleStartRow + 1
    );
    return `${start} - ${end}`;
  });
  const visibleColRange = $derived.by(() => {
    if (!spreadsheetState.hasWorkbook || spreadsheetState.totalCols === 0) {
      return '0 - 0';
    }

    const start = spreadsheetState.horizontalWindow.visibleStartCol + 1;
    const end = Math.max(
      spreadsheetState.horizontalWindow.visibleEndCol,
      spreadsheetState.horizontalWindow.visibleStartCol + 1
    );
    return `${excelColumnLabel(start)} - ${excelColumnLabel(end)}`;
  });
  const viewportCoverage = $derived.by(() => {
    const rowCoverage =
      spreadsheetState.totalRows > 0
        ? Math.min(
            100,
            (spreadsheetState.verticalWindow.visibleRowCount / spreadsheetState.totalRows) * 100
          )
        : 0;
    const colCoverage =
      spreadsheetState.totalCols > 0
        ? Math.min(
            100,
            (spreadsheetState.horizontalWindow.visibleColCount / spreadsheetState.totalCols) * 100
          )
        : 0;

    return {
      rowCoverage,
      colCoverage
    };
  });
  const activeColumnSamples = $derived.by(() => {
    if (selectedCol === null) {
      return [];
    }

    const columnOffset = selectedCol - activeStartCol;
    if (columnOffset < 0 || columnOffset >= activeColLimit) {
      return [];
    }

    const sampleCounts = new Map<string, number>();

    for (const row of activeRows) {
      const normalized = (row[columnOffset] ?? '').trim();
      if (!normalized) {
        continue;
      }

      sampleCounts.set(normalized, (sampleCounts.get(normalized) ?? 0) + 1);
    }

    return Array.from(sampleCounts.entries())
      .map(([value, count]) => ({ value, count }))
      .sort((left, right) => right.count - left.count || left.value.localeCompare(right.value))
      .slice(0, 8);
  });

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

  function toHumanCount(value: number) {
    return value.toLocaleString();
  }

  function formatMetricValue(value: number) {
    return Number.isInteger(value) ? value.toLocaleString() : value.toLocaleString(undefined, {
      maximumFractionDigits: 2
    });
  }

  function formatPercent(value: number) {
    return `${value.toFixed(value >= 10 ? 1 : 2)}%`;
  }

  function columnSortMarker(colIndex: number) {
    if (spreadsheetState.sortCol !== colIndex) {
      return '';
    }

    return spreadsheetState.sortDirection === 'desc' ? 'v' : '^';
  }

  function normalizedSearchQuery() {
    return spreadsheetState.searchQuery.trim().toLowerCase();
  }

  function isNumericCell(value: string | undefined) {
    const normalized = (value ?? '').trim();
    return normalized.length > 0 && /^[-+]?(\d+([.,]\d+)?|[.,]\d+)$/.test(normalized);
  }

  function formatCellValue(value: string | undefined) {
    const normalized = value ?? '';
    return normalized.length > 0 ? normalized : ' ';
  }

  function parseColumnInput(rawValue: string) {
    const normalized = rawValue.trim().toUpperCase();
    if (!normalized) {
      return null;
    }

    if (/^\d+$/.test(normalized)) {
      const numeric = Number.parseInt(normalized, 10);
      return Number.isFinite(numeric) && numeric > 0 ? numeric - 1 : null;
    }

    if (!/^[A-Z]+$/.test(normalized)) {
      return null;
    }

    let index = 0;
    for (const character of normalized) {
      index = index * 26 + (character.charCodeAt(0) - 64);
    }

    return index - 1;
  }

  function focusElement(target: HTMLInputElement | null) {
    target?.focus();
    target?.select();
  }

  function isSearchHighlighted(value: string | undefined) {
    const query = normalizedSearchQuery();
    if (!query) {
      return false;
    }

    return (value ?? '').toLowerCase().includes(query);
  }

  function isActiveSearchCell(sourceRowIndex: number, colIndex: number) {
    if (!activeSearchMatch) {
      return false;
    }

    return activeSearchMatch.source_row === sourceRowIndex && activeSearchMatch.col === colIndex;
  }

  function syncViewportScroll() {
    if (!viewportElement) {
      return;
    }

    viewportElement.scrollTo({
      top: spreadsheetState.scrollTop,
      left: spreadsheetState.scrollLeft,
      behavior: 'auto'
    });
  }

  async function loadWorkbook() {
    await spreadsheetState.loadExcelFile(filePathInput);
    searchInput = '';
    filterInput = spreadsheetState.filterQuery;
    jumpToRowInput = '';
    activeSearchIndex = 0;
    syncViewportScroll();
  }

  async function selectExcelFile() {
    if (!isDesktopRuntime) {
      spreadsheetState.loadError =
        'File picker requires the Tauri desktop runtime. Open the app with `python run.py` or a Tauri bundle.';
      return;
    }

    try {
      const selected = await open({
        multiple: false,
        directory: false,
        filters: [
          {
            name: 'Excel Workbook',
            extensions: ['xlsx']
          }
        ]
      });

      if (!selected || Array.isArray(selected)) {
        return;
      }

      filePathInput = selected;
      await spreadsheetState.loadExcelFile(selected);
      searchInput = '';
      filterInput = spreadsheetState.filterQuery;
      jumpToRowInput = '';
      activeSearchIndex = 0;
      syncViewportScroll();
    } catch (error) {
      spreadsheetState.loadError =
        error instanceof Error ? error.message : 'Excel file picker could not be opened.';
    }
  }

  async function reloadWorkbook() {
    if (filePathInput.trim().length === 0) {
      return;
    }

    await spreadsheetState.loadExcelFile(filePathInput);
    searchInput = '';
    filterInput = spreadsheetState.filterQuery;
    jumpToRowInput = '';
    activeSearchIndex = 0;
    syncViewportScroll();
  }

  function resetViewport() {
    spreadsheetState.resetViewport();
    syncViewportScroll();
  }

  function toggleInspectorPanel() {
    showInspector = !showInspector;
  }

  function goToFirstRow() {
    spreadsheetState.goToRow(1);
    syncViewportScroll();
  }

  function goToLastRow() {
    spreadsheetState.goToRow(Math.max(spreadsheetState.totalRows, 1));
    syncViewportScroll();
  }

  function handleScroll(event: Event) {
    const target = event.currentTarget as HTMLDivElement | null;

    if (!target) {
      return;
    }

    spreadsheetState.setScrollOffsets(target.scrollTop, target.scrollLeft);
  }

  async function handleSheetChange(event: Event) {
    const target = event.currentTarget as HTMLSelectElement | null;

    if (!target) {
      return;
    }

    const sheetIndex = Number.parseInt(target.value, 10);
    if (Number.isNaN(sheetIndex)) {
      return;
    }

    await spreadsheetState.activateSheet(sheetIndex);
    searchInput = '';
    filterInput = spreadsheetState.filterQuery;
    jumpToRowInput = '';
    activeSearchIndex = 0;
    syncViewportScroll();
  }

  async function activateSheetTab(sheetIndex: number) {
    await spreadsheetState.activateSheet(sheetIndex);
    searchInput = '';
    filterInput = spreadsheetState.filterQuery;
    jumpToRowInput = '';
    activeSearchIndex = 0;
    syncViewportScroll();
  }

  async function runSearch() {
    await spreadsheetState.searchWorkbook(searchInput);
    activeSearchIndex = 0;
  }

  async function runSearchForValue(value: string) {
    const normalized = value.trim();
    if (!normalized) {
      return;
    }

    searchInput = normalized;
    await spreadsheetState.searchWorkbook(normalized);
    activeSearchIndex = 0;
  }

  function clearSearch() {
    searchInput = '';
    activeSearchIndex = 0;
    spreadsheetState.clearSearchResults();
  }

  async function applyFilter() {
    await spreadsheetState.applyFilter(filterInput);
    syncViewportScroll();
  }

  async function applyFilterValue(value: string) {
    const normalized = value.trim();
    if (!normalized) {
      return;
    }

    filterInput = normalized;
    await spreadsheetState.applyFilter(normalized);
    syncViewportScroll();
  }

  async function clearViewOptions() {
    filterInput = '';
    await spreadsheetState.clearViewOptions();
    syncViewportScroll();
  }

  async function handleHeaderToggle(event: Event) {
    const target = event.currentTarget as HTMLInputElement | null;

    if (!target) {
      return;
    }

    await spreadsheetState.setHeaderRowEnabled(target.checked);
    syncViewportScroll();
  }

  async function toggleHeaderRowMode() {
    await spreadsheetState.setHeaderRowEnabled(!spreadsheetState.headerRowEnabled);
    syncViewportScroll();
  }

  async function sortByColumn(colIndex: number) {
    await spreadsheetState.cycleColumnSort(colIndex);
    syncViewportScroll();
  }

  async function sortByColumnDirection(colIndex: number, direction: SpreadsheetSortDirection) {
    await spreadsheetState.setColumnSort(colIndex, direction);
    syncViewportScroll();
  }

  async function clearColumnSort() {
    await spreadsheetState.clearColumnSort();
    syncViewportScroll();
  }

  function focusMatch(match: SpreadsheetSearchMatch) {
    spreadsheetState.focusCell(match.display_row, match.col, match.source_row);
    syncViewportScroll();
  }

  function focusSearchResultAt(index: number) {
    const total = spreadsheetState.searchResults.length;
    if (total === 0) {
      return;
    }

    const safeIndex = ((index % total) + total) % total;
    activeSearchIndex = safeIndex;
    const match = spreadsheetState.searchResults[safeIndex];
    if (match) {
      focusMatch(match);
    }
  }

  function focusPreviousSearchResult() {
    focusSearchResultAt(activeSearchIndex - 1);
  }

  function focusNextSearchResult() {
    focusSearchResultAt(activeSearchIndex + 1);
  }

  function goToVisibleRow() {
    const rowNumber = Number.parseInt(jumpToRowInput, 10);

    if (Number.isNaN(rowNumber)) {
      return;
    }

    spreadsheetState.goToRow(rowNumber);
    syncViewportScroll();
  }

  function goToVisibleColumn() {
    const columnIndex = parseColumnInput(jumpToColumnInput);

    if (columnIndex === null) {
      return;
    }

    const targetRow = selectedRow ?? spreadsheetState.verticalWindow.visibleStartRow;
    spreadsheetState.focusCell(
      targetRow,
      columnIndex,
      spreadsheetState.resolveSourceRow(targetRow)
    );
    syncViewportScroll();
  }

  function scrubToRow(event: Event) {
    const target = event.currentTarget as HTMLInputElement | null;
    if (!target) {
      return;
    }

    const rowNumber = Number.parseInt(target.value, 10);
    if (Number.isNaN(rowNumber)) {
      return;
    }

    spreadsheetState.goToRow(rowNumber);
    syncViewportScroll();
  }

  function scrubToColumn(event: Event) {
    const target = event.currentTarget as HTMLInputElement | null;
    if (!target) {
      return;
    }

    const colNumber = Number.parseInt(target.value, 10);
    if (Number.isNaN(colNumber)) {
      return;
    }

    const targetRow = selectedRow ?? spreadsheetState.verticalWindow.visibleStartRow;
    spreadsheetState.focusCell(
      targetRow,
      Math.max(0, colNumber - 1),
      spreadsheetState.resolveSourceRow(targetRow)
    );
    syncViewportScroll();
  }

  function focusCell(displayRowIndex: number, colIndex: number, sourceRowIndex: number) {
    spreadsheetState.setSelectedCell(displayRowIndex, colIndex, sourceRowIndex);
  }

  function focusRecentSelection(item: RecentSelection) {
    spreadsheetState.focusCell(
      item.displayRow,
      item.colIndex,
      item.sourceRow
    );
    syncViewportScroll();
  }

  function setDensityMode(mode: SpreadsheetDensityMode) {
    spreadsheetState.setDensityMode(mode);
    syncViewportScroll();
  }

  function setColumnProfile(profile: SpreadsheetColumnProfile) {
    spreadsheetState.setColumnProfile(profile);
    syncViewportScroll();
  }

  async function copyVisibleChunk() {
    const lines = activeRows.map((row) => row.map((cell) => cell ?? '').join('\t'));
    await copyText(lines.join('\n'));
  }

  async function copySelectedColumnValues() {
    if (selectedCol === null) {
      return;
    }

    const columnOffset = selectedCol - activeStartCol;
    if (columnOffset < 0 || columnOffset >= activeColLimit) {
      return;
    }

    const values = activeRows.map((row) => row[columnOffset] ?? '');
    await copyText(values.join('\n'));
  }

  async function copyText(value: string) {
    if (!value) {
      return;
    }

    try {
      await navigator.clipboard.writeText(value);
    } catch {
      // Clipboard support is optional in browser fallback mode.
    }
  }

  async function copySelectionValue(value = selectedValue) {
    await copyText(value);
  }

  async function copyRowValues(row: string[]) {
    await copyText(row.map((cell) => cell ?? '').join('\t'));
  }

  async function popupNativeContextMenu(event: MouseEvent, items: NativeMenuItem[]) {
    if (!isDesktopRuntime) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();

    const menu = await Menu.new({ items });

    try {
      await menu.popup({ x: event.clientX, y: event.clientY });
    } finally {
      await menu.close().catch(() => undefined);
    }
  }

  function handleViewportContextMenu(event: MouseEvent) {
    if (!spreadsheetState.hasWorkbook) {
      return;
    }

    void popupNativeContextMenu(event, [
      {
        text: 'Reload workbook',
        enabled: Boolean(spreadsheetState.filePath),
        action: () => {
          void reloadWorkbook();
        }
      },
      {
        text: 'Reset viewport',
        enabled: spreadsheetState.hasWorkbook,
        action: () => resetViewport()
      },
      {
        text: 'Copy visible chunk',
        enabled: activeRows.length > 0,
        action: () => {
          void copyVisibleChunk();
        }
      },
      {
        text: spreadsheetState.headerRowEnabled ? 'Disable header row' : 'Use first row as headers',
        enabled: spreadsheetState.hasWorkbook,
        action: () => {
          void toggleHeaderRowMode();
        }
      },
      {
        text: 'Clear filter and sort',
        enabled:
          spreadsheetState.hasWorkbook &&
          (spreadsheetState.activeFilter.length > 0 || spreadsheetState.sortCol !== null),
        action: () => {
          void clearViewOptions();
        }
      }
    ]);
  }

  function handleColumnContextMenu(event: MouseEvent, colIndex: number, label: string) {
    void popupNativeContextMenu(event, [
      {
        text: `Sort ${label} ascending`,
        action: () => {
          void sortByColumnDirection(colIndex, 'asc');
        }
      },
      {
        text: `Sort ${label} descending`,
        action: () => {
          void sortByColumnDirection(colIndex, 'desc');
        }
      },
      {
        text: 'Clear sort',
        enabled: spreadsheetState.sortCol !== null,
        action: () => {
          void clearColumnSort();
        }
      },
      {
        text: 'Copy column label',
        action: () => {
          void copyText(label);
        }
      },
      {
        text: 'Copy visible column values',
        action: () => {
          void copyText(
            activeRows
              .map((row) => row[colIndex - activeStartCol] ?? '')
              .join('\n')
          );
        }
      }
    ]);
  }

  function handleRowContextMenu(
    event: MouseEvent,
    displayRowIndex: number,
    sourceRowIndex: number,
    row: string[]
  ) {
    spreadsheetState.setSelectedCell(
      displayRowIndex,
      Math.max(spreadsheetState.selectedCol ?? 0, 0),
      sourceRowIndex
    );

    void popupNativeContextMenu(event, [
      {
        text: `Copy row ${sourceRowIndex + 1}`,
        action: () => {
          void copyRowValues(row);
        }
      },
      {
        text: `Jump to row ${sourceRowIndex + 1}`,
        action: () => {
          spreadsheetState.goToRow(displayRowIndex + 1);
          syncViewportScroll();
        }
      },
      {
        text: 'Reset viewport',
        action: () => resetViewport()
      },
      {
        text: 'Copy visible chunk',
        action: () => {
          void copyVisibleChunk();
        }
      }
    ]);
  }

  function handleCellContextMenu(
    event: MouseEvent,
    displayRowIndex: number,
    colIndex: number,
    sourceRowIndex: number,
    cellValue: string,
    row: string[]
  ) {
    focusCell(displayRowIndex, colIndex, sourceRowIndex);

    const address = `${excelColumnLabel(colIndex + 1)}${sourceRowIndex + 1}`;
    const normalizedValue = (cellValue ?? '').trim();

    void popupNativeContextMenu(event, [
      {
        text: `Copy ${address}`,
        enabled: normalizedValue.length > 0,
        action: () => {
          void copySelectionValue(cellValue);
        }
      },
      {
        text: `Copy row ${sourceRowIndex + 1}`,
        action: () => {
          void copyRowValues(row);
        }
      },
      {
        text: `Copy visible ${excelColumnLabel(colIndex + 1)} values`,
        action: () => {
          spreadsheetState.setSelectedCell(displayRowIndex, colIndex, sourceRowIndex);
          void copySelectedColumnValues();
        }
      },
      {
        text: `Find "${normalizedValue.slice(0, 28) || 'value'}"`,
        enabled: normalizedValue.length > 0,
        action: () => {
          void runSearchForValue(normalizedValue);
        }
      },
      {
        text: `Filter by "${normalizedValue.slice(0, 28) || 'value'}"`,
        enabled: normalizedValue.length > 0,
        action: () => {
          void applyFilterValue(normalizedValue);
        }
      },
      {
        text: `Sort ${excelColumnLabel(colIndex + 1)} ascending`,
        action: () => {
          void sortByColumnDirection(colIndex, 'asc');
        }
      },
      {
        text: `Sort ${excelColumnLabel(colIndex + 1)} descending`,
        action: () => {
          void sortByColumnDirection(colIndex, 'desc');
        }
      },
      {
        text: 'Clear sort',
        enabled: spreadsheetState.sortCol !== null,
        action: () => {
          void clearColumnSort();
        }
      }
    ]);
  }

  function handleViewportKeydown(event: KeyboardEvent) {
    if (!spreadsheetState.hasWorkbook) {
      return;
    }

    if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.key.toLowerCase() === 'f') {
      event.preventDefault();
      focusElement(filterInputElement);
      return;
    }

    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f') {
      event.preventDefault();
      focusElement(searchInputElement);
      return;
    }

    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'g') {
      event.preventDefault();
      focusElement(jumpToRowInputElement);
      return;
    }

    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'l') {
      event.preventDefault();
      focusElement(filePathInputElement);
      return;
    }

    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'c') {
      event.preventDefault();
      void copySelectionValue();
      return;
    }

    if (event.key === 'F3') {
      event.preventDefault();
      if (event.shiftKey) {
        focusPreviousSearchResult();
      } else {
        focusNextSearchResult();
      }
      return;
    }

    switch (event.key) {
      case 'ArrowUp':
        event.preventDefault();
        spreadsheetState.moveSelection(-1, 0);
        syncViewportScroll();
        break;
      case 'ArrowDown':
        event.preventDefault();
        spreadsheetState.moveSelection(1, 0);
        syncViewportScroll();
        break;
      case 'ArrowLeft':
        event.preventDefault();
        spreadsheetState.moveSelection(0, -1);
        syncViewportScroll();
        break;
      case 'ArrowRight':
        event.preventDefault();
        spreadsheetState.moveSelection(0, 1);
        syncViewportScroll();
        break;
      case 'Enter':
        event.preventDefault();
        spreadsheetState.moveSelection(event.shiftKey ? -1 : 1, 0);
        syncViewportScroll();
        break;
      case 'Tab':
        event.preventDefault();
        spreadsheetState.moveSelection(0, event.shiftKey ? -1 : 1);
        syncViewportScroll();
        break;
      case 'PageUp':
        event.preventDefault();
        spreadsheetState.pageSelection('up');
        syncViewportScroll();
        break;
      case 'PageDown':
        event.preventDefault();
        spreadsheetState.pageSelection('down');
        syncViewportScroll();
        break;
      case 'Home':
        event.preventDefault();
        spreadsheetState.moveSelectionToEdge(event.ctrlKey ? 'grid-start' : 'row-start');
        syncViewportScroll();
        break;
      case 'End':
        event.preventDefault();
        spreadsheetState.moveSelectionToEdge(event.ctrlKey ? 'grid-end' : 'row-end');
        syncViewportScroll();
        break;
      default:
        break;
    }
  }

  $effect(() => {
    spreadsheetState.setViewportMetrics(viewportHeight, viewportWidth);
  });

  $effect(() => {
    if (spreadsheetState.filePath && spreadsheetState.filePath !== filePathInput) {
      filePathInput = spreadsheetState.filePath;
    }
  });

  $effect(() => {
    if (selectedCol !== null) {
      jumpToColumnInput = excelColumnLabel(selectedCol + 1);
    }
  });

  $effect(() => {
    filterInput = spreadsheetState.filterQuery;
  });

  $effect(() => {
    if (activeSearchIndex >= spreadsheetState.searchResults.length) {
      activeSearchIndex = 0;
    }
  });

  $effect(() => {
    if (
      !spreadsheetState.hasWorkbook ||
      selectedRow === null ||
      selectedCol === null
    ) {
      return;
    }

    const entry: RecentSelection = {
      key: `${spreadsheetState.sheetName}:${selectedAddress}`,
      address: selectedAddress,
      sheetName: spreadsheetState.sheetName || 'Sheet',
      displayRow: selectedRow,
      sourceRow: selectedSourceRow,
      colIndex: selectedCol,
      value: selectedValue
    };

    recentSelections = [
      entry,
      ...recentSelections.filter((item) => item.key !== entry.key)
    ].slice(0, 10);
  });
</script>

<section
  class="sheets-module -m-4 flex h-[calc(100%+2rem)] w-[calc(100%+2rem)] min-h-0 flex-col overflow-hidden border"
>
  {#if !isDesktopRuntime}
    <div class="sheets-runtime-banner px-2 py-1 text-[11px]">
      Excel loading only works inside the native Tauri desktop window opened by `python run.py`.
    </div>
  {/if}

  <div class="sheets-toolbar shrink-0 border-b">
    <div class="flex flex-wrap items-center gap-1 border-b px-2 py-1">
      <span class="sheets-chip sheets-chip-strong">SHEETS</span>
      <span class="sheets-chip font-mono" title={spreadsheetState.fileName || filePathInput}>
        {spreadsheetState.fileName || 'NO FILE'}
      </span>
      <span class="sheets-chip">{spreadsheetState.sheetName || 'NO SHEET'}</span>
      <span class="sheets-chip">Rows {toHumanCount(spreadsheetState.totalRows)}</span>
      <span class="sheets-chip">Source {toHumanCount(spreadsheetState.sourceTotalRows)}</span>
      <span class="sheets-chip">Cols {toHumanCount(spreadsheetState.totalCols)}</span>
      <span class="sheets-chip">View {visibleRowRange}</span>
      <span class="sheets-chip">Cols {visibleColRange}</span>
      <span class="sheets-chip">{spreadsheetState.statusLabel}</span>
      <span class="sheets-chip">{spreadsheetState.densityMode}</span>
      <span class="sheets-chip">{spreadsheetState.columnProfile}</span>
      {#if spreadsheetState.loadError}
        <span class="sheets-chip sheets-chip-danger">{spreadsheetState.loadError}</span>
      {/if}
    </div>

    <div class="flex flex-wrap items-center gap-1 border-b px-2 py-1">
      <input
        bind:this={filePathInputElement}
        bind:value={filePathInput}
        class="sheets-input h-8 min-w-[280px] flex-1 font-mono text-xs"
        placeholder="C:\\data\\large-workbook.xlsx"
        spellcheck="false"
        disabled={!isDesktopRuntime}
      />
      <button
        class="sheets-button-primary h-8 px-3 text-xs font-semibold disabled:cursor-not-allowed disabled:opacity-50"
        onclick={() => void loadWorkbook()}
        disabled={!isDesktopRuntime}
      >
        Load
      </button>
      <button
        class="sheets-button h-8 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
        onclick={selectExcelFile}
        disabled={!isDesktopRuntime}
      >
        Select File
      </button>
      <button
        class="sheets-button h-8 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
        onclick={() => void reloadWorkbook()}
        disabled={!isDesktopRuntime || !spreadsheetState.filePath}
      >
        Reload
      </button>
      <button
        class="sheets-button h-8 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
        onclick={resetViewport}
        disabled={!spreadsheetState.hasWorkbook}
      >
        Reset View
      </button>
      <button
        class="sheets-button h-8 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
        onclick={toggleInspectorPanel}
        disabled={!spreadsheetState.hasWorkbook}
      >
        {showInspector ? 'Hide Panel' : 'Show Panel'}
      </button>
      <select
        class="sheets-input h-8 min-w-[220px] max-w-[320px] font-mono text-xs disabled:cursor-not-allowed disabled:opacity-50"
        value={spreadsheetState.activeSheetIndex}
        onchange={handleSheetChange}
        disabled={!spreadsheetState.hasWorkbook}
      >
        {#each spreadsheetState.sheets as sheet, index}
          <option value={index}>
            {sheet.name} ({sheet.total_rows.toLocaleString()} x {sheet.total_cols.toLocaleString()})
          </option>
        {/each}
      </select>
    </div>

    <div class="flex flex-wrap items-center gap-1 border-b px-2 py-1">
      <div class="sheets-name-box h-8 w-24 font-mono text-xs font-semibold">{selectedAddress}</div>
      <div class="sheets-formula-box flex h-8 min-w-[260px] flex-1 items-center gap-2 px-2">
        <span class="sheets-fx">fx</span>
        <span class="min-w-0 flex-1 truncate font-mono text-xs">
          {selectedValue || 'Select a cell to inspect or copy its value.'}
        </span>
      </div>
      <button
        class="sheets-button h-8 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
        onclick={() => void copySelectionValue()}
        disabled={!selectedValue}
        type="button"
      >
        Copy Cell
      </button>
      <label class="inline-flex h-8 items-center gap-2 px-2 text-[11px] text-[var(--sheet-muted)]">
        <input
          type="checkbox"
          class="h-4 w-4 rounded border-[var(--sheet-border-strong)]"
          checked={spreadsheetState.headerRowEnabled}
          onchange={handleHeaderToggle}
          disabled={!spreadsheetState.hasWorkbook}
        />
        Header row
      </label>
    </div>

    <div class="flex flex-wrap items-center gap-1 px-2 py-1">
      <div class="flex min-w-[220px] flex-1 items-center gap-1">
        <input
          bind:this={searchInputElement}
          bind:value={searchInput}
          class="sheets-input h-8 min-w-0 flex-1 font-mono text-xs"
          placeholder="Find in current view"
          spellcheck="false"
          disabled={!spreadsheetState.hasWorkbook}
          onkeydown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              void runSearch();
            }
          }}
        />
        <button
          class="sheets-button h-8 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
          onclick={() => void runSearch()}
          disabled={!spreadsheetState.hasWorkbook || spreadsheetState.isSearching}
        >
          {spreadsheetState.isSearching ? 'Searching' : 'Find'}
        </button>
        <button
          class="sheets-button h-8 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
          onclick={clearSearch}
          disabled={searchResultCount === 0 && searchInput.trim().length === 0}
        >
          Clear
        </button>
        <button
          class="sheets-button h-8 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
          onclick={focusPreviousSearchResult}
          disabled={searchResultCount === 0}
        >
          Prev
        </button>
        <button
          class="sheets-button h-8 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
          onclick={focusNextSearchResult}
          disabled={searchResultCount === 0}
        >
          Next
        </button>
      </div>

      <div class="flex min-w-[260px] flex-[1.1] items-center gap-1">
        <input
          bind:this={filterInputElement}
          bind:value={filterInput}
          class="sheets-input h-8 min-w-0 flex-1 font-mono text-xs"
          placeholder="Filter rows by text"
          spellcheck="false"
          disabled={!spreadsheetState.hasWorkbook}
          onkeydown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              void applyFilter();
            }
          }}
        />
        <button
          class="sheets-button h-8 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
          onclick={() => void applyFilter()}
          disabled={!spreadsheetState.hasWorkbook || spreadsheetState.isApplyingView}
        >
          Apply
        </button>
        <button
          class="sheets-button h-8 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
          onclick={() => void clearViewOptions()}
          disabled={!spreadsheetState.hasWorkbook}
        >
          Clear
        </button>
      </div>

      <div class="flex min-w-[150px] items-center gap-1">
        <input
          bind:this={jumpToRowInputElement}
          bind:value={jumpToRowInput}
          class="sheets-input h-8 w-24 font-mono text-xs"
          placeholder="Row #"
          spellcheck="false"
          disabled={!spreadsheetState.hasWorkbook}
          onkeydown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              goToVisibleRow();
            }
          }}
        />
        <button
          class="sheets-button h-8 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
          onclick={goToVisibleRow}
          disabled={!spreadsheetState.hasWorkbook}
        >
          Go
        </button>
      </div>

      <div class="flex min-w-[170px] items-center gap-1">
        <input
          bind:value={jumpToColumnInput}
          class="sheets-input h-8 w-24 font-mono text-xs uppercase"
          placeholder="Col / AA"
          spellcheck="false"
          disabled={!spreadsheetState.hasWorkbook}
          onkeydown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              goToVisibleColumn();
            }
          }}
        />
        <button
          class="sheets-button h-8 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
          onclick={goToVisibleColumn}
          disabled={!spreadsheetState.hasWorkbook}
        >
          Column
        </button>
      </div>

      <div class="flex items-center gap-1">
        <button
          class="sheets-button h-8 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
          onclick={goToFirstRow}
          disabled={!spreadsheetState.hasWorkbook}
        >
          Top
        </button>
        <button
          class="sheets-button h-8 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
          onclick={goToLastRow}
          disabled={!spreadsheetState.hasWorkbook}
        >
          Bottom
        </button>
        <button
          class="sheets-button h-8 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
          onclick={() => void copyVisibleChunk()}
          disabled={activeRows.length === 0}
        >
          Copy View
        </button>
      </div>
    </div>
  </div>

  {#if spreadsheetState.hasWorkbook}
    <div class="flex min-h-0 flex-1">
      {#if showInspector}
        <aside class="sheets-sidebar enterprise-scrollbar hidden w-[21rem] shrink-0 overflow-auto border-r xl:block">
          <div class="sheets-panel-section border-b px-3 py-3">
            <div class="flex items-center justify-between gap-2">
              <div>
                <p class="sheets-section-label">Workbook</p>
                <h2 class="mt-1 text-sm font-semibold text-[var(--sheet-text)]">
                  {spreadsheetState.sheetName || 'Active Sheet'}
                </h2>
              </div>
              <span class="sheets-panel-badge">{isWorkbookBusy ? 'BUSY' : 'READY'}</span>
            </div>
            <div class="mt-3 grid grid-cols-2 gap-2">
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Visible rows</span>
                <span class="sheets-metric-value">{toHumanCount(spreadsheetState.totalRows)}</span>
              </div>
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Source rows</span>
                <span class="sheets-metric-value">{toHumanCount(spreadsheetState.sourceTotalRows)}</span>
              </div>
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Columns</span>
                <span class="sheets-metric-value">{toHumanCount(spreadsheetState.totalCols)}</span>
              </div>
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Viewport</span>
                <span class="sheets-metric-value">{visibleRowRange}</span>
              </div>
            </div>
            {#if activeSheetSummary}
              <div class="mt-3 rounded-sm border border-[var(--sheet-border)] bg-white/72 px-3 py-2 text-[11px] text-[var(--sheet-muted)]">
                <div class="flex items-center justify-between gap-2">
                  <span class="font-semibold text-[var(--sheet-text)]">{activeSheetSummary.name}</span>
                  <span class="font-mono">
                    {activeSheetSummary.total_rows.toLocaleString()} x {activeSheetSummary.total_cols.toLocaleString()}
                  </span>
                </div>
              </div>
            {/if}
          </div>

          <div class="sheets-panel-section border-b px-3 py-3">
            <p class="sheets-section-label">Presentation</p>
            <div class="mt-3">
              <div class="flex items-center justify-between gap-2">
                <span class="text-[11px] font-medium text-[var(--sheet-muted)]">Density</span>
                <span class="font-mono text-[11px] text-[var(--sheet-text)]">{spreadsheetState.densityMode}</span>
              </div>
              <div class="mt-2 grid grid-cols-3 gap-1">
                {#each ['compact', 'balanced', 'comfortable'] as mode}
                  <button
                    class={`sheets-toggle-button ${
                      spreadsheetState.densityMode === mode ? 'sheets-toggle-button-active' : ''
                    }`}
                    onclick={() => setDensityMode(mode as SpreadsheetDensityMode)}
                    type="button"
                  >
                    {mode}
                  </button>
                {/each}
              </div>
            </div>

            <div class="mt-3">
              <div class="flex items-center justify-between gap-2">
                <span class="text-[11px] font-medium text-[var(--sheet-muted)]">Column profile</span>
                <span class="font-mono text-[11px] text-[var(--sheet-text)]">{spreadsheetState.columnProfile}</span>
              </div>
              <div class="mt-2 grid grid-cols-3 gap-1">
                {#each ['narrow', 'standard', 'wide'] as profile}
                  <button
                    class={`sheets-toggle-button ${
                      spreadsheetState.columnProfile === profile ? 'sheets-toggle-button-active' : ''
                    }`}
                    onclick={() => setColumnProfile(profile as SpreadsheetColumnProfile)}
                    type="button"
                  >
                    {profile}
                  </button>
                {/each}
              </div>
            </div>

            <div class="mt-3 grid grid-cols-2 gap-1">
              <button class="sheets-button h-8 px-3 text-xs" onclick={resetViewport} type="button">
                Reset
              </button>
              <button class="sheets-button h-8 px-3 text-xs" onclick={() => void clearViewOptions()} type="button">
                Clear View
              </button>
              <button class="sheets-button h-8 px-3 text-xs" onclick={() => void reloadWorkbook()} type="button">
                Reload
              </button>
              <button class="sheets-button h-8 px-3 text-xs" onclick={() => void toggleHeaderRowMode()} type="button">
                {spreadsheetState.headerRowEnabled ? 'Disable Header' : 'Enable Header'}
              </button>
            </div>
          </div>

          <div class="sheets-panel-section border-b px-3 py-3">
            <div class="flex items-center justify-between gap-2">
              <p class="sheets-section-label">Navigator</p>
              <span class="font-mono text-[11px] text-[var(--sheet-muted)]">{selectedAddress}</span>
            </div>

            <div class="mt-3 rounded-sm border border-[var(--sheet-border)] bg-white/78 p-3">
              <div class="flex items-center justify-between gap-2">
                <span class="text-[11px] font-medium text-[var(--sheet-muted)]">Row sweep</span>
                <span class="font-mono text-[11px] text-[var(--sheet-text)]">{selectedDisplayRowLabel}</span>
              </div>
              <input
                class="sheets-range mt-3 w-full"
                type="range"
                min="1"
                max={Math.max(1, spreadsheetState.totalRows)}
                value={(selectedRow ?? spreadsheetState.verticalWindow.visibleStartRow) + 1}
                oninput={scrubToRow}
              />
              <div class="mt-2 flex items-center justify-between text-[10px] uppercase tracking-[0.12em] text-[var(--sheet-muted)]">
                <span>1</span>
                <span>{toHumanCount(spreadsheetState.totalRows)}</span>
              </div>
            </div>

            <div class="mt-3 rounded-sm border border-[var(--sheet-border)] bg-white/78 p-3">
              <div class="flex items-center justify-between gap-2">
                <span class="text-[11px] font-medium text-[var(--sheet-muted)]">Column sweep</span>
                <span class="font-mono text-[11px] text-[var(--sheet-text)]">{selectedColumnLabel}</span>
              </div>
              <input
                class="sheets-range mt-3 w-full"
                type="range"
                min="1"
                max={Math.max(1, spreadsheetState.totalCols)}
                value={(selectedCol ?? spreadsheetState.horizontalWindow.visibleStartCol) + 1}
                oninput={scrubToColumn}
              />
              <div class="mt-2 flex items-center justify-between text-[10px] uppercase tracking-[0.12em] text-[var(--sheet-muted)]">
                <span>A</span>
                <span>{excelColumnLabel(Math.max(1, spreadsheetState.totalCols))}</span>
              </div>
            </div>

            <div class="mt-3 grid grid-cols-2 gap-2">
              <div class="sheets-metric-card min-h-[3.5rem]">
                <span class="sheets-metric-label">Row coverage</span>
                <span class="sheets-metric-value">{formatPercent(viewportCoverage.rowCoverage)}</span>
              </div>
              <div class="sheets-metric-card min-h-[3.5rem]">
                <span class="sheets-metric-label">Col coverage</span>
                <span class="sheets-metric-value">{formatPercent(viewportCoverage.colCoverage)}</span>
              </div>
            </div>
          </div>

          <div class="sheets-panel-section border-b px-3 py-3">
            <div class="flex items-center justify-between gap-2">
              <p class="sheets-section-label">Visible Chunk</p>
              <span class="font-mono text-[11px] text-[var(--sheet-muted)]">
                {activeStartRow + 1}-{activeStartRow + activeRowLimit}
              </span>
            </div>
            <div class="mt-3 grid grid-cols-2 gap-2">
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Rows</span>
                <span class="sheets-metric-value">{visibleChunkStats.rowCount}</span>
              </div>
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Cells</span>
                <span class="sheets-metric-value">{visibleChunkStats.cellCount}</span>
              </div>
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Filled</span>
                <span class="sheets-metric-value">{visibleChunkStats.nonEmptyCells}</span>
              </div>
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Numeric</span>
                <span class="sheets-metric-value">{visibleChunkStats.numericCells}</span>
              </div>
            </div>
            <div class="mt-3 rounded-sm border border-[var(--sheet-border)] bg-white/72 px-3 py-2 text-[11px] text-[var(--sheet-muted)]">
              <div class="flex items-center justify-between gap-2">
                <span>Active column</span>
                <span class="font-mono text-[var(--sheet-text)]">{selectedColumnHeader}</span>
              </div>
              <div class="mt-1 flex items-center justify-between gap-2">
                <span>Filled values in chunk</span>
                <span class="font-mono text-[var(--sheet-text)]">{visibleChunkStats.activeColumnFilled}</span>
              </div>
            </div>
            <div class="mt-3 flex gap-1">
              <button class="sheets-button h-8 flex-1 px-3 text-xs" onclick={() => void copyVisibleChunk()} type="button">
                Copy View
              </button>
              <button
                class="sheets-button h-8 flex-1 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
                onclick={() => void copySelectedColumnValues()}
                disabled={selectedCol === null}
                type="button"
              >
                Copy Column
              </button>
            </div>
          </div>

          <div class="sheets-panel-section border-b px-3 py-3">
            <div class="flex items-center justify-between gap-2">
              <p class="sheets-section-label">Active Column Stats</p>
              <span class="sheets-panel-badge">{selectedColumnHeader}</span>
            </div>
            {#if activeColumnNumericStats}
              <div class="mt-3 grid grid-cols-2 gap-2">
                <div class="sheets-metric-card">
                  <span class="sheets-metric-label">Count</span>
                  <span class="sheets-metric-value">{activeColumnNumericStats.count}</span>
                </div>
                <div class="sheets-metric-card">
                  <span class="sheets-metric-label">Average</span>
                  <span class="sheets-metric-value">{formatMetricValue(activeColumnNumericStats.avg)}</span>
                </div>
                <div class="sheets-metric-card">
                  <span class="sheets-metric-label">Min</span>
                  <span class="sheets-metric-value">{formatMetricValue(activeColumnNumericStats.min)}</span>
                </div>
                <div class="sheets-metric-card">
                  <span class="sheets-metric-label">Max</span>
                  <span class="sheets-metric-value">{formatMetricValue(activeColumnNumericStats.max)}</span>
                </div>
              </div>
              <div class="mt-3 rounded-sm border border-[var(--sheet-border)] bg-white/72 px-3 py-2 text-[11px] text-[var(--sheet-muted)]">
                <div class="flex items-center justify-between gap-2">
                  <span>Visible sum</span>
                  <span class="font-mono text-[var(--sheet-text)]">{formatMetricValue(activeColumnNumericStats.sum)}</span>
                </div>
              </div>
            {:else}
              <p class="mt-3 text-[11px] leading-5 text-[var(--sheet-muted)]">
                Active column has no numeric values inside the current viewport chunk.
              </p>
            {/if}
          </div>

          <div class="sheets-panel-section border-b px-3 py-3">
            <div class="flex items-center justify-between gap-2">
              <p class="sheets-section-label">Active Column Samples</p>
              <span class="sheets-panel-badge">{selectedColumnHeader}</span>
            </div>

            {#if activeColumnSamples.length > 0}
              <div class="mt-3 space-y-2">
                {#each activeColumnSamples as sample}
                  <div class="rounded-sm border border-[var(--sheet-border)] bg-white/78 px-3 py-2">
                    <div class="flex items-center justify-between gap-2">
                      <span class="truncate font-mono text-[11px] font-semibold text-[var(--sheet-text)]">
                        {sample.value}
                      </span>
                      <span class="font-mono text-[10px] uppercase tracking-[0.12em] text-[var(--sheet-muted)]">
                        {sample.count} hit
                      </span>
                    </div>
                    <div class="mt-2 flex gap-1">
                      <button
                        class="sheets-button h-7 flex-1 px-2 text-[11px]"
                        onclick={() => void runSearchForValue(sample.value)}
                        type="button"
                      >
                        Find
                      </button>
                      <button
                        class="sheets-button h-7 flex-1 px-2 text-[11px]"
                        onclick={() => void applyFilterValue(sample.value)}
                        type="button"
                      >
                        Filter
                      </button>
                      <button
                        class="sheets-button h-7 px-2 text-[11px]"
                        onclick={() => void copyText(sample.value)}
                        type="button"
                      >
                        Copy
                      </button>
                    </div>
                  </div>
                {/each}
              </div>
            {:else}
              <p class="mt-3 text-[11px] leading-5 text-[var(--sheet-muted)]">
                Distinct sample values appear here for the active column inside the loaded viewport slice.
              </p>
            {/if}
          </div>

          <div class="sheets-panel-section border-b px-3 py-3">
            <div class="flex items-center justify-between gap-2">
              <p class="sheets-section-label">Search Navigator</p>
              <span class="font-mono text-[11px] text-[var(--sheet-muted)]">
                {searchResultCount > 0 ? `${activeSearchIndex + 1}/${searchResultCount}` : '0/0'}
              </span>
            </div>

            {#if activeSearchMatch}
              <button
                class="mt-3 w-full rounded-sm border border-[var(--sheet-accent)] bg-[var(--sheet-accent-soft)] px-3 py-2 text-left"
                onclick={() => focusMatch(activeSearchMatch)}
                type="button"
              >
                <div class="flex items-center justify-between gap-2">
                  <span class="font-mono text-xs font-semibold text-[var(--sheet-accent-strong)]">
                    {excelColumnLabel(activeSearchMatch.col + 1)}{activeSearchMatch.source_row + 1}
                  </span>
                  <span class="font-mono text-[11px] text-[var(--sheet-muted)]">
                    view {activeSearchMatch.display_row + 1}
                  </span>
                </div>
                <p class="mt-1 truncate font-mono text-[11px] text-[var(--sheet-text)]">
                  {activeSearchMatch.value}
                </p>
              </button>
            {/if}

            <div class="mt-3 flex gap-1">
              <button
                class="sheets-button h-8 flex-1 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
                onclick={focusPreviousSearchResult}
                disabled={searchResultCount === 0}
                type="button"
              >
                Previous
              </button>
              <button
                class="sheets-button h-8 flex-1 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
                onclick={focusNextSearchResult}
                disabled={searchResultCount === 0}
                type="button"
              >
                Next
              </button>
              <button
                class="sheets-button h-8 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
                onclick={clearSearch}
                disabled={searchResultCount === 0}
                type="button"
              >
                Clear
              </button>
            </div>

            {#if spreadsheetState.searchResults.length > 0}
              <div class="mt-3 space-y-1">
                {#each spreadsheetState.searchResults.slice(0, 12) as match, index}
                  <button
                    class={`sheets-search-result-item ${
                      activeSearchIndex === index ? 'sheets-search-result-item-active' : ''
                    }`}
                    onclick={() => focusSearchResultAt(index)}
                    type="button"
                  >
                    <span class="font-mono text-[11px] font-semibold">
                      {excelColumnLabel(match.col + 1)}{match.source_row + 1}
                    </span>
                    <span class="truncate font-mono text-[11px]">{match.value}</span>
                  </button>
                {/each}
              </div>
            {:else}
              <p class="mt-3 text-[11px] leading-5 text-[var(--sheet-muted)]">
                Search results appear here. `F3` moves to next match, `Shift+F3` moves backward.
              </p>
            {/if}
          </div>

          <div class="sheets-panel-section px-3 py-3">
            <div class="flex items-center justify-between gap-2">
              <p class="sheets-section-label">Selection Inspector</p>
              <span class="sheets-panel-badge">{selectedAddress}</span>
            </div>
            <div class="mt-3 grid grid-cols-2 gap-2">
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">View row</span>
                <span class="sheets-metric-value">{selectedDisplayRowLabel}</span>
              </div>
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Source row</span>
                <span class="sheets-metric-value">{selectedSourceRowLabel}</span>
              </div>
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Column</span>
                <span class="sheets-metric-value">{selectedColumnLabel}</span>
              </div>
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Length</span>
                <span class="sheets-metric-value">{selectedValue.length}</span>
              </div>
            </div>
            <div class="mt-3 rounded-sm border border-[var(--sheet-border)] bg-white/78 p-3">
              <p class="mb-2 text-[10px] font-semibold uppercase tracking-[0.16em] text-[var(--sheet-muted)]">
                Value Preview
              </p>
              <div class="max-h-40 overflow-auto whitespace-pre-wrap break-words font-mono text-xs leading-5 text-[var(--sheet-text)]">
                {selectedValue || 'No active cell selected.'}
              </div>
            </div>
            <div class="mt-3 rounded-sm border border-[var(--sheet-border)] bg-white/78 p-3">
              <p class="mb-2 text-[10px] font-semibold uppercase tracking-[0.16em] text-[var(--sheet-muted)]">
                Row Snapshot
              </p>
              {#if selectedRowSnapshot.length > 0}
                <div class="space-y-2">
                  {#each selectedRowSnapshot as entry}
                    <div class="rounded-sm border border-[var(--sheet-border)] bg-[var(--sheet-surface)] px-2 py-1.5">
                      <div class="truncate text-[10px] font-semibold uppercase tracking-[0.12em] text-[var(--sheet-muted)]">
                        {entry.label}
                      </div>
                      <div class="mt-1 truncate font-mono text-[11px] text-[var(--sheet-text)]">
                        {entry.value}
                      </div>
                    </div>
                  {/each}
                </div>
              {:else}
                <p class="text-[11px] leading-5 text-[var(--sheet-muted)]">
                  Current row has no non-empty cells inside the loaded viewport slice.
                </p>
              {/if}
            </div>
            <div class="mt-3 flex gap-1">
              <button
                class="sheets-button h-8 flex-1 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
                onclick={() => void copySelectionValue()}
                disabled={!selectedValue}
                type="button"
              >
                Copy Cell
              </button>
              <button
                class="sheets-button h-8 flex-1 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
                onclick={() => void runSearchForValue(selectedValue)}
                disabled={!selectedValue}
                type="button"
              >
                Find Value
              </button>
              <button
                class="sheets-button h-8 flex-1 px-3 text-xs disabled:cursor-not-allowed disabled:opacity-50"
                onclick={() => void copySelectedColumnValues()}
                disabled={selectedCol === null}
                type="button"
              >
                Copy Column
              </button>
            </div>
          </div>

          <div class="sheets-panel-section px-3 py-3">
            <div class="flex items-center justify-between gap-2">
              <p class="sheets-section-label">Recent Focus</p>
              <span class="font-mono text-[11px] text-[var(--sheet-muted)]">
                {recentSelections.length}
              </span>
            </div>
            {#if recentSelections.length > 0}
              <div class="mt-3 space-y-1">
                {#each recentSelections as item}
                  <button
                    class="sheets-search-result-item"
                    onclick={() => focusRecentSelection(item)}
                    type="button"
                  >
                    <div class="flex items-center justify-between gap-2">
                      <span class="font-mono text-[11px] font-semibold">{item.address}</span>
                      <span class="text-[10px] uppercase tracking-[0.12em] text-[var(--sheet-muted)]">
                        {item.sheetName}
                      </span>
                    </div>
                    <span class="truncate font-mono text-[11px]">{item.value || '(empty)'}</span>
                  </button>
                {/each}
              </div>
            {:else}
              <p class="mt-3 text-[11px] leading-5 text-[var(--sheet-muted)]">
                Odaklanan hücreler burada birikir. Sık döndüğünüz alanlara hızlı geri sıçrayabilirsiniz.
              </p>
            {/if}
          </div>
        </aside>
      {/if}

      <div class="flex min-h-0 min-w-0 flex-1 flex-col">
        <div class="sheets-tab-strip enterprise-scrollbar shrink-0 overflow-x-auto border-b px-1 py-1">
          <div class="flex min-w-max items-center gap-1">
            {#each spreadsheetState.sheets as sheet, index}
              <button
                class={`sheets-tab h-8 px-3 text-left ${
                  spreadsheetState.activeSheetIndex === index ? 'sheets-tab-active' : ''
                }`}
                onclick={() => void activateSheetTab(index)}
                type="button"
                title={`${sheet.name} (${sheet.total_rows.toLocaleString()} x ${sheet.total_cols.toLocaleString()})`}
              >
                <span class="block truncate text-xs font-semibold">{sheet.name}</span>
                <span class="block font-mono text-[10px] text-[var(--sheet-muted)]">
                  {sheet.total_rows.toLocaleString()} x {sheet.total_cols.toLocaleString()}
                </span>
              </button>
            {/each}
          </div>
        </div>

        <div class="sheets-statusline flex flex-wrap items-center justify-between gap-2 border-b px-2 py-1 text-[11px]">
          <div class="flex flex-wrap items-center gap-x-3 gap-y-1">
            <span>Filter {spreadsheetState.activeFilter.length > 0 ? spreadsheetState.activeFilter : 'none'}</span>
            <span>
              Sort
              {#if spreadsheetState.sortCol !== null}
                {excelColumnLabel(spreadsheetState.sortCol + 1)} {spreadsheetState.sortDirection}
              {:else}
                none
              {/if}
            </span>
            <span>Chunk rows {activeStartRow + 1} - {activeStartRow + activeRowLimit}</span>
            <span>Chunk cols {excelColumnLabel(activeStartCol + 1)} - {excelColumnLabel(activeStartCol + activeColLimit)}</span>
            <span>Matches {searchResultCount}</span>
            <span>Keys Arrows / Tab / Enter / PgUp / PgDn / F3 / Ctrl+F / Ctrl+Shift+F / Ctrl+G / Ctrl+L / Ctrl+C</span>
          </div>
          {#if isWorkbookBusy}
            <span class="font-semibold text-[var(--sheet-accent-strong)]">Working...</span>
          {/if}
        </div>

        <div class="grid grid-cols-1 gap-2 border-b border-[var(--sheet-border)] bg-white/45 px-2 py-2 xl:grid-cols-[minmax(0,1fr)_18rem]">
          <div class="rounded-sm border border-[var(--sheet-border)] bg-white/72 px-3 py-2">
            <div class="flex items-center justify-between gap-2">
              <div>
                <p class="text-[10px] font-semibold uppercase tracking-[0.16em] text-[var(--sheet-muted)]">
                  Fast Navigator
                </p>
                <p class="mt-1 text-[11px] text-[var(--sheet-muted)]">
                  Sweep across large row and column ranges without rebuilding the full dataset in the browser.
                </p>
              </div>
              <span class="sheets-panel-badge">Viewport</span>
            </div>
            <div class="mt-3 grid gap-3 lg:grid-cols-2">
              <label class="block">
                <div class="mb-2 flex items-center justify-between gap-2 text-[11px]">
                  <span class="font-medium text-[var(--sheet-muted)]">Rows</span>
                  <span class="font-mono text-[var(--sheet-text)]">{visibleRowRange}</span>
                </div>
                <input
                  class="sheets-range w-full"
                  type="range"
                  min="1"
                  max={Math.max(1, spreadsheetState.totalRows)}
                  value={(selectedRow ?? spreadsheetState.verticalWindow.visibleStartRow) + 1}
                  oninput={scrubToRow}
                />
              </label>
              <label class="block">
                <div class="mb-2 flex items-center justify-between gap-2 text-[11px]">
                  <span class="font-medium text-[var(--sheet-muted)]">Columns</span>
                  <span class="font-mono text-[var(--sheet-text)]">{visibleColRange}</span>
                </div>
                <input
                  class="sheets-range w-full"
                  type="range"
                  min="1"
                  max={Math.max(1, spreadsheetState.totalCols)}
                  value={(selectedCol ?? spreadsheetState.horizontalWindow.visibleStartCol) + 1}
                  oninput={scrubToColumn}
                />
              </label>
            </div>
          </div>

          <div class="rounded-sm border border-[var(--sheet-border)] bg-white/72 px-3 py-2">
            <div class="flex items-center justify-between gap-2">
              <p class="text-[10px] font-semibold uppercase tracking-[0.16em] text-[var(--sheet-muted)]">
                Packet Health
              </p>
              <span class="font-mono text-[11px] text-[var(--sheet-text)]">
                cache {spreadsheetState.cacheEntryCount}
              </span>
            </div>
            <div class="mt-3 grid grid-cols-2 gap-2">
              <div class="sheets-metric-card min-h-[3.25rem]">
                <span class="sheets-metric-label">Row coverage</span>
                <span class="sheets-metric-value">{formatPercent(viewportCoverage.rowCoverage)}</span>
              </div>
              <div class="sheets-metric-card min-h-[3.25rem]">
                <span class="sheets-metric-label">Col coverage</span>
                <span class="sheets-metric-value">{formatPercent(viewportCoverage.colCoverage)}</span>
              </div>
            </div>
          </div>
        </div>

        <div class="flex min-h-0 flex-1 flex-col">
          <div
            bind:this={viewportElement}
            bind:clientHeight={viewportHeight}
            bind:clientWidth={viewportWidth}
            class="sheets-grid enterprise-scrollbar min-h-0 flex-1 overflow-auto outline-none"
            onscroll={handleScroll}
            onkeydown={handleViewportKeydown}
            oncontextmenu={handleViewportContextMenu}
            role="grid"
            tabindex="0"
          >
            <div
              class="min-w-max"
              style={`height: ${spreadsheetState.totalContentHeight + spreadsheetState.rowHeight}px; width: ${
                spreadsheetState.rowHeaderWidth + spreadsheetState.totalContentWidth
              }px;`}
            >
              <div class="sheets-header sticky top-0 z-30 flex" style={`height: ${spreadsheetState.rowHeight}px;`}>
                <div
                  class="sheets-corner sticky left-0 z-40 flex items-center px-2 font-mono text-[10px] font-semibold uppercase tracking-[0.16em]"
                  style={`width: ${spreadsheetState.rowHeaderWidth}px; flex: 0 0 ${spreadsheetState.rowHeaderWidth}px;`}
                >
                  Row
                </div>
                <div
                  class="flex"
                  style={`width: ${spreadsheetState.totalContentWidth}px; flex: 0 0 ${spreadsheetState.totalContentWidth}px;`}
                >
                  <div
                    aria-hidden="true"
                    style={`width: ${spreadsheetState.horizontalWindow.leftSpacerWidth}px; flex: 0 0 ${spreadsheetState.horizontalWindow.leftSpacerWidth}px;`}
                  ></div>
                  {#each activeColumnLabels as label, colOffset}
                    {@const colIndex = activeStartCol + colOffset}
                    <button
                      class={`sheets-column-header flex shrink-0 items-center gap-2 border-r px-2 font-mono text-[10px] font-semibold tracking-[0.12em] ${
                        spreadsheetState.sortCol === colIndex ? 'sheets-column-header-active' : ''
                      }`}
                      style={`width: ${spreadsheetState.columnWidth}px; height: ${spreadsheetState.rowHeight}px; flex: 0 0 ${spreadsheetState.columnWidth}px;`}
                      onclick={() => void sortByColumn(colIndex)}
                      oncontextmenu={(event) => handleColumnContextMenu(event, colIndex, label)}
                      type="button"
                      title="Left click cycles sort. Right click opens context menu."
                    >
                      <span class="block truncate text-left uppercase">{label}</span>
                      <span class="text-[10px] text-[var(--sheet-muted)]">{columnSortMarker(colIndex)}</span>
                    </button>
                  {/each}
                  <div
                    aria-hidden="true"
                    style={`width: ${spreadsheetState.horizontalWindow.rightSpacerWidth}px; flex: 0 0 ${spreadsheetState.horizontalWindow.rightSpacerWidth}px;`}
                  ></div>
                </div>
              </div>

              <div style={`height: ${spreadsheetState.verticalWindow.topSpacerHeight}px;`} aria-hidden="true"></div>

              {#each activeRows as row, rowOffset (activeStartRow + rowOffset)}
                {@const displayRowIndex = activeStartRow + rowOffset}
                {@const sourceRowIndex = activeSourceRows[rowOffset] ?? displayRowIndex}
                <div
                  class={`sheets-row flex border-b ${
                    selectedRow === displayRowIndex ? 'sheets-row-selected' : ''
                  }`}
                  style={`height: ${spreadsheetState.rowHeight}px;`}
                >
                  <button
                    class={`sheets-row-header sticky left-0 z-20 flex items-center border-r px-2 font-mono text-[11px] ${
                      selectedRow === displayRowIndex ? 'sheets-row-header-selected' : ''
                    }`}
                    style={`width: ${spreadsheetState.rowHeaderWidth}px; height: ${spreadsheetState.rowHeight}px; flex: 0 0 ${spreadsheetState.rowHeaderWidth}px;`}
                    onclick={() =>
                      focusCell(displayRowIndex, Math.max(selectedCol ?? 0, 0), sourceRowIndex)}
                    oncontextmenu={(event) =>
                      handleRowContextMenu(event, displayRowIndex, sourceRowIndex, row)}
                    type="button"
                  >
                    {sourceRowIndex + 1}
                  </button>

                  <div
                    class="flex"
                    style={`width: ${spreadsheetState.totalContentWidth}px; flex: 0 0 ${spreadsheetState.totalContentWidth}px;`}
                  >
                    <div
                      aria-hidden="true"
                      style={`width: ${spreadsheetState.horizontalWindow.leftSpacerWidth}px; flex: 0 0 ${spreadsheetState.horizontalWindow.leftSpacerWidth}px;`}
                    ></div>

                    {#each row as cell, colOffset}
                      {@const colIndex = activeStartCol + colOffset}
                      <button
                        class={`sheets-cell flex shrink-0 items-center border-r px-2 font-mono text-xs ${
                          isNumericCell(cell)
                            ? 'justify-end text-right sheets-cell-numeric'
                            : 'justify-start text-left'
                        } ${
                          isActiveSearchCell(sourceRowIndex, colIndex)
                            ? 'sheets-cell-search-active'
                            : isSearchHighlighted(cell)
                              ? 'sheets-cell-search-hit'
                              : ''
                        } ${
                          selectedRow === displayRowIndex && selectedCol === colIndex
                            ? 'sheets-cell-selected'
                            : ''
                        }`}
                        style={`width: ${spreadsheetState.columnWidth}px; height: ${spreadsheetState.rowHeight}px; flex: 0 0 ${spreadsheetState.columnWidth}px;`}
                        onclick={() => focusCell(displayRowIndex, colIndex, sourceRowIndex)}
                        oncontextmenu={(event) =>
                          handleCellContextMenu(
                            event,
                            displayRowIndex,
                            colIndex,
                            sourceRowIndex,
                            cell ?? '',
                            row
                          )}
                        title={cell ?? ''}
                        type="button"
                      >
                        <span class="block w-full truncate">{formatCellValue(cell)}</span>
                      </button>
                    {/each}

                    <div
                      aria-hidden="true"
                      style={`width: ${spreadsheetState.horizontalWindow.rightSpacerWidth}px; flex: 0 0 ${spreadsheetState.horizontalWindow.rightSpacerWidth}px;`}
                    ></div>
                  </div>
                </div>
              {/each}

              <div style={`height: ${spreadsheetState.verticalWindow.bottomSpacerHeight}px;`} aria-hidden="true"></div>
            </div>
          </div>

          <div class="sheets-footer-bar flex flex-wrap items-center justify-between gap-2 border-t px-2 py-1 text-[11px]">
            <div class="flex flex-wrap items-center gap-x-3 gap-y-1">
              <span>Cell {selectedAddress}</span>
              <span>View row {selectedDisplayRowLabel}</span>
              <span>Source row {selectedSourceRowLabel}</span>
              <span>Column {selectedColumnLabel}</span>
              <span>Chunk cache {spreadsheetState.cacheEntryCount}</span>
              <span>Row coverage {formatPercent(viewportCoverage.rowCoverage)}</span>
              <span>Col coverage {formatPercent(viewportCoverage.colCoverage)}</span>
            </div>
            <div class="flex items-center gap-1">
              <span class="text-[var(--sheet-muted)]">Density</span>
              {#each ['compact', 'balanced', 'comfortable'] as mode}
                <button
                  class={`sheets-inline-toggle ${
                    spreadsheetState.densityMode === mode ? 'sheets-inline-toggle-active' : ''
                  }`}
                  onclick={() => setDensityMode(mode as SpreadsheetDensityMode)}
                  type="button"
                >
                  {mode}
                </button>
              {/each}
            </div>
          </div>
        </div>
      </div>
    </div>
  {:else}
    <div class="flex min-h-0 flex-1 items-center justify-center p-4">
      <div class="sheets-empty max-w-4xl border px-6 py-6">
        <p class="text-[10px] font-semibold uppercase tracking-[0.22em] text-[var(--sheet-muted)]">
          Spreadsheet Virtualization
        </p>
        <h2 class="mt-2 text-xl font-semibold tracking-tight text-[var(--sheet-text)]">
          Load a local Excel workbook
        </h2>
        <p class="mt-3 max-w-2xl text-sm leading-6 text-[var(--sheet-muted)]">
          FLEXBOX keeps workbook data in Rust memory and only paints the active viewport on the
          Svelte side. Sorting, filtering and header-row transforms stay in the backend so very
          large sheets remain responsive.
        </p>
        <div class="mt-5 grid gap-3 sm:grid-cols-3">
          <div class="sheets-empty-card">
            <span class="sheets-metric-label">Runtime</span>
            <span class="sheets-metric-value">{isDesktopRuntime ? 'Tauri' : 'Browser'}</span>
          </div>
          <div class="sheets-empty-card">
            <span class="sheets-metric-label">Virtualization</span>
            <span class="sheets-metric-value">Row + Column</span>
          </div>
          <div class="sheets-empty-card">
            <span class="sheets-metric-label">Backend</span>
            <span class="sheets-metric-value">Rust Memory</span>
          </div>
        </div>
        <div class="mt-4 rounded-sm border border-[var(--sheet-border)] bg-[var(--sheet-surface-strong)] px-3 py-2 font-mono text-xs text-[var(--sheet-muted)]">
          Example path: C:\data\finance-ledger.xlsx
        </div>
        <div class="mt-4 flex flex-wrap gap-2">
          <button
            class="sheets-button-primary px-4 py-2 text-sm font-semibold disabled:cursor-not-allowed disabled:opacity-50"
            onclick={() => void loadWorkbook()}
            disabled={!isDesktopRuntime}
          >
            Load current path
          </button>
          <button
            class="sheets-button px-4 py-2 text-sm disabled:cursor-not-allowed disabled:opacity-50"
            onclick={selectExcelFile}
            disabled={!isDesktopRuntime}
          >
            Select file
          </button>
        </div>
      </div>
    </div>
  {/if}
</section>

<style>
  .sheets-module {
    --sheet-bg: #f1efe4;
    --sheet-surface: #faf8ef;
    --sheet-surface-strong: #ece8d6;
    --sheet-surface-muted: #e3dfcb;
    --sheet-border: #c9c3ab;
    --sheet-border-strong: #a79f82;
    --sheet-text: #24271d;
    --sheet-muted: #696d58;
    --sheet-accent: #6f7458;
    --sheet-accent-strong: #4f563e;
    --sheet-accent-soft: rgba(111, 116, 88, 0.16);
    background:
      radial-gradient(circle at top left, rgba(255, 255, 255, 0.58), transparent 26%),
      linear-gradient(180deg, rgba(255, 255, 255, 0.52), rgba(255, 255, 255, 0)) 0 0 / 100% 70px
        no-repeat,
      var(--sheet-bg);
    border-color: var(--sheet-border);
    color: var(--sheet-text);
  }

  .sheets-toolbar,
  .sheets-tab-strip,
  .sheets-statusline,
  .sheets-runtime-banner,
  .sheets-footer-bar {
    background: rgba(250, 248, 239, 0.96);
    border-color: var(--sheet-border);
  }

  .sheets-runtime-banner {
    color: #8a4b11;
    background: rgba(255, 191, 77, 0.16);
    border-bottom: 1px solid rgba(138, 75, 17, 0.18);
  }

  .sheets-sidebar {
    background:
      linear-gradient(180deg, rgba(255, 255, 255, 0.24), rgba(255, 255, 255, 0)) 0 0 / 100% 80px
        no-repeat,
      rgba(244, 241, 229, 0.94);
    border-color: var(--sheet-border);
  }

  .sheets-panel-section {
    border-color: rgba(169, 159, 130, 0.48);
  }

  .sheets-section-label {
    font-size: 10px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.18em;
    color: var(--sheet-muted);
  }

  .sheets-panel-badge,
  .sheets-chip {
    display: inline-flex;
    align-items: center;
    height: 1.5rem;
    border: 1px solid var(--sheet-border);
    background: rgba(255, 255, 255, 0.75);
    color: var(--sheet-muted);
    border-radius: 0.22rem;
    padding: 0 0.5rem;
    font-size: 11px;
    white-space: nowrap;
  }

  .sheets-chip-strong {
    border-color: var(--sheet-accent-strong);
    background: linear-gradient(180deg, #73795b 0%, #565c44 100%);
    color: white;
    letter-spacing: 0.12em;
    font-weight: 700;
  }

  .sheets-chip-danger {
    max-width: 32rem;
    overflow: hidden;
    text-overflow: ellipsis;
    color: #8a1f1f;
    border-color: rgba(138, 31, 31, 0.2);
    background: rgba(191, 50, 50, 0.08);
  }

  .sheets-input,
  .sheets-button,
  .sheets-button-primary,
  .sheets-name-box,
  .sheets-formula-box,
  .sheets-tab,
  .sheets-match,
  .sheets-toggle-button,
  .sheets-inline-toggle,
  .sheets-search-result-item {
    border-radius: 0.22rem;
    border: 1px solid var(--sheet-border);
  }

  .sheets-input,
  .sheets-name-box,
  .sheets-formula-box {
    background: rgba(255, 255, 255, 0.92);
    color: var(--sheet-text);
  }

  .sheets-input::placeholder {
    color: #8a8d79;
  }

  .sheets-input:focus {
    outline: none;
    border-color: var(--sheet-accent);
    box-shadow: 0 0 0 3px rgba(111, 116, 88, 0.12);
  }

  .sheets-range {
    accent-color: var(--sheet-accent-strong);
  }

  .sheets-button,
  .sheets-match,
  .sheets-toggle-button,
  .sheets-inline-toggle,
  .sheets-search-result-item {
    background: linear-gradient(180deg, rgba(255, 255, 255, 0.9), rgba(244, 240, 225, 0.9));
    color: var(--sheet-text);
  }

  .sheets-button:hover,
  .sheets-match:hover,
  .sheets-tab:hover,
  .sheets-toggle-button:hover,
  .sheets-inline-toggle:hover,
  .sheets-search-result-item:hover {
    background: rgba(255, 255, 255, 0.98);
  }

  .sheets-button-primary {
    background: linear-gradient(180deg, #707656 0%, #565d41 100%);
    border-color: #51573e;
    color: white;
  }

  .sheets-button-primary:hover {
    background: linear-gradient(180deg, #7a8160 0%, #5d6548 100%);
  }

  .sheets-name-box {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    color: var(--sheet-accent-strong);
  }

  .sheets-formula-box {
    box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.6);
  }

  .sheets-fx {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 1.5rem;
    height: 1.5rem;
    border-radius: 999px;
    background: var(--sheet-surface-muted);
    color: var(--sheet-accent-strong);
    font-size: 10px;
    font-weight: 700;
    text-transform: uppercase;
  }

  .sheets-tab-strip {
    background: linear-gradient(180deg, rgba(240, 237, 224, 0.98), rgba(248, 246, 237, 0.98));
  }

  .sheets-tab {
    min-width: 8rem;
    background: rgba(236, 232, 214, 0.72);
    color: var(--sheet-text);
  }

  .sheets-tab-active {
    background: rgba(255, 255, 255, 0.96);
    border-color: var(--sheet-accent);
    box-shadow:
      inset 0 2px 0 var(--sheet-accent),
      0 1px 0 rgba(255, 255, 255, 0.6);
  }

  .sheets-statusline,
  .sheets-footer-bar {
    color: var(--sheet-muted);
  }

  .sheets-match {
    white-space: nowrap;
  }

  .sheets-grid {
    background:
      linear-gradient(180deg, rgba(255, 255, 255, 0.5), rgba(255, 255, 255, 0)) 0 0 / 100% 52px
        no-repeat,
      var(--sheet-surface);
    scrollbar-color: #b3ad90 transparent;
  }

  .sheets-header {
    background: rgba(247, 244, 233, 0.98);
    border-color: var(--sheet-border-strong);
    box-shadow: 0 1px 0 rgba(135, 126, 95, 0.16);
  }

  .sheets-corner,
  .sheets-column-header,
  .sheets-row-header {
    background: rgba(237, 233, 217, 0.98);
    border-color: var(--sheet-border);
    color: var(--sheet-muted);
  }

  .sheets-column-header-active {
    background: rgba(224, 219, 197, 0.98);
    color: var(--sheet-accent-strong);
  }

  .sheets-row {
    border-color: rgba(201, 195, 171, 0.52);
    background: rgba(255, 255, 255, 0.88);
  }

  .sheets-row:nth-child(even) {
    background: rgba(251, 250, 244, 0.9);
  }

  .sheets-row:hover {
    background: rgba(247, 244, 233, 0.96);
  }

  .sheets-row-selected {
    background: rgba(229, 224, 206, 0.78) !important;
  }

  .sheets-row-header-selected {
    background: rgba(219, 213, 188, 0.96);
    color: var(--sheet-accent-strong);
  }

  .sheets-cell {
    background: transparent;
    border-color: rgba(211, 206, 186, 0.48);
    color: var(--sheet-text);
  }

  .sheets-cell-numeric {
    font-feature-settings: 'tnum' 1;
    font-variant-numeric: tabular-nums;
  }

  .sheets-cell-selected {
    background: rgba(223, 218, 197, 0.96);
    box-shadow: inset 0 0 0 1px var(--sheet-accent-strong);
  }

  .sheets-cell-search-hit {
    background: rgba(111, 116, 88, 0.08);
  }

  .sheets-cell-search-active {
    background: rgba(111, 116, 88, 0.18);
    box-shadow: inset 0 0 0 1px rgba(79, 86, 62, 0.42);
  }

  .sheets-metric-card,
  .sheets-empty-card {
    display: flex;
    min-height: 4rem;
    flex-direction: column;
    justify-content: space-between;
    gap: 0.35rem;
    border-radius: 0.28rem;
    border: 1px solid var(--sheet-border);
    background: rgba(255, 255, 255, 0.76);
    padding: 0.75rem;
  }

  .sheets-metric-label {
    font-size: 10px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.16em;
    color: var(--sheet-muted);
  }

  .sheets-metric-value {
    font-family: 'IBM Plex Mono', 'Cascadia Code', 'SFMono-Regular', Consolas, monospace;
    font-size: 13px;
    font-weight: 700;
    color: var(--sheet-text);
  }

  .sheets-toggle-button,
  .sheets-inline-toggle {
    height: 2rem;
    padding: 0 0.6rem;
    font-size: 11px;
    text-transform: capitalize;
  }

  .sheets-toggle-button-active,
  .sheets-inline-toggle-active,
  .sheets-search-result-item-active {
    border-color: var(--sheet-accent);
    background: var(--sheet-accent-soft);
    color: var(--sheet-accent-strong);
  }

  .sheets-search-result-item {
    display: flex;
    width: 100%;
    flex-direction: column;
    gap: 0.2rem;
    padding: 0.55rem 0.65rem;
    text-align: left;
  }

  .sheets-empty {
    border-color: var(--sheet-border);
    background: rgba(250, 248, 239, 0.96);
    box-shadow: 0 20px 42px rgba(79, 86, 62, 0.08);
  }
</style>
