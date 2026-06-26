<script lang="ts">
  import { isTauri } from '@tauri-apps/api/core';
  import { Menu } from '@tauri-apps/api/menu';
  import { open } from '@tauri-apps/plugin-dialog';
  import {
    spreadsheetState,
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

  const isDesktopRuntime = isTauri();
  let filePathInput = $state('');
  let searchInput = $state('');
  let filterInput = $state('');
  let jumpToRowInput = $state('');
  let viewportElement = $state<HTMLDivElement | null>(null);
  let viewportHeight = $state(0);
  let viewportWidth = $state(0);

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
  const searchResultCount = $derived(spreadsheetState.searchResults.length);
  const isWorkbookBusy = $derived(
    spreadsheetState.isLoading || spreadsheetState.isApplyingView || spreadsheetState.isSearching
  );
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

  function columnSortMarker(colIndex: number) {
    if (spreadsheetState.sortCol !== colIndex) {
      return '';
    }

    return spreadsheetState.sortDirection === 'desc' ? 'v' : '^';
  }

  function isNumericCell(value: string | undefined) {
    const normalized = (value ?? '').trim();
    return normalized.length > 0 && /^[-+]?(\d+([.,]\d+)?|[.,]\d+)$/.test(normalized);
  }

  function formatCellValue(value: string | undefined) {
    const normalized = value ?? '';
    return normalized.length > 0 ? normalized : ' ';
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
    syncViewportScroll();
  }

  function resetViewport() {
    spreadsheetState.resetViewport();
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
    syncViewportScroll();
  }

  async function activateSheetTab(sheetIndex: number) {
    await spreadsheetState.activateSheet(sheetIndex);
    searchInput = '';
    filterInput = spreadsheetState.filterQuery;
    jumpToRowInput = '';
    syncViewportScroll();
  }

  async function runSearch() {
    await spreadsheetState.searchWorkbook(searchInput);
  }

  async function runSearchForValue(value: string) {
    const normalized = value.trim();
    if (!normalized) {
      return;
    }

    searchInput = normalized;
    await spreadsheetState.searchWorkbook(normalized);
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

  function goToVisibleRow() {
    const rowNumber = Number.parseInt(jumpToRowInput, 10);

    if (Number.isNaN(rowNumber)) {
      return;
    }

    spreadsheetState.goToRow(rowNumber);
    syncViewportScroll();
  }

  function focusCell(displayRowIndex: number, colIndex: number, sourceRowIndex: number) {
    spreadsheetState.setSelectedCell(displayRowIndex, colIndex, sourceRowIndex);
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

    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'c') {
      event.preventDefault();
      void copySelectionValue();
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
    filterInput = spreadsheetState.filterQuery;
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
      <span class="sheets-chip">Viewport {visibleRowRange}</span>
      <span class="sheets-chip">Cols {visibleColRange}</span>
      <span class="sheets-chip">{spreadsheetState.statusLabel}</span>
      {#if spreadsheetState.loadError}
        <span class="sheets-chip sheets-chip-danger">{spreadsheetState.loadError}</span>
      {/if}
    </div>

    <div class="flex flex-wrap items-center gap-1 border-b px-2 py-1">
      <input
        bind:value={filePathInput}
        class="sheets-input h-8 min-w-[260px] flex-1 font-mono text-xs"
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
      <div class="sheets-formula-box flex h-8 min-w-[280px] flex-1 items-center gap-2 px-2">
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
      </div>

      <div class="flex min-w-[260px] flex-[1.1] items-center gap-1">
        <input
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
    </div>
  </div>

  {#if spreadsheetState.hasWorkbook}
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

    <div class="sheets-statusline flex flex-wrap items-center gap-x-3 gap-y-1 border-b px-2 py-1 text-[11px]">
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
      <span>Keys Arrows / Tab / Enter / PgUp / PgDn / Ctrl+C</span>
      {#if isWorkbookBusy}
        <span class="font-semibold text-[var(--sheet-accent-strong)]">Working...</span>
      {/if}
    </div>

    {#if spreadsheetState.searchResults.length > 0}
      <div class="enterprise-scrollbar shrink-0 overflow-x-auto border-b px-2 py-1">
        <div class="flex min-w-max items-center gap-1">
          {#each spreadsheetState.searchResults.slice(0, 18) as match}
            <button
              class="sheets-match px-2 py-1 font-mono text-[11px]"
              onclick={() => focusMatch(match)}
            >
              {excelColumnLabel(match.col + 1)}{match.source_row + 1} [view {match.display_row + 1}]
              {match.value.slice(0, 28)}
            </button>
          {/each}
        </div>
      </div>
    {/if}

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
          <div class="sheets-header sticky top-0 z-30 flex h-7 border-b">
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
                  class={`sheets-column-header flex h-7 shrink-0 items-center gap-2 border-r px-2 font-mono text-[10px] font-semibold tracking-[0.12em] ${
                    spreadsheetState.sortCol === colIndex ? 'sheets-column-header-active' : ''
                  }`}
                  style={`width: ${spreadsheetState.columnWidth}px; flex: 0 0 ${spreadsheetState.columnWidth}px;`}
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
              class={`sheets-row flex h-7 border-b ${
                selectedRow === displayRowIndex ? 'sheets-row-selected' : ''
              }`}
            >
              <button
                class={`sheets-row-header sticky left-0 z-20 flex items-center border-r px-2 font-mono text-[11px] ${
                  selectedRow === displayRowIndex ? 'sheets-row-header-selected' : ''
                }`}
                style={`width: ${spreadsheetState.rowHeaderWidth}px; flex: 0 0 ${spreadsheetState.rowHeaderWidth}px;`}
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
                    class={`sheets-cell flex h-7 shrink-0 items-center border-r px-2 font-mono text-xs ${
                      isNumericCell(cell)
                        ? 'justify-end text-right sheets-cell-numeric'
                        : 'justify-start text-left'
                    } ${
                      selectedRow === displayRowIndex && selectedCol === colIndex
                        ? 'sheets-cell-selected'
                        : ''
                    }`}
                    style={`width: ${spreadsheetState.columnWidth}px; flex: 0 0 ${spreadsheetState.columnWidth}px;`}
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
    </div>
  {:else}
    <div class="flex min-h-0 flex-1 items-center justify-center p-4">
      <div class="sheets-empty max-w-3xl border px-6 py-6">
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
      linear-gradient(180deg, rgba(255, 255, 255, 0.6), rgba(255, 255, 255, 0)) 0 0 / 100% 60px
        no-repeat,
      var(--sheet-bg);
    border-color: var(--sheet-border);
    color: var(--sheet-text);
  }

  .sheets-toolbar,
  .sheets-tab-strip,
  .sheets-statusline,
  .sheets-runtime-banner {
    background: rgba(250, 248, 239, 0.96);
    border-color: var(--sheet-border);
  }

  .sheets-runtime-banner {
    color: #8a4b11;
    background: rgba(255, 191, 77, 0.16);
    border-bottom: 1px solid rgba(138, 75, 17, 0.18);
  }

  .sheets-chip {
    display: inline-flex;
    align-items: center;
    height: 1.5rem;
    border: 1px solid var(--sheet-border);
    background: rgba(255, 255, 255, 0.75);
    color: var(--sheet-muted);
    border-radius: 0.2rem;
    padding: 0 0.45rem;
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
  .sheets-match {
    border-radius: 0.2rem;
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

  .sheets-button,
  .sheets-match {
    background: linear-gradient(180deg, rgba(255, 255, 255, 0.88), rgba(244, 240, 225, 0.88));
    color: var(--sheet-text);
  }

  .sheets-button:hover,
  .sheets-match:hover,
  .sheets-tab:hover {
    background: rgba(255, 255, 255, 0.96);
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

  .sheets-statusline {
    color: var(--sheet-muted);
  }

  .sheets-match {
    white-space: nowrap;
  }

  .sheets-grid {
    background:
      linear-gradient(180deg, rgba(255, 255, 255, 0.5), rgba(255, 255, 255, 0)) 0 0 / 100% 48px
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

  .sheets-row:hover {
    background: rgba(247, 244, 233, 0.9);
  }

  .sheets-row-selected {
    background: rgba(229, 224, 206, 0.72);
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
    background: rgba(223, 218, 197, 0.95);
    box-shadow: inset 0 0 0 1px var(--sheet-accent-strong);
  }

  .sheets-empty {
    border-color: var(--sheet-border);
    background: rgba(250, 248, 239, 0.96);
    box-shadow: 0 20px 42px rgba(79, 86, 62, 0.08);
  }
</style>
