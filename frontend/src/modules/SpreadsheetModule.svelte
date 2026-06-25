<script lang="ts">
  import { isTauri } from '@tauri-apps/api/core';
  import { open } from '@tauri-apps/plugin-dialog';
  import {
    spreadsheetState,
    type SpreadsheetSearchMatch
  } from '$lib/state/spreadsheetState.svelte';

  const isDesktopRuntime = isTauri();
  let filePathInput = $state('');
  let searchInput = $state('');
  let jumpToRowInput = $state('');
  let viewportElement = $state<HTMLDivElement | null>(null);
  let viewportHeight = $state(0);
  let viewportWidth = $state(0);

  const activeRows = $derived(spreadsheetState.activeRows);
  const activeStartRow = $derived(spreadsheetState.activePacket?.start_row ?? 0);
  const activeRowLimit = $derived(spreadsheetState.activePacket?.row_limit ?? 0);
  const activeStartCol = $derived(spreadsheetState.activePacket?.start_col ?? 0);
  const activeColLimit = $derived(spreadsheetState.activePacket?.col_limit ?? 0);
  const activeColumnLabels = $derived(spreadsheetState.visibleColumnLabels);
  const selectedRow = $derived(spreadsheetState.selectedRow);
  const selectedCol = $derived(spreadsheetState.selectedCol);
  const selectedAddress = $derived.by(() =>
    selectedRow !== null && selectedCol !== null
      ? `${excelColumnLabel(selectedCol + 1)}${selectedRow + 1}`
      : 'None'
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

  function isNumericCell(value: string | undefined) {
    const normalized = (value ?? '').trim();
    return normalized.length > 0 && /^[-+]?(\d+([.,]\d+)?|[.,]\d+)$/.test(normalized);
  }

  function formatCellValue(value: string | undefined) {
    return value ?? '';
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

  function loadWorkbook() {
    void spreadsheetState.loadExcelFile(filePathInput);
  }

  async function selectExcelFile() {
    if (!isTauri()) {
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
      jumpToRowInput = '';
      syncViewportScroll();
    } catch (error) {
      spreadsheetState.loadError =
        error instanceof Error ? error.message : 'Excel file picker could not be opened.';
    }
  }

  function reloadWorkbook() {
    if (filePathInput.trim().length === 0) {
      return;
    }

    void spreadsheetState.loadExcelFile(filePathInput);
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
    jumpToRowInput = '';
    syncViewportScroll();
  }

  async function runSearch() {
    await spreadsheetState.searchWorkbook(searchInput);
  }

  function focusMatch(match: SpreadsheetSearchMatch) {
    spreadsheetState.focusCell(match.row, match.col);
    syncViewportScroll();
  }

  function goToRow() {
    const rowNumber = Number.parseInt(jumpToRowInput, 10);

    if (Number.isNaN(rowNumber)) {
      return;
    }

    spreadsheetState.goToRow(rowNumber);
    syncViewportScroll();
  }

  function focusCell(rowIndex: number, colIndex: number) {
    spreadsheetState.setSelectedCell(rowIndex, colIndex);
  }

  $effect(() => {
    spreadsheetState.setViewportMetrics(viewportHeight, viewportWidth);
  });

  $effect(() => {
    if (spreadsheetState.filePath && spreadsheetState.filePath !== filePathInput) {
      filePathInput = spreadsheetState.filePath;
    }
  });
</script>

<section class="flex h-full min-h-0 flex-col border border-slate-300 bg-enterprise-surface text-enterprise-text-primary">
  <div class="flex flex-wrap items-end gap-3 border-b border-slate-200 bg-white px-3 py-2">
    <div class="min-w-0 flex-1">
      <label
        class="text-[10px] font-semibold uppercase tracking-[0.18em] text-enterprise-text-muted"
        for="spreadsheet-path"
      >
        Excel path
      </label>
      {#if !isDesktopRuntime}
        <p class="mt-2 rounded-sm border border-[rgba(217,119,6,0.22)] bg-[rgba(255,159,10,0.10)] px-3 py-1.5 text-data-xs text-[var(--enterprise-warning)]">
          You are viewing the Vite browser page. Excel loading only works inside the native Tauri desktop window opened by `python run.py`.
        </p>
      {/if}
      <div class="mt-1 flex flex-wrap gap-2">
        <input
          id="spreadsheet-path"
          bind:value={filePathInput}
          class="enterprise-input h-9 min-w-0 flex-1 rounded-sm px-3 font-mono text-sm"
          placeholder="C:\\data\\large-workbook.xlsx"
          spellcheck="false"
          disabled={!isDesktopRuntime}
        />
        <button
          class="enterprise-button-primary h-9 rounded-sm px-4 text-sm font-semibold disabled:cursor-not-allowed disabled:opacity-50"
          onclick={loadWorkbook}
          disabled={!isDesktopRuntime}
        >
          Load Excel
        </button>
        <button
          class="toolbar-button h-9 rounded-sm px-4 text-sm font-medium disabled:cursor-not-allowed disabled:opacity-50"
          onclick={selectExcelFile}
          disabled={!isDesktopRuntime}
        >
          Select File
        </button>
        <button
          class="toolbar-button h-9 rounded-sm px-4 text-sm font-medium disabled:cursor-not-allowed disabled:opacity-50"
          onclick={reloadWorkbook}
          disabled={!isDesktopRuntime || !spreadsheetState.filePath}
        >
          Reload
        </button>
        <button
          class="toolbar-button h-9 rounded-sm px-4 text-sm font-medium disabled:cursor-not-allowed disabled:opacity-50"
          onclick={resetViewport}
          disabled={!isDesktopRuntime || !spreadsheetState.hasWorkbook}
        >
          Reset View
        </button>
      </div>
      {#if spreadsheetState.loadError}
        <p class="mt-2 rounded-sm border border-[rgba(255,69,58,0.22)] bg-[rgba(255,69,58,0.10)] px-3 py-1.5 text-data-xs text-[var(--enterprise-danger)]">
          {spreadsheetState.loadError}
        </p>
      {:else}
        <p class="mt-2 text-data-xs text-enterprise-text-muted">
          Workbook is held in Rust memory. FLEXBOX only requests the active row and column window for rendering.
        </p>
      {/if}
    </div>

    <div class="grid min-w-[340px] flex-none gap-2 sm:grid-cols-2 xl:grid-cols-4">
      <div class="rounded-sm border border-slate-200 bg-slate-50 px-3 py-2">
        <p class="text-[10px] uppercase tracking-[0.16em] text-slate-500">File</p>
        <p
          class="mt-1 truncate font-mono text-data-sm font-semibold text-slate-700"
          title={spreadsheetState.fileName || filePathInput}
        >
          {spreadsheetState.fileName || 'No file loaded'}
        </p>
      </div>
      <div class="rounded-sm border border-slate-200 bg-slate-50 px-3 py-2">
        <p class="text-[10px] uppercase tracking-[0.16em] text-slate-500">Sheet</p>
        <p class="mt-1 font-mono text-data-sm font-semibold text-slate-700">
          {spreadsheetState.sheetName || 'None'}
        </p>
      </div>
      <div class="rounded-sm border border-slate-200 bg-slate-50 px-3 py-2">
        <p class="text-[10px] uppercase tracking-[0.16em] text-slate-500">Rows</p>
        <p class="mt-1 font-mono text-data-sm font-semibold text-slate-700">
          {spreadsheetState.totalRows.toLocaleString()}
        </p>
      </div>
      <div class="rounded-sm border border-slate-200 bg-slate-50 px-3 py-2">
        <p class="text-[10px] uppercase tracking-[0.16em] text-slate-500">Columns</p>
        <p class="mt-1 font-mono text-data-sm font-semibold text-slate-700">
          {spreadsheetState.totalCols.toLocaleString()}
        </p>
      </div>
    </div>
  </div>

  <div class="flex flex-wrap items-end gap-3 border-b border-slate-200 bg-slate-50 px-3 py-2">
    <div class="min-w-[220px] flex-1">
      <label class="text-[10px] font-semibold uppercase tracking-[0.16em] text-slate-500" for="sheet-select">
        Workbook Sheets
      </label>
      <select
        id="sheet-select"
        class="enterprise-input mt-1 h-9 w-full rounded-sm px-3 font-mono text-sm disabled:cursor-not-allowed disabled:opacity-50"
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

    <div class="min-w-[260px] flex-1">
      <label class="text-[10px] font-semibold uppercase tracking-[0.16em] text-slate-500" for="search-input">
        Find In Sheet
      </label>
      <div class="mt-1 flex gap-2">
        <input
          id="search-input"
          bind:value={searchInput}
          class="enterprise-input h-9 min-w-0 flex-1 rounded-sm px-3 font-mono text-sm"
          placeholder="Search visible workbook values"
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
          class="toolbar-button h-9 rounded-sm px-4 text-sm font-medium disabled:cursor-not-allowed disabled:opacity-50"
          onclick={runSearch}
          disabled={!spreadsheetState.hasWorkbook || spreadsheetState.isSearching}
        >
          {spreadsheetState.isSearching ? 'Searching' : 'Find'}
        </button>
      </div>
    </div>

    <div class="min-w-[180px]">
      <label class="text-[10px] font-semibold uppercase tracking-[0.16em] text-slate-500" for="goto-row">
        Jump To Row
      </label>
      <div class="mt-1 flex gap-2">
        <input
          id="goto-row"
          bind:value={jumpToRowInput}
          class="enterprise-input h-9 min-w-0 flex-1 rounded-sm px-3 font-mono text-sm"
          placeholder="125000"
          spellcheck="false"
          disabled={!spreadsheetState.hasWorkbook}
          onkeydown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              goToRow();
            }
          }}
        />
        <button
          class="toolbar-button h-9 rounded-sm px-4 text-sm font-medium disabled:cursor-not-allowed disabled:opacity-50"
          onclick={goToRow}
          disabled={!spreadsheetState.hasWorkbook}
        >
          Go
        </button>
      </div>
    </div>
  </div>

  {#if spreadsheetState.searchResults.length > 0}
    <div class="border-b border-slate-200 bg-white px-3 py-2">
      <div class="flex flex-wrap gap-2">
        {#each spreadsheetState.searchResults.slice(0, 12) as match}
          <button
            class="rounded-sm border border-slate-300 bg-slate-50 px-2 py-1 font-mono text-[11px] text-slate-700 transition hover:border-slate-400 hover:bg-slate-100"
            onclick={() => focusMatch(match)}
          >
            {excelColumnLabel(match.col + 1)}{match.row + 1}: {match.value.slice(0, 36)}
          </button>
        {/each}
      </div>
    </div>
  {/if}

  <div class="flex min-h-0 flex-1 flex-col">
    {#if spreadsheetState.hasWorkbook}
      <div class="border-b border-slate-200 bg-slate-50 px-3 py-2">
        <div class="flex flex-wrap items-center gap-2 text-data-xs text-slate-600">
          <span class="font-mono font-semibold text-slate-700">{spreadsheetState.sheetName}</span>
          <span class="h-3 w-px bg-slate-300"></span>
          <span>Rows {visibleRowRange}</span>
          <span class="h-3 w-px bg-slate-300"></span>
          <span>Cols {visibleColRange}</span>
          <span class="h-3 w-px bg-slate-300"></span>
          <span>Chunk rows {activeStartRow + 1} - {activeStartRow + activeRowLimit}</span>
          <span class="h-3 w-px bg-slate-300"></span>
          <span>Chunk cols {excelColumnLabel(activeStartCol + 1)} - {excelColumnLabel(activeStartCol + activeColLimit)}</span>
          <span class="h-3 w-px bg-slate-300"></span>
          <span>Cell {selectedAddress}</span>
          <span class="h-3 w-px bg-slate-300"></span>
          <span>Status {spreadsheetState.statusLabel}</span>
        </div>
      </div>

      <div
        bind:this={viewportElement}
        bind:clientHeight={viewportHeight}
        bind:clientWidth={viewportWidth}
        class="enterprise-scrollbar min-h-0 flex-1 overflow-auto bg-white"
        onscroll={handleScroll}
      >
        <div
          class="min-w-max"
          style={`height: ${spreadsheetState.totalContentHeight + spreadsheetState.rowHeight}px; width: ${
            spreadsheetState.rowHeaderWidth + spreadsheetState.totalContentWidth
          }px;`}
        >
          <div class="sticky top-0 z-30 flex h-7 border-b border-slate-300 bg-white shadow-[0_1px_0_rgba(148,163,184,0.24)]">
            <div
              class="sticky left-0 z-40 flex items-center border-r border-slate-300 bg-slate-50 px-2 py-1 font-mono text-[10px] font-semibold uppercase tracking-[0.16em] text-slate-500"
              style={`width: ${spreadsheetState.rowHeaderWidth}px; flex: 0 0 ${spreadsheetState.rowHeaderWidth}px;`}
            >
              Row
            </div>
            <div class="flex" style={`width: ${spreadsheetState.totalContentWidth}px; flex: 0 0 ${spreadsheetState.totalContentWidth}px;`}>
              <div
                aria-hidden="true"
                style={`width: ${spreadsheetState.horizontalWindow.leftSpacerWidth}px; flex: 0 0 ${spreadsheetState.horizontalWindow.leftSpacerWidth}px;`}
              ></div>
              {#each activeColumnLabels as label}
                <div
                  class="flex h-7 shrink-0 items-center border-r border-slate-200 px-2 py-1 font-mono text-[10px] font-semibold uppercase tracking-[0.12em] text-slate-500"
                  style={`width: ${spreadsheetState.columnWidth}px; flex: 0 0 ${spreadsheetState.columnWidth}px;`}
                >
                  {label}
                </div>
              {/each}
              <div
                aria-hidden="true"
                style={`width: ${spreadsheetState.horizontalWindow.rightSpacerWidth}px; flex: 0 0 ${spreadsheetState.horizontalWindow.rightSpacerWidth}px;`}
              ></div>
            </div>
          </div>

          <div style={`height: ${spreadsheetState.verticalWindow.topSpacerHeight}px;`} aria-hidden="true"></div>

          {#each activeRows as row, rowOffset (activeStartRow + rowOffset)}
            {@const rowIndex = activeStartRow + rowOffset}
            <div
              class={`flex h-7 border-b border-slate-200 text-enterprise-text-primary transition ${
                selectedRow === rowIndex ? 'bg-slate-50' : 'bg-white hover:bg-slate-50'
              }`}
            >
              <div
                class={`sticky left-0 z-20 flex items-center border-r border-slate-200 px-2 py-1 font-mono text-[11px] ${
                  selectedRow === rowIndex ? 'bg-slate-100 text-slate-700' : 'bg-inherit text-slate-500'
                }`}
                style={`width: ${spreadsheetState.rowHeaderWidth}px; flex: 0 0 ${spreadsheetState.rowHeaderWidth}px;`}
              >
                {rowIndex + 1}
              </div>

              <div class="flex" style={`width: ${spreadsheetState.totalContentWidth}px; flex: 0 0 ${spreadsheetState.totalContentWidth}px;`}>
                <div
                  aria-hidden="true"
                  style={`width: ${spreadsheetState.horizontalWindow.leftSpacerWidth}px; flex: 0 0 ${spreadsheetState.horizontalWindow.leftSpacerWidth}px;`}
                ></div>

                {#each row as cell, colOffset}
                  {@const colIndex = activeStartCol + colOffset}
                  <button
                    class={`flex h-7 shrink-0 items-center border-r border-slate-100 px-2 py-1 font-mono text-xs ${
                      isNumericCell(cell) ? 'justify-end text-right' : 'justify-start text-left'
                    } ${
                      selectedRow === rowIndex && selectedCol === colIndex
                        ? 'bg-slate-100 shadow-[inset_0_0_0_1px_var(--enterprise-primary)]'
                        : ''
                    }`}
                    style={`width: ${spreadsheetState.columnWidth}px; flex: 0 0 ${spreadsheetState.columnWidth}px;`}
                    onclick={() => focusCell(rowIndex, colIndex)}
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
    {:else}
      <div class="flex flex-1 items-center justify-center px-6">
        <div class="max-w-2xl rounded-sm border border-slate-300 bg-white px-6 py-8 shadow-sm">
          <p class="text-[10px] font-semibold uppercase tracking-[0.22em] text-enterprise-text-muted">
            Virtual Spreadsheet
          </p>
          <h2 class="mt-2 text-xl font-semibold tracking-tight text-enterprise-text-primary">
            Load a local Excel workbook
          </h2>
          <p class="mt-3 max-w-xl text-sm leading-6 text-enterprise-text-secondary">
            FLEXBOX reads workbook data in Rust and only paints the active viewport on the Svelte side.
            Sheet switching, search and large row jumps stay inside the same workbook session.
          </p>
          <div class="mt-5 rounded-sm border border-slate-200 bg-slate-50 px-4 py-3 font-mono text-data-sm text-slate-600">
            Example: `C:\\data\\finance-ledger.xlsx`
          </div>
          <div class="mt-5 flex flex-wrap gap-2">
            <button
              class="enterprise-button-primary rounded-sm px-4 py-2 text-sm font-semibold disabled:cursor-not-allowed disabled:opacity-50"
              onclick={loadWorkbook}
              disabled={!isDesktopRuntime}
            >
              Load current path
            </button>
            <button
              class="toolbar-button rounded-sm px-4 py-2 text-sm font-medium disabled:cursor-not-allowed disabled:opacity-50"
              onclick={selectExcelFile}
              disabled={!isDesktopRuntime}
            >
              Select file
            </button>
          </div>
        </div>
      </div>
    {/if}
  </div>
</section>
