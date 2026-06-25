<script lang="ts">
  import { isTauri } from '@tauri-apps/api/core';
  import { open } from '@tauri-apps/plugin-dialog';
  import { spreadsheetState } from '$lib/state/spreadsheetState.svelte';

  const isDesktopRuntime = isTauri();
  let filePathInput = $state('');
  let viewportElement = $state<HTMLDivElement | null>(null);
  let viewportHeight = $state(0);
  let viewportWidth = $state(0);

  const rowGridTemplate = $derived(
    `grid-template-columns: 4.5rem repeat(${Math.max(spreadsheetState.totalCols, 1)}, minmax(8rem, 10rem));`
  );
  const columnLabels = $derived.by(() =>
    Array.from({ length: spreadsheetState.totalCols }, (_, index) => excelColumnLabel(index + 1))
  );
  const activeRows = $derived(spreadsheetState.activeRows);
  const activeStartRow = $derived(spreadsheetState.activePacket?.start_row ?? 0);
  const activeLimit = $derived(spreadsheetState.activePacket?.limit ?? 0);
  const selectedRow = $derived(spreadsheetState.selectedRow);

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

  function formatCellValue(value: string | undefined) {
    const normalized = (value ?? '').trim();
    return normalized.length > 0 ? normalized : '-';
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

    if (viewportElement) {
      viewportElement.scrollTo({
        top: 0,
        left: 0,
        behavior: 'auto'
      });
    }
  }

  function focusSelectedRow(rowIndex: number) {
    spreadsheetState.setSelectedRow(rowIndex);
  }

  function handleScroll(event: Event) {
    const target = event.currentTarget as HTMLDivElement | null;

    if (!target) {
      return;
    }

    spreadsheetState.setScrollTop(target.scrollTop);
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
      <label class="text-[10px] font-semibold uppercase tracking-[0.18em] text-enterprise-text-muted" for="spreadsheet-path">
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
        <button class="enterprise-button-primary h-9 rounded-sm px-4 text-sm font-semibold disabled:cursor-not-allowed disabled:opacity-50" onclick={loadWorkbook} disabled={!isDesktopRuntime}>
          Load Excel
        </button>
        <button class="toolbar-button h-9 rounded-sm px-4 text-sm font-medium disabled:cursor-not-allowed disabled:opacity-50" onclick={selectExcelFile} disabled={!isDesktopRuntime}>
          Select File
        </button>
        <button class="toolbar-button h-9 rounded-sm px-4 text-sm font-medium disabled:cursor-not-allowed disabled:opacity-50" onclick={reloadWorkbook} disabled={!isDesktopRuntime || !spreadsheetState.filePath}>
          Reload
        </button>
        <button class="toolbar-button h-9 rounded-sm px-4 text-sm font-medium disabled:cursor-not-allowed disabled:opacity-50" onclick={resetViewport} disabled={!isDesktopRuntime || !spreadsheetState.hasWorkbook}>
          Reset View
        </button>
      </div>
      {#if spreadsheetState.loadError}
        <p class="mt-2 rounded-sm border border-[rgba(255,69,58,0.22)] bg-[rgba(255,69,58,0.10)] px-3 py-1.5 text-data-xs text-[var(--enterprise-danger)]">
          {spreadsheetState.loadError}
        </p>
      {:else}
        <p class="mt-2 text-data-xs text-enterprise-text-muted">
          Manual path entry keeps the load path deterministic for local Excel sources.
        </p>
      {/if}
    </div>

    <div class="grid min-w-[340px] flex-none gap-2 sm:grid-cols-2 xl:grid-cols-4">
      <div class="rounded-sm border border-slate-200 bg-slate-50 px-3 py-2">
        <p class="text-[10px] uppercase tracking-[0.16em] text-slate-500">File</p>
        <p class="mt-1 truncate font-mono text-data-sm font-semibold text-slate-700" title={spreadsheetState.fileName || filePathInput}>
          {spreadsheetState.fileName || 'No file loaded'}
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
      <div class="rounded-sm border border-slate-200 bg-slate-50 px-3 py-2">
        <p class="text-[10px] uppercase tracking-[0.16em] text-slate-500">State</p>
        <p class="mt-1 font-mono text-data-sm font-semibold text-slate-700">
          {spreadsheetState.statusLabel}
        </p>
      </div>
    </div>
  </div>

  <div class="flex min-h-0 flex-1 flex-col">
    {#if spreadsheetState.hasWorkbook}
      <div class="border-b border-slate-200 bg-slate-50 px-3 py-2">
        <div class="flex flex-wrap items-center gap-2 text-data-xs text-slate-600">
          <span class="font-mono font-semibold text-slate-700">{spreadsheetState.sheetName}</span>
          <span class="h-3 w-px bg-slate-300"></span>
          <span class="font-mono">{spreadsheetState.fileName}</span>
          <span class="h-3 w-px bg-slate-300"></span>
          <span>Viewport {spreadsheetState.virtualWindow.visibleStartRow + 1} - {Math.max(spreadsheetState.virtualWindow.visibleEndRow, spreadsheetState.virtualWindow.visibleStartRow + 1)}</span>
          <span class="h-3 w-px bg-slate-300"></span>
          <span>Chunk {activeStartRow + 1} - {activeStartRow + activeLimit}</span>
          <span class="h-3 w-px bg-slate-300"></span>
          <span>Selected {selectedRow !== null ? selectedRow + 1 : 'none'}</span>
        </div>
      </div>

      <div
        bind:this={viewportElement}
        bind:clientHeight={viewportHeight}
        bind:clientWidth={viewportWidth}
        class="enterprise-scrollbar min-h-0 flex-1 overflow-auto bg-white"
        onscroll={handleScroll}
      >
        <div class="min-w-max" style={`height: ${spreadsheetState.totalContentHeight + spreadsheetState.rowHeight}px; width: max-content;`}>
          <div class="sticky top-0 z-30 grid border-b border-slate-300 bg-white shadow-[0_1px_0_rgba(148,163,184,0.24)]" style={rowGridTemplate}>
            <div class="sticky left-0 z-40 flex h-7 items-center border-r border-slate-300 bg-slate-50 px-2 py-1 font-mono text-[10px] font-semibold uppercase tracking-[0.16em] text-slate-500">
              Row
            </div>
            {#each columnLabels as label}
              <div class="flex h-7 items-center border-r border-slate-200 px-2 py-1 font-mono text-[10px] font-semibold uppercase tracking-[0.12em] text-slate-500">
                {label}
              </div>
            {/each}
          </div>

          <div style={`height: ${spreadsheetState.virtualWindow.topSpacerHeight}px;`} aria-hidden="true"></div>

          {#each activeRows as row, index (activeStartRow + index)}
            {@const rowIndex = activeStartRow + index}
            <div
              class={`grid h-7 cursor-pointer border-b border-slate-200 text-enterprise-text-primary transition ${
                selectedRow === rowIndex ? 'border-l-2 border-enterprise-primary bg-slate-100' : 'border-l-2 border-transparent bg-white hover:bg-slate-50'
              }`}
              style={rowGridTemplate}
              onclick={() => focusSelectedRow(rowIndex)}
              onkeydown={(event) => {
                if (event.key === 'Enter' || event.key === ' ') {
                  event.preventDefault();
                  focusSelectedRow(rowIndex);
                }
              }}
              role="button"
              tabindex="0"
            >
              <div class="sticky left-0 z-20 flex h-7 items-center border-r border-slate-200 bg-inherit px-2 py-1 font-mono text-[11px] text-slate-500">
                {rowIndex + 1}
              </div>
              {#each row as cell, colIndex}
                <div class="flex h-7 min-w-0 items-center border-r border-slate-100 px-2 py-1 font-mono text-xs text-slate-700">
                  <span class="block truncate" title={cell ?? ''}>{formatCellValue(cell)}</span>
                </div>
              {/each}
            </div>
          {/each}

          <div
            style={`height: ${spreadsheetState.virtualWindow.bottomSpacerHeight}px;`}
            aria-hidden="true"
          ></div>
        </div>
      </div>
    {:else}
      <div class="flex flex-1 items-center justify-center px-6">
        <div class="max-w-2xl rounded-sm border border-slate-300 bg-white px-6 py-8 shadow-sm">
          <p class="text-[10px] font-semibold uppercase tracking-[0.22em] text-enterprise-text-muted">Virtual Spreadsheet</p>
          <h2 class="mt-2 text-xl font-semibold tracking-tight text-enterprise-text-primary">
            Load a local Excel workbook
          </h2>
          <p class="mt-3 max-w-xl text-sm leading-6 text-enterprise-text-secondary">
            Enter a full `.xlsx` path and load it into Rust memory. The frontend will only render the active viewport chunk, so scrolling stays smooth even on very large sheets.
          </p>
          <div class="mt-5 rounded-sm border border-slate-200 bg-slate-50 px-4 py-3 font-mono text-data-sm text-slate-600">
            Example: `C:\\data\\finance-ledger.xlsx`
          </div>
          <div class="mt-5 flex flex-wrap gap-2">
            <button class="enterprise-button-primary rounded-sm px-4 py-2 text-sm font-semibold disabled:cursor-not-allowed disabled:opacity-50" onclick={loadWorkbook} disabled={!isDesktopRuntime}>
              Load current path
            </button>
            <button class="toolbar-button rounded-sm px-4 py-2 text-sm font-medium disabled:cursor-not-allowed disabled:opacity-50" onclick={resetViewport} disabled={!isDesktopRuntime}>
              Reset viewport
            </button>
          </div>
        </div>
      </div>
    {/if}
  </div>
</section>
