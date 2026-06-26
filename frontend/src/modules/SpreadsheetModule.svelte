<script lang="ts">
  import { isTauri } from "@tauri-apps/api/core";
  import { Menu } from "@tauri-apps/api/menu";
  import { PhysicalPosition } from "@tauri-apps/api/window";
  import { open } from "@tauri-apps/plugin-dialog";
  import {
    spreadsheetState,
    type SpreadsheetColumnProfile,
    type SpreadsheetDensityMode,
    type SpreadsheetSearchMatch,
    type SpreadsheetSortDirection,
  } from "$lib/state/spreadsheetState.svelte";

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
  let filePathInput = $state("");
  let searchInput = $state("");
  let filterInput = $state("");
  let jumpToRowInput = $state("");
  let jumpToColumnInput = $state("");
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

  let canvasElement = $state<HTMLCanvasElement | null>(null);
  let canvasContext = $state<CanvasRenderingContext2D | null>(null);

  // $derived State Tanımlamaları
  const activeRows = $derived(spreadsheetState.activeRows);
  const activeSourceRows = $derived(spreadsheetState.activeSourceRows);
  const activeStartRow = $derived(
    spreadsheetState.activePacket?.start_row ?? 0,
  );
  const activeRowLimit = $derived(
    spreadsheetState.activePacket?.row_limit ?? 0,
  );
  const activeStartCol = $derived(
    spreadsheetState.activePacket?.start_col ?? 0,
  );
  const activeColLimit = $derived(
    spreadsheetState.activePacket?.col_limit ?? 0,
  );
  const activeColumnLabels = $derived(spreadsheetState.visibleColumnLabels);
  const selectedRow = $derived(spreadsheetState.selectedRow);
  const selectedSourceRow = $derived(spreadsheetState.selectedSourceRow);
  const selectedCol = $derived(spreadsheetState.selectedCol);
  const selectedAddress = $derived.by(() =>
    selectedCol !== null
      ? `${excelColumnLabel(selectedCol + 1)}${(selectedSourceRow ?? selectedRow ?? 0) + 1}`
      : "A1",
  );
  const selectedValue = $derived(spreadsheetState.selectedCellValue);
  const activeSheetSummary = $derived(
    spreadsheetState.sheets[spreadsheetState.activeSheetIndex] ?? null,
  );
  const searchResultCount = $derived(spreadsheetState.searchResults.length);
  const activeSearchMatch = $derived(
    spreadsheetState.searchResults[activeSearchIndex] ?? null,
  );
  const isWorkbookBusy = $derived(
    spreadsheetState.isLoading ||
      spreadsheetState.isApplyingView ||
      spreadsheetState.isSearching,
  );
  const selectedDisplayRowLabel = $derived(
    selectedRow !== null ? (selectedRow + 1).toLocaleString() : "-",
  );
  const selectedSourceRowLabel = $derived(
    selectedSourceRow !== null ? (selectedSourceRow + 1).toLocaleString() : "-",
  );
  const selectedColumnLabel = $derived(
    selectedCol !== null ? excelColumnLabel(selectedCol + 1) : "-",
  );
  const selectedColumnHeader = $derived.by(() => {
    if (selectedCol === null) return "-";
    return (
      activeColumnLabels[selectedCol - activeStartCol] ??
      excelColumnLabel(selectedCol + 1)
    );
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
        const normalized = (cell ?? "").trim();
        if (normalized.length === 0) return;

        nonEmptyCells += 1;
        if (isNumericCell(cell)) numericCells += 1;
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
      activeColumnFilled,
    };
  });

  const activeColumnNumericStats = $derived.by(() => {
    if (selectedCol === null) return null;

    const columnOffset = selectedCol - activeStartCol;
    if (columnOffset < 0 || columnOffset >= activeColLimit) return null;

    const numericValues = activeRows
      .map((row) => row[columnOffset] ?? "")
      .map((value) => {
        const normalized = value.trim();
        if (!isNumericCell(normalized)) return null;
        const parsed = Number.parseFloat(normalized.replace(",", "."));
        return Number.isFinite(parsed) ? parsed : null;
      })
      .filter((value): value is number => value !== null);

    if (numericValues.length === 0) return null;

    const sum = numericValues.reduce(
      (accumulator, value) => accumulator + value,
      0,
    );
    const min = Math.min(...numericValues);
    const max = Math.max(...numericValues);

    return {
      count: numericValues.length,
      sum,
      min,
      max,
      avg: sum / numericValues.length,
    };
  });

  const selectedRowSnapshot = $derived.by(() => {
    if (
      !spreadsheetState.activePacket ||
      selectedRow === null ||
      selectedRow < spreadsheetState.activePacket.start_row ||
      selectedRow >=
        spreadsheetState.activePacket.start_row +
          spreadsheetState.activePacket.row_limit
    ) {
      return [];
    }

    const rowOffset = selectedRow - spreadsheetState.activePacket.start_row;
    const row = spreadsheetState.activePacket.rows[rowOffset] ?? [];
    return row
      .map((value, index) => ({
        label:
          spreadsheetState.activePacket?.headers[index]?.trim() ||
          excelColumnLabel(
            (spreadsheetState.activePacket?.start_col ?? 0) + index + 1,
          ),
        value,
      }))
      .filter((entry) => entry.value.trim().length > 0)
      .slice(0, 8);
  });

  const visibleRowRange = $derived.by(() => {
    if (!spreadsheetState.hasWorkbook || spreadsheetState.totalRows === 0)
      return "0 - 0";
    const start = spreadsheetState.verticalWindow.visibleStartRow + 1;
    const end = Math.max(
      spreadsheetState.verticalWindow.visibleEndRow,
      spreadsheetState.verticalWindow.visibleStartRow + 1,
    );
    return `${start} - ${end}`;
  });

  const visibleColRange = $derived.by(() => {
    if (!spreadsheetState.hasWorkbook || spreadsheetState.totalCols === 0)
      return "0 - 0";
    const start = spreadsheetState.horizontalWindow.visibleStartCol + 1;
    const end = Math.max(
      spreadsheetState.horizontalWindow.visibleEndCol,
      spreadsheetState.horizontalWindow.visibleStartCol + 1,
    );
    return `${excelColumnLabel(start)} - ${excelColumnLabel(end)}`;
  });

  const viewportCoverage = $derived.by(() => {
    const rowCoverage =
      spreadsheetState.totalRows > 0
        ? Math.min(
            100,
            (spreadsheetState.verticalWindow.visibleRowCount /
              spreadsheetState.totalRows) *
              100,
          )
        : 0;
    const colCoverage =
      spreadsheetState.totalCols > 0
        ? Math.min(
            100,
            (spreadsheetState.horizontalWindow.visibleColCount /
              spreadsheetState.totalCols) *
              100,
          )
        : 0;

    return { rowCoverage, colCoverage };
  });

  const activeColumnSamples = $derived.by(() => {
    if (selectedCol === null) return [];

    const columnOffset = selectedCol - activeStartCol;
    if (columnOffset < 0 || columnOffset >= activeColLimit) return [];

    const sampleCounts = new Map<string, number>();

    for (const row of activeRows) {
      const normalized = (row[columnOffset] ?? "").trim();
      if (!normalized) continue;
      sampleCounts.set(normalized, (sampleCounts.get(normalized) ?? 0) + 1);
    }

    return Array.from(sampleCounts.entries())
      .map(([value, count]) => ({ value, count }))
      .sort(
        (left, right) =>
          right.count - left.count || left.value.localeCompare(right.value),
      )
      .slice(0, 8);
  });

  // Yardımcı Fonksiyonlar
  function excelColumnLabel(index: number) {
    let value = index;
    let label = "";
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
  function formatPercent(value: number) {
    return `${value.toFixed(value >= 10 ? 1 : 2)}%`;
  }
  function columnSortMarker(colIndex: number) {
    if (spreadsheetState.sortCol !== colIndex) return "";
    return spreadsheetState.sortDirection === "desc" ? "▼" : "▲";
  }
  function normalizedSearchQuery() {
    return spreadsheetState.searchQuery.trim().toLowerCase();
  }
  function isNumericCell(value: string | undefined) {
    const normalized = (value ?? "").trim();
    return (
      normalized.length > 0 && /^[-+]?(\d+([.,]\d+)?|[.,]\d+)$/.test(normalized)
    );
  }
  function formatCellValue(value: string | undefined) {
    const normalized = value ?? "";
    return normalized.length > 0 ? normalized : " ";
  }
  function isSearchHighlighted(value: string | undefined) {
    const query = normalizedSearchQuery();
    if (!query) return false;
    return (value ?? "").toLowerCase().includes(query);
  }
  function isActiveSearchCell(sourceRowIndex: number, colIndex: number) {
    if (!activeSearchMatch) return false;
    return (
      activeSearchMatch.source_row === sourceRowIndex &&
      activeSearchMatch.col === colIndex
    );
  }
  function focusElement(target: HTMLInputElement | null) {
    target?.focus();
    target?.select();
  }
  function formatMetricValue(value: number) {
    return Number.isInteger(value)
      ? value.toLocaleString()
      : value.toLocaleString(undefined, { maximumFractionDigits: 2 });
  }

  function parseColumnInput(rawValue: string) {
    const normalized = rawValue.trim().toUpperCase();
    if (!normalized) return null;
    if (/^\d+$/.test(normalized)) {
      const numeric = Number.parseInt(normalized, 10);
      return Number.isFinite(numeric) && numeric > 0 ? numeric - 1 : null;
    }
    if (!/^[A-Z]+$/.test(normalized)) return null;
    let index = 0;
    for (const character of normalized) {
      index = index * 26 + (character.charCodeAt(0) - 64);
    }
    return index - 1;
  }

  // --- CANVAS ENGINE (ÇİZİM VE BOYAMA DÖNGÜSÜ) ---
  function drawCanvas() {
    if (
      !canvasElement ||
      !canvasContext ||
      viewportWidth === 0 ||
      viewportHeight === 0
    )
      return;

    const ctx = canvasContext;
    const dpr = window.devicePixelRatio || 1;

    canvasElement.width = viewportWidth * dpr;
    canvasElement.height = viewportHeight * dpr;
    ctx.scale(dpr, dpr);

    ctx.clearRect(0, 0, viewportWidth, viewportHeight);

    if (!spreadsheetState.hasWorkbook) return;

    const {
      scrollTop,
      scrollLeft,
      rowHeight,
      columnWidth,
      rowHeaderWidth,
      selectedRow,
      selectedCol,
      activePacket,
    } = spreadsheetState;

    const headerHeight = rowHeight;
    const startRowIdx = activePacket?.start_row ?? 0;
    const startColIdx = activePacket?.start_col ?? 0;

    // 1. GRID VE HÜCRE İÇERİKLERİ
    ctx.font = "12px 'IBM Plex Mono', Consolas, monospace";
    ctx.textBaseline = "middle";

    activeRows.forEach((row, rIdx) => {
      const displayRow = startRowIdx + rIdx;
      const y = displayRow * rowHeight - scrollTop + headerHeight;

      if (y > viewportHeight || y + rowHeight < headerHeight) return;

      if (selectedRow === displayRow) {
        ctx.fillStyle = "rgba(229, 224, 206, 0.25)";
        ctx.fillRect(
          rowHeaderWidth,
          y,
          viewportWidth - rowHeaderWidth,
          rowHeight,
        );
      }

      row.forEach((cell, cIdx) => {
        const displayCol = startColIdx + cIdx;
        const x = displayCol * columnWidth - scrollLeft + rowHeaderWidth;

        if (x > viewportWidth || x + columnWidth < rowHeaderWidth) return;

        const isSelected =
          selectedRow === displayRow && selectedCol === displayCol;
        const isSearchHit = isSearchHighlighted(cell);
        const currentSourceRow = activeSourceRows[rIdx] ?? displayRow;
        const isActiveSearch = isActiveSearchCell(currentSourceRow, displayCol);

        if (isSelected) {
          ctx.fillStyle = "rgba(223, 218, 197, 0.96)";
          ctx.fillRect(x, y, columnWidth, rowHeight);
        } else if (isActiveSearch) {
          ctx.fillStyle = "rgba(111, 116, 88, 0.22)";
          ctx.fillRect(x, y, columnWidth, rowHeight);
        } else if (isSearchHit) {
          ctx.fillStyle = "rgba(111, 116, 88, 0.08)";
          ctx.fillRect(x, y, columnWidth, rowHeight);
        }

        ctx.strokeStyle = "rgba(211, 206, 186, 0.48)";
        ctx.strokeRect(x, y, columnWidth, rowHeight);

        if (isSelected) {
          ctx.strokeStyle = "#4f563e";
          ctx.lineWidth = 1.5;
          ctx.strokeRect(x, y, columnWidth, rowHeight);
          ctx.lineWidth = 1;
        }

        if (cell !== null && cell !== undefined) {
          ctx.fillStyle = "#24271d";
          ctx.save();
          ctx.beginPath();
          ctx.rect(x + 6, y, columnWidth - 12, rowHeight);
          ctx.clip();

          const normalizedCell = formatCellValue(cell);
          if (isNumericCell(normalizedCell)) {
            ctx.textAlign = "right";
            ctx.fillText(
              normalizedCell,
              x + columnWidth - 6,
              y + rowHeight / 2,
            );
          } else {
            ctx.textAlign = "left";
            ctx.fillText(normalizedCell, x + 6, y + rowHeight / 2);
          }
          ctx.restore();
        }
      });
    });

    // 2. STICKY COLUMN HEADERS
    ctx.fillStyle = "rgba(247, 244, 233, 0.98)";
    ctx.fillRect(
      rowHeaderWidth,
      0,
      viewportWidth - rowHeaderWidth,
      headerHeight,
    );

    ctx.font = "bold 10px 'IBM Plex Mono', Consolas, monospace";
    ctx.textBaseline = "middle";

    activeColumnLabels.forEach((label, cIdx) => {
      const colIndex = startColIdx + cIdx;
      const x = colIndex * columnWidth - scrollLeft + rowHeaderWidth;

      if (x > viewportWidth || x + columnWidth < rowHeaderWidth) return;

      const isSorted = spreadsheetState.sortCol === colIndex;
      ctx.fillStyle = isSorted
        ? "rgba(224, 219, 197, 0.98)"
        : "rgba(237, 233, 217, 0.98)";
      ctx.fillRect(x, 0, columnWidth, headerHeight);

      ctx.strokeStyle = "#c9c3ab";
      ctx.strokeRect(x, 0, columnWidth, headerHeight);

      ctx.fillStyle = isSorted ? "#4f563e" : "#696d58";
      ctx.textAlign = "left";
      ctx.fillText(label.toUpperCase(), x + 8, headerHeight / 2);

      if (isSorted) {
        ctx.textAlign = "right";
        ctx.fillText(
          columnSortMarker(colIndex),
          x + columnWidth - 8,
          headerHeight / 2,
        );
      }
    });

    // 3. STICKY ROW HEADERS
    ctx.fillStyle = "rgba(247, 244, 233, 0.98)";
    ctx.fillRect(
      0,
      headerHeight,
      rowHeaderWidth,
      viewportHeight - headerHeight,
    );
    ctx.textAlign = "center";

    activeRows.forEach((_, rIdx) => {
      const displayRow = startRowIdx + rIdx;
      const y = displayRow * rowHeight - scrollTop + headerHeight;

      if (y > viewportHeight || y + rowHeight < headerHeight) return;

      const isRowSelected = selectedRow === displayRow;
      const sourceRowIndex = activeSourceRows[rIdx] ?? displayRow;

      ctx.fillStyle = isRowSelected
        ? "rgba(219, 213, 188, 0.96)"
        : "rgba(237, 233, 217, 0.98)";
      ctx.fillRect(0, y, rowHeaderWidth, rowHeight);

      ctx.strokeStyle = "#c9c3ab";
      ctx.strokeRect(0, y, rowHeaderWidth, rowHeight);

      ctx.fillStyle = isRowSelected ? "#4f563e" : "#696d58";
      ctx.fillText(
        (sourceRowIndex + 1).toString(),
        rowHeaderWidth / 2,
        y + rowHeight / 2,
      );
    });

    // 4. TOP-LEFT CORNER BLOCK
    ctx.fillStyle = "rgba(237, 233, 217, 1)";
    ctx.fillRect(0, 0, rowHeaderWidth, headerHeight);
    ctx.strokeStyle = "#c9c3ab";
    ctx.strokeRect(0, 0, rowHeaderWidth, headerHeight);
    ctx.fillStyle = "#696d58";
    ctx.textAlign = "center";
    ctx.fillText("ROW", rowHeaderWidth / 2, headerHeight / 2);
  }

  // --- REAKTİVİTE & FARE MATEMATİĞİ ENTEGRASYONU ---
  $effect(() => {
    activeRows;
    spreadsheetState.scrollTop;
    spreadsheetState.scrollLeft;
    selectedRow;
    selectedCol;
    spreadsheetState.searchResults;
    viewportWidth;
    viewportHeight;

    requestAnimationFrame(() => drawCanvas());
  });

  $effect(() => {
    if (canvasElement && !canvasContext) {
      canvasContext = canvasElement.getContext("2d", { alpha: false });
    }
  });

  function resolveEventCoordinates(event: MouseEvent) {
    const rect = canvasElement!.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;

    const { scrollTop, scrollLeft, rowHeight, columnWidth, rowHeaderWidth } =
      spreadsheetState;
    const headerHeight = rowHeight;

    const isCorner = x <= rowHeaderWidth && y <= headerHeight;
    const isColHeader = x > rowHeaderWidth && y <= headerHeight;
    const isRowHeader = x <= rowHeaderWidth && y > headerHeight;
    const isCell = x > rowHeaderWidth && y > headerHeight;

    const clickedDisplayCol =
      isColHeader || isCell
        ? Math.floor((x - rowHeaderWidth + scrollLeft) / columnWidth)
        : null;

    const clickedDisplayRow =
      isRowHeader || isCell
        ? Math.floor((y - headerHeight + scrollTop) / rowHeight)
        : null;

    return {
      isCorner,
      isColHeader,
      isRowHeader,
      isCell,
      clickedDisplayCol,
      clickedDisplayRow,
    };
  }

  function handleCanvasClick(event: MouseEvent) {
    if (!spreadsheetState.hasWorkbook) return;
    const {
      isColHeader,
      isRowHeader,
      isCell,
      clickedDisplayCol,
      clickedDisplayRow,
    } = resolveEventCoordinates(event);

    if (isColHeader && clickedDisplayCol !== null) {
      void sortByColumn(clickedDisplayCol);
    } else if (isRowHeader && clickedDisplayRow !== null) {
      focusCell(
        clickedDisplayRow,
        Math.max(selectedCol ?? 0, 0),
        spreadsheetState.resolveSourceRow(clickedDisplayRow),
      );
    } else if (
      isCell &&
      clickedDisplayRow !== null &&
      clickedDisplayCol !== null
    ) {
      focusCell(
        clickedDisplayRow,
        clickedDisplayCol,
        spreadsheetState.resolveSourceRow(clickedDisplayRow),
      );
    }
  }

  function handleCanvasContextMenu(event: MouseEvent) {
    if (!spreadsheetState.hasWorkbook) return;
    const {
      isColHeader,
      isRowHeader,
      isCell,
      clickedDisplayCol,
      clickedDisplayRow,
    } = resolveEventCoordinates(event);

    if (isColHeader && clickedDisplayCol !== null) {
      const label =
        activeColumnLabels[clickedDisplayCol - activeStartCol] ??
        excelColumnLabel(clickedDisplayCol + 1);
      handleColumnContextMenu(event, clickedDisplayCol, label);
    } else if (isRowHeader && clickedDisplayRow !== null) {
      const sourceRow = spreadsheetState.resolveSourceRow(clickedDisplayRow);
      const rIdx = clickedDisplayRow - activeStartRow;
      handleRowContextMenu(
        event,
        clickedDisplayRow,
        sourceRow ?? clickedDisplayRow,
        activeRows[rIdx] ?? [],
      );
    } else if (
      isCell &&
      clickedDisplayRow !== null &&
      clickedDisplayCol !== null
    ) {
      const sourceRow = spreadsheetState.resolveSourceRow(clickedDisplayRow);
      const rIdx = clickedDisplayRow - activeStartRow;
      const cIdx = clickedDisplayCol - activeStartCol;
      const rowData = activeRows[rIdx] ?? [];
      handleCellContextMenu(
        event,
        clickedDisplayRow,
        clickedDisplayCol,
        sourceRow ?? clickedDisplayRow,
        rowData[cIdx] ?? "",
        rowData,
      );
    } else {
      handleViewportContextMenu(event);
    }
  }

  function syncViewportScroll() {
    if (!viewportElement) return;
    viewportElement.scrollTo({
      top: spreadsheetState.scrollTop,
      left: spreadsheetState.scrollLeft,
      behavior: "auto",
    });
  }

  // İş Akış Kontrolleri (Core Logic)
  async function loadWorkbook() {
    await spreadsheetState.loadExcelFile(filePathInput);
    searchInput = "";
    filterInput = spreadsheetState.filterQuery;
    jumpToRowInput = "";
    activeSearchIndex = 0;
    syncViewportScroll();
  }

  async function selectExcelFile() {
    if (!isDesktopRuntime) {
      spreadsheetState.loadError =
        "File picker requires the Tauri desktop runtime.";
      return;
    }

    try {
      const selected = await open({
        multiple: false,
        directory: false,
        filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
      });

      if (!selected || Array.isArray(selected)) return;

      filePathInput = selected;
      await spreadsheetState.loadExcelFile(selected);
      searchInput = "";
      filterInput = spreadsheetState.filterQuery;
      jumpToRowInput = "";
      activeSearchIndex = 0;
      syncViewportScroll();
    } catch (error) {
      spreadsheetState.loadError =
        error instanceof Error ? error.message : "File picker failed.";
    }
  }

  async function reloadWorkbook() {
    if (!filePathInput.trim()) return;
    await spreadsheetState.loadExcelFile(filePathInput);
    searchInput = "";
    filterInput = spreadsheetState.filterQuery;
    jumpToRowInput = "";
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
    if (!target) return;
    spreadsheetState.setScrollOffsets(target.scrollTop, target.scrollLeft);
  }

  async function handleSheetChange(event: Event) {
    const target = event.currentTarget as HTMLSelectElement | null;
    if (!target) return;
    const sheetIndex = Number.parseInt(target.value, 10);
    if (Number.isNaN(sheetIndex)) return;

    await spreadsheetState.activateSheet(sheetIndex);
    searchInput = "";
    filterInput = spreadsheetState.filterQuery;
    jumpToRowInput = "";
    activeSearchIndex = 0;
    syncViewportScroll();
  }

  async function activateSheetTab(sheetIndex: number) {
    await spreadsheetState.activateSheet(sheetIndex);
    searchInput = "";
    filterInput = spreadsheetState.filterQuery;
    jumpToRowInput = "";
    activeSearchIndex = 0;
    syncViewportScroll();
  }

  async function runSearch() {
    await spreadsheetState.searchWorkbook(searchInput);
    activeSearchIndex = 0;
  }

  async function runSearchForValue(value: string) {
    const normalized = value.trim();
    if (!normalized) return;
    searchInput = normalized;
    await spreadsheetState.searchWorkbook(normalized);
    activeSearchIndex = 0;
  }

  function clearSearch() {
    searchInput = "";
    activeSearchIndex = 0;
    spreadsheetState.clearSearchResults();
  }

  async function applyFilter() {
    await spreadsheetState.applyFilter(filterInput);
    syncViewportScroll();
  }

  async function applyFilterValue(value: string) {
    const normalized = value.trim();
    if (!normalized) return;
    filterInput = normalized;
    await spreadsheetState.applyFilter(normalized);
    syncViewportScroll();
  }

  async function clearViewOptions() {
    filterInput = "";
    await spreadsheetState.clearViewOptions();
    syncViewportScroll();
  }

  async function handleHeaderToggle(event: Event) {
    const target = event.currentTarget as HTMLInputElement | null;
    if (!target) return;
    await spreadsheetState.setHeaderRowEnabled(target.checked);
    syncViewportScroll();
  }

  async function toggleHeaderRowMode() {
    await spreadsheetState.setHeaderRowEnabled(
      !spreadsheetState.headerRowEnabled,
    );
    syncViewportScroll();
  }

  async function sortByColumn(colIndex: number) {
    await spreadsheetState.cycleColumnSort(colIndex);
    syncViewportScroll();
  }

  async function sortByColumnDirection(
    colIndex: number,
    direction: SpreadsheetSortDirection,
  ) {
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
    if (total === 0) return;
    const safeIndex = ((index % total) + total) % total;
    activeSearchIndex = safeIndex;
    const match = spreadsheetState.searchResults[safeIndex];
    if (match) focusMatch(match);
  }

  function focusPreviousSearchResult() {
    focusSearchResultAt(activeSearchIndex - 1);
  }
  function focusNextSearchResult() {
    focusSearchResultAt(activeSearchIndex + 1);
  }

  // Hızlı satır gitme logiği
  function goToVisibleRow() {
    const rowNumber = Number.parseInt(jumpToRowInput, 10);
    if (Number.isNaN(rowNumber)) return;
    spreadsheetState.goToRow(rowNumber);
    syncViewportScroll();
  }

  function goToVisibleColumn() {
    const columnIndex = parseColumnInput(jumpToColumnInput);
    if (columnIndex === null) return;
    const targetRow =
      selectedRow ?? spreadsheetState.verticalWindow.visibleStartRow;
    spreadsheetState.focusCell(
      targetRow,
      columnIndex,
      spreadsheetState.resolveSourceRow(targetRow),
    );
    syncViewportScroll();
  }

  function scrubToRow(event: Event) {
    const target = event.currentTarget as HTMLInputElement | null;
    if (!target) return;
    const rowNumber = Number.parseInt(target.value, 10);
    if (Number.isNaN(rowNumber)) return;
    spreadsheetState.goToRow(rowNumber);
    syncViewportScroll();
  }

  function scrubToColumn(event: Event) {
    const target = event.currentTarget as HTMLInputElement | null;
    if (!target) return;
    const colNumber = Number.parseInt(target.value, 10);
    if (Number.isNaN(colNumber)) return;
    const targetRow =
      selectedRow ?? spreadsheetState.verticalWindow.visibleStartRow;
    spreadsheetState.focusCell(
      targetRow,
      Math.max(0, colNumber - 1),
      spreadsheetState.resolveSourceRow(targetRow),
    );
    syncViewportScroll();
  }

  function focusCell(
    displayRowIndex: number,
    colIndex: number,
    sourceRowIndex: number | null,
  ) {
    spreadsheetState.setSelectedCell(displayRowIndex, colIndex, sourceRowIndex);
  }

  function focusRecentSelection(item: RecentSelection) {
    spreadsheetState.focusCell(item.displayRow, item.colIndex, item.sourceRow);
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

  // Panoya Kopyalama İşlemleri
  async function copyText(value: string) {
    if (!value) return;
    try {
      await navigator.clipboard.writeText(value);
    } catch {
      /* Fallback */
    }
  }

  async function copyVisibleChunk() {
    const lines = activeRows.map((row) =>
      row.map((cell) => cell ?? "").join("\t"),
    );
    await copyText(lines.join("\n"));
  }

  async function copySelectedColumnValues() {
    if (selectedCol === null) return;
    const columnOffset = selectedCol - activeStartCol;
    if (columnOffset < 0 || columnOffset >= activeColLimit) return;
    const values = activeRows.map((row) => row[columnOffset] ?? "");
    await copyText(values.join("\n"));
  }

  async function copySelectionValue(value = selectedValue) {
    await copyText(value);
  }
  async function copyRowValues(row: string[]) {
    await copyText(row.map((cell) => cell ?? "").join("\t"));
  }

  // Yerel Sağ Tık (Context) Menüleri
  async function popupNativeContextMenu(
    event: MouseEvent,
    items: NativeMenuItem[],
  ) {
    if (!isDesktopRuntime) return;
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
    if (!spreadsheetState.hasWorkbook) return;
    void popupNativeContextMenu(event, [
      {
        text: "Reload workbook",
        enabled: Boolean(spreadsheetState.filePath),
        action: () => {
          void reloadWorkbook();
        },
      },
      {
        text: "Reset viewport",
        enabled: spreadsheetState.hasWorkbook,
        action: () => resetViewport(),
      },
      {
        text: "Copy visible chunk",
        enabled: activeRows.length > 0,
        action: () => {
          void copyVisibleChunk();
        },
      },
      {
        text: spreadsheetState.headerRowEnabled
          ? "Disable header row"
          : "Use first row as headers",
        enabled: spreadsheetState.hasWorkbook,
        action: () => {
          void toggleHeaderRowMode();
        },
      },
      {
        text: "Clear filter and sort",
        enabled:
          spreadsheetState.hasWorkbook &&
          (spreadsheetState.activeFilter.length > 0 ||
            spreadsheetState.sortCol !== null),
        action: () => {
          void clearViewOptions();
        },
      },
    ]);
  }

  function handleColumnContextMenu(
    event: MouseEvent,
    colIndex: number,
    label: string,
  ) {
    void popupNativeContextMenu(event, [
      {
        text: `Sort ${label} ascending`,
        action: () => {
          void sortByColumnDirection(colIndex, "asc");
        },
      },
      {
        text: `Sort ${label} descending`,
        action: () => {
          void sortByColumnDirection(colIndex, "desc");
        },
      },
      {
        text: "Clear sort",
        enabled: spreadsheetState.sortCol !== null,
        action: () => {
          void clearColumnSort();
        },
      },
      {
        text: "Copy column label",
        action: () => {
          void copyText(label);
        },
      },
      {
        text: "Copy visible column values",
        action: () => {
          void copyText(
            activeRows
              .map((row) => row[colIndex - activeStartCol] ?? "")
              .join("\n"),
          );
        },
      },
    ]);
  }

  function handleRowContextMenu(
    event: MouseEvent,
    displayRowIndex: number,
    sourceRowIndex: number,
    row: string[],
  ) {
    spreadsheetState.setSelectedCell(
      displayRowIndex,
      Math.max(selectedCol ?? 0, 0),
      sourceRowIndex,
    );
    void popupNativeContextMenu(event, [
      {
        text: `Copy row ${sourceRowIndex + 1}`,
        action: () => {
          void copyRowValues(row);
        },
      },
      {
        text: `Jump to row ${sourceRowIndex + 1}`,
        action: () => {
          spreadsheetState.goToRow(displayRowIndex + 1);
          syncViewportScroll();
        },
      },
      { text: "Reset viewport", action: () => resetViewport() },
      {
        text: "Copy visible chunk",
        action: () => {
          void copyVisibleChunk();
        },
      },
    ]);
  }

  function handleCellContextMenu(
    event: MouseEvent,
    displayRowIndex: number,
    colIndex: number,
    sourceRowIndex: number,
    cellValue: string,
    row: string[],
  ) {
    focusCell(displayRowIndex, colIndex, sourceRowIndex);
    const address = `${excelColumnLabel(colIndex + 1)}${sourceRowIndex + 1}`;
    const normalizedValue = (cellValue ?? "").trim();

    void popupNativeContextMenu(event, [
      {
        text: `Copy ${address}`,
        enabled: normalizedValue.length > 0,
        action: () => {
          void copySelectionValue(cellValue);
        },
      },
      {
        text: `Copy row ${sourceRowIndex + 1}`,
        action: () => {
          void copyRowValues(row);
        },
      },
      {
        text: `Copy visible ${excelColumnLabel(colIndex + 1)} values`,
        action: () => {
          spreadsheetState.setSelectedCell(
            displayRowIndex,
            colIndex,
            sourceRowIndex,
          );
          void copySelectedColumnValues();
        },
      },
      {
        text: `Find "${normalizedValue.slice(0, 28) || "value"}"`,
        enabled: normalizedValue.length > 0,
        action: () => {
          void runSearchForValue(normalizedValue);
        },
      },
      {
        text: `Filter by "${normalizedValue.slice(0, 28) || "value"}"`,
        enabled: normalizedValue.length > 0,
        action: () => {
          void applyFilterValue(normalizedValue);
        },
      },
      {
        text: `Sort ${excelColumnLabel(colIndex + 1)} ascending`,
        action: () => {
          void sortByColumnDirection(colIndex, "asc");
        },
      },
      {
        text: `Sort ${excelColumnLabel(colIndex + 1)} descending`,
        action: () => {
          void sortByColumnDirection(colIndex, "desc");
        },
      },
      {
        text: "Clear sort",
        enabled: spreadsheetState.sortCol !== null,
        action: () => {
          void clearColumnSort();
        },
      },
    ]);
  }

  // Klavye Navigasyon Yakalamaları
  function handleViewportKeydown(event: KeyboardEvent) {
    if (!spreadsheetState.hasWorkbook) return;

    if (
      (event.ctrlKey || event.metaKey) &&
      event.shiftKey &&
      event.key.toLowerCase() === "f"
    ) {
      event.preventDefault();
      focusElement(filterInputElement);
      return;
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "f") {
      event.preventDefault();
      focusElement(searchInputElement);
      return;
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "g") {
      event.preventDefault();
      focusElement(jumpToRowInputElement);
      return;
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "l") {
      event.preventDefault();
      focusElement(filePathInputElement);
      return;
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
      event.preventDefault();
      void copySelectionValue();
      return;
    }
    if (event.key === "F3") {
      event.preventDefault();
      if (event.shiftKey) focusPreviousSearchResult();
      else focusNextSearchResult();
      return;
    }

    switch (event.key) {
      case "ArrowUp":
        event.preventDefault();
        spreadsheetState.moveSelection(-1, 0);
        syncViewportScroll();
        break;
      case "ArrowDown":
        event.preventDefault();
        spreadsheetState.moveSelection(1, 0);
        syncViewportScroll();
        break;
      case "ArrowLeft":
        event.preventDefault();
        spreadsheetState.moveSelection(0, -1);
        syncViewportScroll();
        break;
      case "ArrowRight":
        event.preventDefault();
        spreadsheetState.moveSelection(0, 1);
        syncViewportScroll();
        break;
      case "Enter":
        event.preventDefault();
        spreadsheetState.moveSelection(event.shiftKey ? -1 : 1, 0);
        syncViewportScroll();
        break;
      case "Tab":
        event.preventDefault();
        spreadsheetState.moveSelection(0, event.shiftKey ? -1 : 1);
        syncViewportScroll();
        break;
      case "PageUp":
        event.preventDefault();
        spreadsheetState.pageSelection("up");
        syncViewportScroll();
        break;
      case "PageDown":
        event.preventDefault();
        spreadsheetState.pageSelection("down");
        syncViewportScroll();
        break;
      case "Home":
        event.preventDefault();
        spreadsheetState.moveSelectionToEdge(
          event.ctrlKey ? "grid-start" : "row-start",
        );
        syncViewportScroll();
        break;
      case "End":
        event.preventDefault();
        spreadsheetState.moveSelectionToEdge(
          event.ctrlKey ? "grid-end" : "row-end",
        );
        syncViewportScroll();
        break;
    }
  }

  // Senkronizasyon Efektleri ($effect)
  $effect(() => {
    spreadsheetState.setViewportMetrics(viewportHeight, viewportWidth);
  });
  $effect(() => {
    if (
      spreadsheetState.filePath &&
      spreadsheetState.filePath !== filePathInput
    )
      filePathInput = spreadsheetState.filePath;
  });
  $effect(() => {
    if (selectedCol !== null)
      jumpToColumnInput = excelColumnLabel(selectedCol + 1);
  });
  $effect(() => {
    filterInput = spreadsheetState.filterQuery;
  });

  // Hatalı satır düzeltildi (Sözdizimi hatası giderildi)
  $effect(() => {
    if (
      spreadsheetState.searchResults.length > 0 &&
      activeSearchIndex >= spreadsheetState.searchResults.length
    ) {
      activeSearchIndex = 0;
    }
  });

  $effect(() => {
    if (
      !spreadsheetState.hasWorkbook ||
      selectedRow === null ||
      selectedCol === null
    )
      return;
    const entry: RecentSelection = {
      key: `${spreadsheetState.sheetName}:${selectedAddress}`,
      address: selectedAddress,
      sheetName: spreadsheetState.sheetName || "Sheet",
      displayRow: selectedRow,
      sourceRow: selectedSourceRow,
      colIndex: selectedCol,
      value: selectedValue,
    };
    recentSelections = [
      entry,
      ...recentSelections.filter((item) => item.key !== entry.key),
    ].slice(0, 10);
  });
</script>

<section
  class="sheets-module -m-4 flex h-[calc(100%+2rem)] w-[calc(100%+2rem)] min-h-0 flex-col overflow-hidden border"
>
  {#if !isDesktopRuntime}
    <div class="sheets-runtime-banner px-2 py-1 text-[11px]">
      Excel loading only works inside the native Tauri desktop window opened by
      `python run.py`.
    </div>
  {/if}

  <div class="sheets-toolbar shrink-0 border-b">
    <div class="flex flex-wrap items-center gap-1 border-b px-2 py-1">
      <span class="sheets-chip sheets-chip-strong">SHEETS</span>
      <span
        class="sheets-chip font-mono"
        title={spreadsheetState.fileName || filePathInput}
      >
        {spreadsheetState.fileName || "NO FILE"}
      </span>
      <span class="sheets-chip">{spreadsheetState.sheetName || "NO SHEET"}</span
      >
      <span class="sheets-chip"
        >Rows {toHumanCount(spreadsheetState.totalRows)}</span
      >
      <span class="sheets-chip"
        >Source {toHumanCount(spreadsheetState.sourceTotalRows)}</span
      >
      <span class="sheets-chip"
        >Cols {toHumanCount(spreadsheetState.totalCols)}</span
      >
      <span class="sheets-chip">View {visibleRowRange}</span>
      <span class="sheets-chip">Cols {visibleColRange}</span>
      <span class="sheets-chip">{spreadsheetState.statusLabel}</span>
      <span class="sheets-chip">{spreadsheetState.densityMode}</span>
      <span class="sheets-chip">{spreadsheetState.columnProfile}</span>
      {#if spreadsheetState.loadError}
        <span class="sheets-chip sheets-chip-danger"
          >{spreadsheetState.loadError}</span
        >
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
        class="sheets-button-primary h-8 px-3 text-xs font-semibold disabled:opacity-50"
        onclick={() => void loadWorkbook()}
        disabled={!isDesktopRuntime}>Load</button
      >
      <button
        class="sheets-button h-8 px-3 text-xs disabled:opacity-50"
        onclick={selectExcelFile}
        disabled={!isDesktopRuntime}>Select File</button
      >
      <button
        class="sheets-button h-8 px-3 text-xs disabled:opacity-50"
        onclick={() => void reloadWorkbook()}
        disabled={!isDesktopRuntime || !spreadsheetState.filePath}
        >Reload</button
      >
      <button
        class="sheets-button h-8 px-3 text-xs disabled:opacity-50"
        onclick={resetViewport}
        disabled={!spreadsheetState.hasWorkbook}>Reset View</button
      >
      <button
        class="sheets-button h-8 px-3 text-xs disabled:opacity-50"
        onclick={toggleInspectorPanel}
        disabled={!spreadsheetState.hasWorkbook}
        >{showInspector ? "Hide Panel" : "Show Panel"}</button
      >

      <select
        class="sheets-input h-8 min-w-[220px] max-w-[320px] font-mono text-xs disabled:opacity-50"
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
      <div class="sheets-name-box h-8 w-24 font-mono text-xs font-semibold">
        {selectedAddress}
      </div>
      <div
        class="sheets-formula-box flex h-8 min-w-[260px] flex-1 items-center gap-2 px-2"
      >
        <span class="sheets-fx">fx</span>
        <span class="min-w-0 flex-1 truncate font-mono text-xs">
          {selectedValue || "Select a cell to inspect or copy its value."}
        </span>
      </div>
      <button
        class="sheets-button h-8 px-3 text-xs disabled:opacity-50"
        onclick={() => void copySelectionValue()}
        disabled={!selectedValue}
        type="button">Copy Cell</button
      >
      <label
        class="inline-flex h-8 items-center gap-2 px-2 text-[11px] text-[var(--sheet-muted)]"
      >
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
          onkeydown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              void runSearch();
            }
          }}
        />
        <button
          class="sheets-button h-8 px-3 text-xs disabled:opacity-50"
          onclick={() => void runSearch()}
          disabled={!spreadsheetState.hasWorkbook ||
            spreadsheetState.isSearching}
          >{spreadsheetState.isSearching ? "Searching" : "Find"}</button
        >
        <button
          class="sheets-button h-8 px-3 text-xs disabled:opacity-50"
          onclick={clearSearch}
          disabled={searchResultCount === 0 && searchInput.trim().length === 0}
          >Clear</button
        >
        <button
          class="sheets-button h-8 px-3 text-xs disabled:opacity-50"
          onclick={focusPreviousSearchResult}
          disabled={searchResultCount === 0}>Prev</button
        >
        <button
          class="sheets-button h-8 px-3 text-xs disabled:opacity-50"
          onclick={focusNextSearchResult}
          disabled={searchResultCount === 0}>Next</button
        >
      </div>

      <div class="flex min-w-[260px] flex-[1.1] items-center gap-1">
        <input
          bind:this={filterInputElement}
          bind:value={filterInput}
          class="sheets-input h-8 min-w-0 flex-1 font-mono text-xs"
          placeholder="Filter rows by text"
          spellcheck="false"
          disabled={!spreadsheetState.hasWorkbook}
          onkeydown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              void applyFilter();
            }
          }}
        />
        <button
          class="sheets-button h-8 px-3 text-xs disabled:opacity-50"
          onclick={() => void applyFilter()}
          disabled={!spreadsheetState.hasWorkbook ||
            spreadsheetState.isApplyingView}>Apply</button
        >
        <button
          class="sheets-button h-8 px-3 text-xs disabled:opacity-50"
          onclick={() => void clearViewOptions()}
          disabled={!spreadsheetState.hasWorkbook}>Clear</button
        >
      </div>

      <div class="flex min-w-[150px] items-center gap-1">
        <input
          bind:this={jumpToRowInputElement}
          bind:value={jumpToRowInput}
          class="sheets-input h-8 w-24 font-mono text-xs"
          placeholder="Row #"
          disabled={!spreadsheetState.hasWorkbook}
          onkeydown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              goToVisibleRow();
            }
          }}
        />
        <button
          class="sheets-button h-8 px-3 text-xs disabled:opacity-50"
          onclick={goToVisibleRow}
          disabled={!spreadsheetState.hasWorkbook}>Go</button
        >
      </div>

      <div class="flex min-w-[170px] items-center gap-1">
        <input
          bind:value={jumpToColumnInput}
          class="sheets-input h-8 w-24 font-mono text-xs uppercase"
          placeholder="Col / AA"
          disabled={!spreadsheetState.hasWorkbook}
          onkeydown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              goToVisibleColumn();
            }
          }}
        />
        <button
          class="sheets-button h-8 px-3 text-xs disabled:opacity-50"
          onclick={goToVisibleColumn}
          disabled={!spreadsheetState.hasWorkbook}>Column</button
        >
      </div>

      <div class="flex items-center gap-1">
        <button
          class="sheets-button h-8 px-3 text-xs disabled:opacity-50"
          onclick={goToFirstRow}
          disabled={!spreadsheetState.hasWorkbook}>Top</button
        >
        <button
          class="sheets-button h-8 px-3 text-xs disabled:opacity-50"
          onclick={goToLastRow}
          disabled={!spreadsheetState.hasWorkbook}>Bottom</button
        >
        <button
          class="sheets-button h-8 px-3 text-xs disabled:opacity-50"
          onclick={() => void copyVisibleChunk()}
          disabled={activeRows.length === 0}>Copy View</button
        >
      </div>
    </div>
  </div>

  {#if spreadsheetState.hasWorkbook}
    <div class="flex min-h-0 flex-1">
      {#if showInspector}
        <aside
          class="sheets-sidebar enterprise-scrollbar hidden w-[21rem] shrink-0 overflow-auto border-r xl:block"
        >
          <div class="sheets-panel-section border-b px-3 py-3">
            <div class="flex items-center justify-between gap-2">
              <div>
                <p class="sheets-section-label">Workbook</p>
                <h2 class="mt-1 text-sm font-semibold text-[var(--sheet-text)]">
                  {spreadsheetState.sheetName || "Active Sheet"}
                </h2>
              </div>
              <span class="sheets-panel-badge"
                >{isWorkbookBusy ? "BUSY" : "READY"}</span
              >
            </div>
            <div class="mt-3 grid grid-cols-2 gap-2">
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Visible rows</span><span
                  class="sheets-metric-value"
                  >{toHumanCount(spreadsheetState.totalRows)}</span
                >
              </div>
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Source rows</span><span
                  class="sheets-metric-value"
                  >{toHumanCount(spreadsheetState.sourceTotalRows)}</span
                >
              </div>
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Columns</span><span
                  class="sheets-metric-value"
                  >{toHumanCount(spreadsheetState.totalCols)}</span
                >
              </div>
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Viewport</span><span
                  class="sheets-metric-value">{visibleRowRange}</span
                >
              </div>
            </div>
            {#if activeSheetSummary}
              <div
                class="mt-3 rounded-sm border border-[var(--sheet-border)] bg-white/72 px-3 py-2 text-[11px] text-[var(--sheet-muted)]"
              >
                <div class="flex items-center justify-between gap-2">
                  <span class="font-semibold text-[var(--sheet-text)]"
                    >{activeSheetSummary.name}</span
                  >
                  <span class="font-mono"
                    >{activeSheetSummary.total_rows.toLocaleString()} x {activeSheetSummary.total_cols.toLocaleString()}</span
                  >
                </div>
              </div>
            {/if}
          </div>

          <div class="sheets-panel-section border-b px-3 py-3">
            <p class="sheets-section-label">Presentation</p>
            <div class="mt-3">
              <div class="flex items-center justify-between gap-2">
                <span class="text-[11px] font-medium text-[var(--sheet-muted)]"
                  >Density</span
                >
                <span class="font-mono text-[11px] text-[var(--sheet-text)]"
                  >{spreadsheetState.densityMode}</span
                >
              </div>
              <div class="mt-2 grid grid-cols-3 gap-1">
                {#each ["compact", "balanced", "comfortable"] as mode}
                  <button
                    class={`sheets-toggle-button ${spreadsheetState.densityMode === mode ? "sheets-toggle-button-active" : ""}`}
                    onclick={() =>
                      setDensityMode(mode as SpreadsheetDensityMode)}
                    type="button">{mode}</button
                  >
                {/each}
              </div>
            </div>

            <div class="mt-3">
              <div class="flex items-center justify-between gap-2">
                <span class="text-[11px] font-medium text-[var(--sheet-muted)]"
                  >Column profile</span
                >
                <span class="font-mono text-[11px] text-[var(--sheet-text)]"
                  >{spreadsheetState.columnProfile}</span
                >
              </div>
              <div class="mt-2 grid grid-cols-3 gap-1">
                {#each ["narrow", "standard", "wide"] as profile}
                  <button
                    class={`sheets-toggle-button ${spreadsheetState.columnProfile === profile ? "sheets-toggle-button-active" : ""}`}
                    onclick={() =>
                      setColumnProfile(profile as SpreadsheetColumnProfile)}
                    type="button">{profile}</button
                  >
                {/each}
              </div>
            </div>
          </div>

          <div class="sheets-panel-section border-b px-3 py-3">
            <div class="flex items-center justify-between gap-2">
              <p class="sheets-section-label">Navigator</p>
              <span class="font-mono text-[11px] text-[var(--sheet-muted)]"
                >{selectedAddress}</span
              >
            </div>
            <div
              class="mt-3 rounded-sm border border-[var(--sheet-border)] bg-white/78 p-3"
            >
              <div class="flex items-center justify-between gap-2">
                <span class="text-[11px] font-medium text-[var(--sheet-muted)]"
                  >Row sweep</span
                >
                <span class="font-mono text-[11px] text-[var(--sheet-text)]"
                  >{selectedDisplayRowLabel}</span
                >
              </div>
              <input
                class="sheets-range mt-3 w-full"
                type="range"
                min="1"
                max={Math.max(1, spreadsheetState.totalRows)}
                value={(selectedRow ??
                  spreadsheetState.verticalWindow.visibleStartRow) + 1}
                oninput={scrubToRow}
              />
              <div
                class="mt-2 flex items-center justify-between text-[10px] tracking-[0.12em] text-[var(--sheet-muted)]"
              >
                <span>1</span><span
                  >{toHumanCount(spreadsheetState.totalRows)}</span
                >
              </div>
            </div>

            <div
              class="mt-3 rounded-sm border border-[var(--sheet-border)] bg-white/78 p-3"
            >
              <div class="flex items-center justify-between gap-2">
                <span class="text-[11px] font-medium text-[var(--sheet-muted)]"
                  >Column sweep</span
                >
                <span class="font-mono text-[11px] text-[var(--sheet-text)]"
                  >{selectedColumnLabel}</span
                >
              </div>
              <input
                class="sheets-range mt-3 w-full"
                type="range"
                min="1"
                max={Math.max(1, spreadsheetState.totalCols)}
                value={(selectedCol ??
                  spreadsheetState.horizontalWindow.visibleStartCol) + 1}
                oninput={scrubToColumn}
              />
              <div
                class="mt-2 flex items-center justify-between text-[10px] tracking-[0.12em] text-[var(--sheet-muted)]"
              >
                <span>A</span><span
                  >{excelColumnLabel(
                    Math.max(1, spreadsheetState.totalCols),
                  )}</span
                >
              </div>
            </div>
          </div>

          <div class="sheets-panel-section border-b px-3 py-3">
            <p class="sheets-section-label">Visible Chunk Data</p>
            <div class="mt-3 grid grid-cols-2 gap-2">
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Rows</span><span
                  class="sheets-metric-value">{visibleChunkStats.rowCount}</span
                >
              </div>
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Cells</span><span
                  class="sheets-metric-value"
                  >{visibleChunkStats.cellCount}</span
                >
              </div>
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Filled</span><span
                  class="sheets-metric-value"
                  >{visibleChunkStats.nonEmptyCells}</span
                >
              </div>
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Numeric</span><span
                  class="sheets-metric-value"
                  >{visibleChunkStats.numericCells}</span
                >
              </div>
            </div>
          </div>

          <div class="sheets-panel-section border-b px-3 py-3">
            <p class="sheets-section-label">
              Active Column Stats ({selectedColumnHeader})
            </p>
            {#if activeColumnNumericStats}
              <div class="mt-3 grid grid-cols-2 gap-2">
                <div class="sheets-metric-card">
                  <span class="sheets-metric-label">Count</span><span
                    class="sheets-metric-value"
                    >{activeColumnNumericStats.count}</span
                  >
                </div>
                <div class="sheets-metric-card">
                  <span class="sheets-metric-label">Average</span><span
                    class="sheets-metric-value"
                    >{formatMetricValue(activeColumnNumericStats.avg)}</span
                  >
                </div>
                <div class="sheets-metric-card">
                  <span class="sheets-metric-label">Min</span><span
                    class="sheets-metric-value"
                    >{formatMetricValue(activeColumnNumericStats.min)}</span
                  >
                </div>
                <div class="sheets-metric-card">
                  <span class="sheets-metric-label">Max</span><span
                    class="sheets-metric-value"
                    >{formatMetricValue(activeColumnNumericStats.max)}</span
                  >
                </div>
              </div>
              <div
                class="mt-3 rounded-sm border border-[var(--sheet-border)] bg-white/72 px-3 py-2 text-[11px]"
              >
                <div class="flex items-center justify-between">
                  <span>Sum</span><span
                    class="font-mono font-bold text-[var(--sheet-text)]"
                    >{formatMetricValue(activeColumnNumericStats.sum)}</span
                  >
                </div>
              </div>
            {:else}
              <p class="mt-2 text-[11px] text-[var(--sheet-muted)]">
                No active numeric dataset in chunk viewport.
              </p>
            {/if}
          </div>

          <div class="sheets-panel-section border-b px-3 py-3">
            <p class="sheets-section-label">Distinct Value Hits</p>
            {#if activeColumnSamples.length > 0}
              <div class="mt-3 space-y-1.5">
                {#each activeColumnSamples as sample}
                  <div
                    class="rounded-sm border border-[var(--sheet-border)] bg-white/78 px-3 py-1.5 flex items-center justify-between"
                  >
                    <span
                      class="truncate font-mono text-[11px] font-semibold text-[var(--sheet-text)]"
                      >{sample.value}</span
                    >
                    <span
                      class="font-mono text-[10px] text-[var(--sheet-muted)]"
                      >{sample.count} hit</span
                    >
                  </div>
                {/each}
              </div>
            {:else}
              <p class="mt-2 text-[11px] text-[var(--sheet-muted)]">
                Samples populate upon active row/col context.
              </p>
            {/if}
          </div>

          <div class="sheets-panel-section border-b px-3 py-3">
            <div class="flex items-center justify-between">
              <p class="sheets-section-label">Search Matches</p>
              <span class="font-mono text-[11px]"
                >{searchResultCount > 0
                  ? `${activeSearchIndex + 1}/${searchResultCount}`
                  : "0/0"}</span
              >
            </div>
            {#if activeSearchMatch}
              <button
                class="mt-3 w-full rounded-sm border border-[var(--sheet-accent)] bg-[var(--sheet-accent-soft)] p-2 text-left"
                onclick={() => focusMatch(activeSearchMatch)}
                type="button"
              >
                <span
                  class="font-mono text-xs font-bold text-[var(--sheet-accent-strong)]"
                  >{excelColumnLabel(
                    activeSearchMatch.col + 1,
                  )}{activeSearchMatch.source_row + 1}</span
                >
                <p
                  class="mt-0.5 truncate font-mono text-[11px] text-[var(--sheet-text)]"
                >
                  {activeSearchMatch.value}
                </p>
              </button>
            {/if}
          </div>

          <div class="sheets-panel-section px-3 py-3">
            <p class="sheets-section-label">Cell Preview</p>
            <div
              class="mt-2 rounded-sm border border-[var(--sheet-border)] bg-white/78 p-3 max-h-32 overflow-auto font-mono text-xs whitespace-pre-wrap break-all"
            >
              {selectedValue || "(empty cell)"}
            </div>
          </div>
        </aside>
      {/if}

      <div class="flex min-h-0 min-w-0 flex-1 flex-col">
        <div
          class="sheets-tab-strip enterprise-scrollbar shrink-0 overflow-x-auto border-b px-1 py-1"
        >
          <div class="flex min-w-max items-center gap-1">
            {#each spreadsheetState.sheets as sheet, index}
              <button
                class={`sheets-tab h-8 px-3 text-left ${spreadsheetState.activeSheetIndex === index ? "sheets-tab-active" : ""}`}
                onclick={() => void activateSheetTab(index)}
                type="button"
                title={`${sheet.name} (${sheet.total_rows.toLocaleString()} x {sheet.total_cols.toLocaleString()})`}
              >
                <span class="block truncate text-xs font-semibold"
                  >{sheet.name}</span
                >
                <span
                  class="block font-mono text-[10px] text-[var(--sheet-muted)]"
                  >{sheet.total_rows.toLocaleString()} x {sheet.total_cols.toLocaleString()}</span
                >
              </button>
            {/each}
          </div>
        </div>

        <div
          class="sheets-statusline flex flex-wrap items-center justify-between gap-2 border-b px-2 py-1 text-[11px]"
        >
          <div class="flex flex-wrap items-center gap-x-3 gap-y-1">
            <span
              >Filter: <b class="text-[var(--sheet-text)]"
                >{spreadsheetState.activeFilter || "none"}</b
              ></span
            >
            <span
              >Sort: <b class="text-[var(--sheet-text)]"
                >{spreadsheetState.sortCol !== null
                  ? `${excelColumnLabel(spreadsheetState.sortCol + 1)} ${spreadsheetState.sortDirection}`
                  : "none"}</b
              ></span
            >
            <span
              >Keys: Arrows / Tab / Enter / PgUp / PgDn / F3 / Ctrl+F / Ctrl+G</span
            >
          </div>
          {#if isWorkbookBusy}
            <span
              class="font-semibold text-[var(--sheet-accent-strong)] animate-pulse"
              >Processing GPU Repaint...</span
            >
          {/if}
        </div>

        <div
          class="flex min-h-0 flex-1 flex-col relative bg-[var(--sheet-surface)]"
        >
          <div
            bind:this={viewportElement}
            bind:clientHeight={viewportHeight}
            bind:clientWidth={viewportWidth}
            class="sheets-grid enterprise-scrollbar absolute inset-0 overflow-auto outline-none"
            onscroll={handleScroll}
            onkeydown={handleViewportKeydown}
            role="grid"
            tabindex="0"
          >
            <div
              style={`width: ${spreadsheetState.totalContentWidth + spreadsheetState.rowHeaderWidth}px; 
                      height: ${spreadsheetState.totalContentHeight + spreadsheetState.rowHeight}px;`}
              aria-hidden="true"
            ></div>

            <canvas
              bind:this={canvasElement}
              class="absolute left-0 top-0 pointer-events-auto"
              style={`transform: translate(${spreadsheetState.scrollLeft}px, ${spreadsheetState.scrollTop}px); z-index: 10;`}
              onclick={handleCanvasClick}
              oncontextmenu={handleCanvasContextMenu}
            ></canvas>
          </div>
        </div>

        <div
          class="sheets-footer-bar flex items-center justify-between gap-2 border-t px-2 py-1 text-[11px]"
        >
          <div class="flex items-center gap-3 font-mono">
            <span>Focus: {selectedAddress}</span>
            <span>Cache Load: {spreadsheetState.cacheEntryCount}/8 Packets</span
            >
          </div>
          <div class="flex items-center gap-1">
            <span class="text-[var(--sheet-muted)] mr-1">Density:</span>
            {#each ["compact", "balanced", "comfortable"] as mode}
              <button
                class={`sheets-inline-toggle ${spreadsheetState.densityMode === mode ? "sheets-inline-toggle-active" : ""}`}
                onclick={() => setDensityMode(mode as SpreadsheetDensityMode)}
                type="button">{mode}</button
              >
            {/each}
          </div>
        </div>
      </div>
    </div>
  {:else}
    <div class="flex min-h-0 flex-1 items-center justify-center p-4">
      <div class="sheets-empty max-w-4xl border px-6 py-6">
        <p
          class="text-[10px] font-semibold uppercase tracking-[0.22em] text-[var(--sheet-muted)]"
        >
          Spreadsheet Engine
        </p>
        <h2
          class="mt-2 text-xl font-semibold tracking-tight text-[var(--sheet-text)]"
        >
          FLEXBOX High-Performance Canvas Grid
        </h2>
        <p class="mt-3 max-w-2xl text-sm leading-6 text-[var(--sheet-muted)]">
          FLEXBOX keeps workbook data in Rust memory stack and paints the active
          viewport inside an isolated HTML5 Canvas loop. Sorting, filtering and
          query transforms stay completely back-end bound so millions of data
          rows scroll at 60 FPS.
        </p>
        <div
          class="mt-4 rounded-sm border border-[var(--sheet-border)] bg-[var(--sheet-surface-strong)] px-3 py-2 font-mono text-xs text-[var(--sheet-muted)]"
        >
          Example absolute system path: C:\data\production_logs.xlsx
        </div>
        <div class="mt-5 flex gap-2">
          <button
            class="sheets-button-primary px-4 py-2 text-sm font-semibold disabled:opacity-50"
            onclick={() => void loadWorkbook()}
            disabled={!isDesktopRuntime}>Load explicit context</button
          >
          <button
            class="sheets-button px-4 py-2 text-sm disabled:opacity-50"
            onclick={selectExcelFile}
            disabled={!isDesktopRuntime}>Select OS path picker</button
          >
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
    background: radial-gradient(
        circle at top left,
        rgba(255, 255, 255, 0.58),
        transparent 26%
      ),
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
    background: rgba(244, 241, 229, 0.94);
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
    font-weight: 700;
  }
  .sheets-chip-danger {
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
  .sheets-toggle-button,
  .sheets-inline-toggle {
    border-radius: 0.22rem;
    border: 1px solid var(--sheet-border);
  }

  .sheets-input,
  .sheets-name-box,
  .sheets-formula-box {
    background: rgba(255, 255, 255, 0.92);
    color: var(--sheet-text);
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
  .sheets-toggle-button,
  .sheets-inline-toggle {
    background: linear-gradient(
      180deg,
      rgba(255, 255, 255, 0.9),
      rgba(244, 240, 225, 0.9)
    );
    color: var(--sheet-text);
  }
  .sheets-button:hover,
  .sheets-tab:hover,
  .sheets-toggle-button:hover,
  .sheets-inline-toggle:hover {
    background: rgba(255, 255, 255, 0.98);
  }
  .sheets-button-primary {
    background: linear-gradient(180deg, #707656 0%, #565d41 100%);
    border-color: #51573e;
    color: white;
  }

  .sheets-name-box {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    color: var(--sheet-accent-strong);
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
  }

  .sheets-tab {
    min-width: 8rem;
    background: rgba(236, 232, 214, 0.72);
    color: var(--sheet-text);
  }
  .sheets-tab-active {
    background: rgba(255, 255, 255, 0.96);
    border-color: var(--sheet-accent);
    box-shadow: inset 0 2px 0 var(--sheet-accent);
  }

  .sheets-statusline,
  .sheets-footer-bar {
    color: var(--sheet-muted);
  }
  .sheets-grid {
    scrollbar-color: #b3ad90 transparent;
  }

  .sheets-metric-card,
  .sheets-empty-card {
    display: flex;
    min-height: 4rem;
    flex-direction: column;
    justify-content: space-between;
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
    font-family: monospace;
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
  .sheets-inline-toggle-active {
    border-color: var(--sheet-accent);
    background: var(--sheet-accent-soft);
    color: var(--sheet-accent-strong);
  }

  .sheets-empty {
    border-color: var(--sheet-border);
    background: rgba(250, 248, 239, 0.96);
    box-shadow: 0 20px 42px rgba(79, 86, 62, 0.08);
  }
</style>
