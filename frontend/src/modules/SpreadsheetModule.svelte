<script lang="ts">
  import { isTauri } from "@tauri-apps/api/core";
  import { Menu } from "@tauri-apps/api/menu";
  import { open } from "@tauri-apps/plugin-dialog";
  import { PhysicalPosition } from "@tauri-apps/api/window";
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
        if (selectedCol !== null && activeStartCol + index === selectedCol)
          activeColumnFilled += 1;
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
    const sum = numericValues.reduce((acc, val) => acc + val, 0);
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
    return `${spreadsheetState.verticalWindow.visibleStartRow + 1} - ${spreadsheetState.verticalWindow.visibleEndRow}`;
  });

  const visibleColRange = $derived.by(() => {
    if (!spreadsheetState.hasWorkbook || spreadsheetState.totalCols === 0)
      return "0 - 0";
    return `${excelColumnLabel(spreadsheetState.horizontalWindow.visibleStartCol + 1)} - ${excelColumnLabel(spreadsheetState.horizontalWindow.visibleEndCol)}`;
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
      .sort((a, b) => b.count - a.count)
      .slice(0, 8);
  });

  // Core Helpers
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
    return `${value.toFixed(2)}%`;
  }
  function columnSortMarker(colIndex: number) {
    if (spreadsheetState.sortCol !== colIndex) return "";
    return spreadsheetState.sortDirection === "desc" ? " ▼" : " ▲";
  }
  function isNumericCell(value: string | undefined) {
    const normalized = (value ?? "").trim();
    return (
      normalized.length > 0 && /^[-+]?(\d+([.,]\d+)?|[.,]\d+)$/.test(normalized)
    );
  }
  function formatCellValue(value: string | undefined) {
    return value ?? " ";
  }
  function isSearchHighlighted(value: string | undefined) {
    const query = spreadsheetState.searchQuery.trim().toLowerCase();
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
    return value.toLocaleString(undefined, { maximumFractionDigits: 2 });
  }

  function parseColumnInput(rawValue: string) {
    const normalized = rawValue.trim().toUpperCase();
    if (!normalized) return null;
    if (/^\d+$/.test(normalized))
      return Math.max(0, Number.parseInt(normalized, 10) - 1);
    let index = 0;
    for (let i = 0; i < normalized.length; i++) {
      index = index * 26 + (normalized.charCodeAt(i) - 64);
    }
    return index - 1;
  }

  // --- EMİR KİPİ ÇİZİM KONTROLÜ ---
  let renderScheduled = false;
  function requestRender() {
    if (renderScheduled) return;
    renderScheduled = true;
    requestAnimationFrame(() => {
      renderScheduled = false;
      drawCanvas();
    });
  }

  function focusCell(
    displayRowIndex: number,
    colIndex: number,
    sourceRowIndex: number | null,
  ) {
    spreadsheetState.setSelectedCell(displayRowIndex, colIndex, sourceRowIndex);
    requestRender();
  }

  async function sortByColumn(colIndex: number) {
    await spreadsheetState.cycleColumnSort(colIndex);
    requestRender();
  }

  async function sortByColumnDirection(
    colIndex: number,
    direction: SpreadsheetSortDirection,
  ) {
    await spreadsheetState.setColumnSort(colIndex, direction);
    requestRender();
  }

  async function clearColumnSort() {
    await spreadsheetState.clearColumnSort();
    requestRender();
  }

  function focusMatch(match: SpreadsheetSearchMatch) {
    spreadsheetState.focusCell(match.display_row, match.col, match.source_row);
    requestRender();
  }

  function focusSearchResultAt(index: number) {
    const total = spreadsheetState.searchResults.length;
    if (total === 0) return;
    activeSearchIndex = ((index % total) + total) % total;
    const match = spreadsheetState.searchResults[activeSearchIndex];
    if (match) focusMatch(match);
  }

  function focusPreviousSearchResult() {
    focusSearchResultAt(activeSearchIndex - 1);
  }
  function focusNextSearchResult() {
    focusSearchResultAt(activeSearchIndex + 1);
  }

  async function runSearchForValue(value: string) {
    if (!value.trim()) return;
    searchInput = value;
    await spreadsheetState.searchWorkbook(value);
    activeSearchIndex = 0;
    requestRender();
  }
  async function applyFilterValue(value: string) {
    if (!value.trim()) return;
    filterInput = value;
    await spreadsheetState.applyFilter(value);
    requestRender();
  }
  async function clearViewOptions() {
    filterInput = "";
    await spreadsheetState.clearViewOptions();
    requestRender();
  }
  async function toggleHeaderRowMode() {
    await spreadsheetState.setHeaderRowEnabled(
      !spreadsheetState.headerRowEnabled,
    );
    requestRender();
  }
  async function reloadWorkbook() {
    if (filePathInput.trim())
      await spreadsheetState.loadExcelFile(filePathInput);
    requestRender();
  }
  function resetViewport() {
    spreadsheetState.resetViewport();
    requestRender();
  }

  // --- CLIPBOARD & OS CONTEXT MENUS ---
  async function copyText(value: string) {
    if (value) await navigator.clipboard.writeText(value).catch(() => {});
  }
  async function copyVisibleChunk() {
    await copyText(
      activeRows.map((r) => r.map((c) => c ?? "").join("\t")).join("\n"),
    );
  }
  async function copyRowValues(row: string[]) {
    await copyText(row.map((c) => c ?? "").join("\t"));
  }
  async function copySelectionValue(value = selectedValue) {
    await copyText(value);
  }
  async function copySelectedColumnValues() {
    if (selectedCol === null) return;
    const offset = selectedCol - activeStartCol;
    if (offset >= 0 && offset < activeColLimit)
      await copyText(activeRows.map((r) => r[offset] ?? "").join("\n"));
  }

  async function popupNativeContextMenu(
    event: MouseEvent,
    items: NativeMenuItem[],
  ) {
    if (!isDesktopRuntime) return;
    event.preventDefault();
    event.stopPropagation();
    const menu = await Menu.new({ items });
    try {
      const pos = new PhysicalPosition(
        Math.round(event.clientX),
        Math.round(event.clientY),
      );
      await menu.popup(pos);
    } finally {
      await menu.close().catch(() => {});
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
      { text: "Reset viewport", action: () => resetViewport() },
      {
        text: "Copy visible chunk",
        action: () => {
          void copyVisibleChunk();
        },
      },
      {
        text: spreadsheetState.headerRowEnabled
          ? "Disable header"
          : "Use first row as headers",
        action: () => {
          void toggleHeaderRowMode();
        },
      },
      {
        text: "Clear filter and sort",
        enabled:
          spreadsheetState.activeFilter.length > 0 ||
          spreadsheetState.sortCol !== null,
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
      { text: "Reset viewport", action: () => resetViewport() },
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
    const normalizedValue = (cellValue ?? "").trim();
    void popupNativeContextMenu(event, [
      {
        text: `Copy Cell Value`,
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
        text: `Find "${normalizedValue.slice(0, 20)}"`,
        enabled: normalizedValue.length > 0,
        action: () => {
          void runSearchForValue(normalizedValue);
        },
      },
      {
        text: `Filter by "${normalizedValue.slice(0, 20)}"`,
        enabled: normalizedValue.length > 0,
        action: () => {
          void applyFilterValue(normalizedValue);
        },
      },
    ]);
  }

  // --- INTERACTIVE FARE VE SCROLL KOORDİNAT HESAPLAMALARI (DERLEYİCİ HATASINI ÇÖZEN EN ÜSTE ALINAN BÖLÜM) ---
  function resolveEventCoordinates(event: MouseEvent) {
    const rect = canvasElement!.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;

    const { scrollTop, scrollLeft, rowHeight, columnWidth, rowHeaderWidth } =
      spreadsheetState;
    const headerHeight = rowHeight;

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
      handleRowContextMenu(
        event,
        clickedDisplayRow,
        sourceRow ?? clickedDisplayRow,
        activeRows[clickedDisplayRow - activeStartRow] ?? [],
      );
    } else if (
      isCell &&
      clickedDisplayRow !== null &&
      clickedDisplayCol !== null
    ) {
      const sourceRow = spreadsheetState.resolveSourceRow(clickedDisplayRow);
      const rowData = activeRows[clickedDisplayRow - activeStartRow] ?? [];
      handleCellContextMenu(
        event,
        clickedDisplayRow,
        clickedDisplayCol,
        sourceRow ?? clickedDisplayRow,
        rowData[clickedDisplayCol - activeStartCol] ?? "",
        rowData,
      );
    } else {
      handleViewportContextMenu(event);
    }
  }

  function handleGridWheel(event: WheelEvent) {
    if (!spreadsheetState.hasWorkbook) return;
    event.preventDefault();

    const nextTop = Math.min(
      Math.max(0, spreadsheetState.scrollTop + event.deltaY),
      spreadsheetState.totalContentHeight -
        viewportHeight +
        spreadsheetState.rowHeight,
    );
    const nextLeft = Math.min(
      Math.max(0, spreadsheetState.scrollLeft + event.deltaX),
      spreadsheetState.totalContentWidth -
        viewportWidth +
        spreadsheetState.rowHeaderWidth,
    );

    spreadsheetState.setScrollOffsets(nextTop, nextLeft);
    requestRender();
  }

  // --- CANVAS ASIL BOYAMA MOTORU ---
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

    if (
      canvasElement.width !== viewportWidth * dpr ||
      canvasElement.height !== viewportHeight * dpr
    ) {
      canvasElement.width = viewportWidth * dpr;
      canvasElement.height = viewportHeight * dpr;
      ctx.scale(dpr, dpr);
    }

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

    ctx.font = "12px 'IBM Plex Mono', monospace";
    ctx.textBaseline = "middle";

    activeRows.forEach((row, rIdx) => {
      const displayRow = startRowIdx + rIdx;
      const y = displayRow * rowHeight - scrollTop + headerHeight;

      if (y > viewportHeight || y + rowHeight < headerHeight) return;

      if (selectedRow === displayRow) {
        ctx.fillStyle = "rgba(111, 116, 88, 0.05)";
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
          ctx.fillStyle = "#eadebe";
        } else if (isActiveSearch) {
          ctx.fillStyle = "rgba(255, 191, 77, 0.35)";
        } else if (isSearchHit) {
          ctx.fillStyle = "rgba(255, 255, 0, 0.2)";
        } else {
          ctx.fillStyle = "#faf8ef";
        }

        ctx.fillRect(x, y, columnWidth, rowHeight);
        ctx.strokeStyle = "#c9c3ab";
        ctx.strokeRect(x, y, columnWidth, rowHeight);

        if (cell !== null && cell !== undefined && cell !== "") {
          ctx.fillStyle = "#24271d";
          ctx.save();
          ctx.beginPath();
          ctx.rect(x + 6, y, columnWidth - 12, rowHeight);
          ctx.clip();

          const strVal = formatCellValue(cell);
          if (isNumericCell(strVal)) {
            ctx.textAlign = "right";
            ctx.fillText(strVal, x + columnWidth - 6, y + rowHeight / 2);
          } else {
            ctx.textAlign = "left";
            ctx.fillText(strVal, x + 6, y + rowHeight / 2);
          }
          ctx.restore();
        }

        if (isSelected) {
          ctx.strokeStyle = "#4f563e";
          ctx.lineWidth = 2;
          ctx.strokeRect(x, y, columnWidth, rowHeight);
          ctx.lineWidth = 1;
        }
      });
    });

    ctx.fillStyle = "#ededf0";
    ctx.fillRect(
      rowHeaderWidth,
      0,
      viewportWidth - rowHeaderWidth,
      headerHeight,
    );
    ctx.font = "bold 11px 'IBM Plex Mono', monospace";

    activeColumnLabels.forEach((label, cIdx) => {
      const colIndex = startColIdx + cIdx;
      const x = colIndex * columnWidth - scrollLeft + rowHeaderWidth;

      if (x > viewportWidth || x + columnWidth < rowHeaderWidth) return;

      const isSorted = spreadsheetState.sortCol === colIndex;
      ctx.fillStyle = isSorted ? "#dfdac5" : "#ede9d9";
      ctx.fillRect(x, 0, columnWidth, headerHeight);
      ctx.strokeStyle = "#a79f82";
      ctx.strokeRect(x, 0, columnWidth, headerHeight);

      ctx.fillStyle = "#24271d";
      ctx.textAlign = "left";
      ctx.fillText(
        label.toUpperCase() + columnSortMarker(colIndex),
        x + 8,
        headerHeight / 2,
      );
    });

    ctx.fillStyle = "#ede9d9";
    ctx.fillRect(
      0,
      headerHeight,
      rowHeaderWidth,
      viewportHeight - headerHeight,
    );
    ctx.font = "11px 'IBM Plex Mono', monospace";
    ctx.textAlign = "center";

    activeRows.forEach((_, rIdx) => {
      const displayRow = startRowIdx + rIdx;
      const y = displayRow * rowHeight - scrollTop + headerHeight;

      if (y > viewportHeight || y + rowHeight < headerHeight) return;

      const isRowSelected = selectedRow === displayRow;
      const sourceRowIndex = activeSourceRows[rIdx] ?? displayRow;

      ctx.fillStyle = isRowSelected ? "#dbd5bc" : "#ede9d9";
      ctx.fillRect(0, y, rowHeaderWidth, rowHeight);
      ctx.strokeStyle = "#a79f82";
      ctx.strokeRect(0, y, rowHeaderWidth, rowHeight);

      ctx.fillStyle = "#696d58";
      ctx.fillText(
        (sourceRowIndex + 1).toString(),
        rowHeaderWidth / 2,
        y + rowHeight / 2,
      );
    });

    ctx.fillStyle = "#dbd5bc";
    ctx.fillRect(0, 0, rowHeaderWidth, headerHeight);
    ctx.strokeStyle = "#a79f82";
    ctx.strokeRect(0, 0, rowHeaderWidth, headerHeight);
    ctx.fillStyle = "#4f563e";
    ctx.font = "bold 10px 'IBM Plex Mono', monospace";
    ctx.fillText("INDEX", rowHeaderWidth / 2, headerHeight / 2);
  }

  // UI Fonksiyon Akışları
  async function loadWorkbook() {
    await spreadsheetState.loadExcelFile(filePathInput);
    searchInput = "";
    filterInput = spreadsheetState.filterQuery;
    jumpToRowInput = "";
    activeSearchIndex = 0;
    requestRender();
  }
  async function selectExcelFile() {
    if (!isDesktopRuntime) {
      spreadsheetState.loadError = "OS runtime required.";
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
      activeSearchIndex = 0;
      requestRender();
    } catch (e) {
      spreadsheetState.loadError = "Picker error.";
    }
  }
  function toggleInspectorPanel() {
    showInspector = !showInspector;
  }
  function goToFirstRow() {
    spreadsheetState.goToRow(1);
    requestRender();
  }
  function goToLastRow() {
    spreadsheetState.goToRow(Math.max(spreadsheetState.totalRows, 1));
    requestRender();
  }
  async function handleSheetChange(event: Event) {
    const target = event.currentTarget as HTMLSelectElement | null;
    if (!target) return;
    const idx = Number.parseInt(target.value, 10);
    if (!Number.isNaN(idx)) {
      await spreadsheetState.activateSheet(idx);
      requestRender();
    }
  }
  async function activateSheetTab(sheetIndex: number) {
    await spreadsheetState.activateSheet(sheetIndex);
    requestRender();
  }
  async function runSearch() {
    await spreadsheetState.searchWorkbook(searchInput);
    activeSearchIndex = 0;
    requestRender();
  }
  function clearSearch() {
    searchInput = "";
    activeSearchIndex = 0;
    spreadsheetState.clearSearchResults();
    requestRender();
  }
  async function applyFilter() {
    await spreadsheetState.applyFilter(filterInput);
    requestRender();
  }
  async function handleHeaderToggle(event: Event) {
    const target = event.currentTarget as HTMLInputElement | null;
    if (target) {
      await spreadsheetState.setHeaderRowEnabled(target.checked);
      requestRender();
    }
  }
  function goToVisibleRow() {
    const num = Number.parseInt(jumpToRowInput, 10);
    if (!Number.isNaN(num)) {
      spreadsheetState.goToRow(num);
      requestRender();
    }
  }
  function goToVisibleColumn() {
    const idx = parseColumnInput(jumpToColumnInput);
    if (idx === null) return;
    spreadsheetState.focusCell(
      selectedRow ?? spreadsheetState.verticalWindow.visibleStartRow,
      idx,
      spreadsheetState.resolveSourceRow(
        selectedRow ?? spreadsheetState.verticalWindow.visibleStartRow,
      ),
    );
    requestRender();
  }
  function scrubToRow(event: Event) {
    const t = event.currentTarget as HTMLInputElement | null;
    if (t) {
      spreadsheetState.goToRow(Number.parseInt(t.value, 10));
      requestRender();
    }
  }
  function scrubToColumn(event: Event) {
    const t = event.currentTarget as HTMLInputElement | null;
    if (!t) return;
    spreadsheetState.focusCell(
      selectedRow ?? spreadsheetState.verticalWindow.visibleStartRow,
      Math.max(0, Number.parseInt(t.value, 10) - 1),
      spreadsheetState.resolveSourceRow(
        selectedRow ?? spreadsheetState.verticalWindow.visibleStartRow,
      ),
    );
    requestRender();
  }
  function setDensityMode(mode: SpreadsheetDensityMode) {
    spreadsheetState.setDensityMode(mode);
    requestRender();
  }
  function setColumnProfile(profile: SpreadsheetColumnProfile) {
    spreadsheetState.setColumnProfile(profile);
    requestRender();
  }

  // Senkronizasyon Efektleri ($effect)
  $effect(() => {
    spreadsheetState.setViewportMetrics(viewportHeight, viewportWidth);
    requestRender();
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
  $effect(() => {
    if (
      spreadsheetState.searchResults.length > 0 &&
      activeSearchIndex >= spreadsheetState.searchResults.length
    )
      activeSearchIndex = 0;
  });

  $effect(() => {
    activeRows;
    spreadsheetState.searchResults;
    requestRender();
  });

  $effect(() => {
    if (canvasElement && !canvasContext) {
      canvasContext = canvasElement.getContext("2d", { alpha: false });
      requestRender();
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
      Excel loading requires the native Tauri desktop window.
    </div>
  {/if}

  <div class="sheets-toolbar shrink-0 border-b">
    <div class="flex flex-wrap items-center gap-1 border-b px-2 py-1">
      <span class="sheets-chip sheets-chip-strong">SHEETS</span>
      <span class="sheets-chip font-mono truncate max-w-xs"
        >{spreadsheetState.fileName || "NO FILE"}</span
      >
      <span class="sheets-chip">{spreadsheetState.sheetName || "NO SHEET"}</span
      >
      <span class="sheets-chip"
        >Rows: {toHumanCount(spreadsheetState.totalRows)}</span
      >
      <span class="sheets-chip"
        >Cols: {toHumanCount(spreadsheetState.totalCols)}</span
      >
      <span class="sheets-chip">Viewport: {visibleRowRange}</span>
      <span class="sheets-chip">{spreadsheetState.statusLabel}</span>
      {#if spreadsheetState.loadError}<span
          class="sheets-chip sheets-chip-danger"
          >{spreadsheetState.loadError}</span
        >{/if}
    </div>

    <div class="flex flex-wrap items-center gap-1 border-b px-2 py-1">
      <input
        bind:this={filePathInputElement}
        bind:value={filePathInput}
        class="sheets-input h-8 min-w-[280px] flex-1 font-mono text-xs"
        placeholder="Excel file path..."
        disabled={!isDesktopRuntime}
      />
      <button
        class="sheets-button-primary h-8 px-3 text-xs font-semibold"
        onclick={() => void loadWorkbook()}
        disabled={!isDesktopRuntime}>Load</button
      >
      <button
        class="sheets-button h-8 px-3 text-xs"
        onclick={selectExcelFile}
        disabled={!isDesktopRuntime}>Select</button
      >
      <button
        class="sheets-button h-8 px-3 text-xs"
        onclick={() => void reloadWorkbook()}
        disabled={!isDesktopRuntime || !spreadsheetState.filePath}
        >Reload</button
      >
      <button
        class="sheets-button h-8 px-3 text-xs"
        onclick={resetViewport}
        disabled={!spreadsheetState.hasWorkbook}>Reset</button
      >
      <button
        class="sheets-button h-8 px-3 text-xs"
        onclick={toggleInspectorPanel}
        disabled={!spreadsheetState.hasWorkbook}
        >{showInspector ? "Hide Panel" : "Show Panel"}</button
      >

      <select
        class="sheets-input h-8 min-w-[200px] font-mono text-xs"
        value={spreadsheetState.activeSheetIndex}
        onchange={handleSheetChange}
        disabled={!spreadsheetState.hasWorkbook}
      >
        {#each spreadsheetState.sheets as sheet, index}
          <option value={index}
            >{sheet.name} ({sheet.total_rows.toLocaleString()})</option
          >
        {/each}
      </select>
    </div>

    <div class="flex flex-wrap items-center gap-1 border-b px-2 py-1">
      <div class="sheets-name-box h-8 w-24 font-mono text-xs font-semibold">
        {selectedAddress}
      </div>
      <div class="sheets-formula-box flex h-8 flex-1 items-center gap-2 px-2">
        <span class="sheets-fx">fx</span><span
          class="truncate font-mono text-xs"
          >{selectedValue || "Select a cell..."}</span
        >
      </div>
      <button
        class="sheets-button h-8 px-3 text-xs"
        onclick={() => void copySelectionValue()}
        disabled={!selectedValue}>Copy</button
      >
      <label
        class="inline-flex h-8 items-center gap-2 px-2 text-[11px] text-[var(--sheet-muted)]"
      >
        <input
          type="checkbox"
          checked={spreadsheetState.headerRowEnabled}
          onchange={handleHeaderToggle}
          disabled={!spreadsheetState.hasWorkbook}
        /> Header row
      </label>
    </div>

    <div class="flex flex-wrap items-center gap-1 px-2 py-1">
      <div class="flex flex-1 items-center gap-1">
        <input
          bind:this={searchInputElement}
          bind:value={searchInput}
          class="sheets-input h-8 flex-1 font-mono text-xs"
          placeholder="Find text..."
          disabled={!spreadsheetState.hasWorkbook}
          onkeydown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              void runSearch();
            }
          }}
        />
        <button
          class="sheets-button h-8 px-3 text-xs"
          onclick={() => void runSearch()}
          disabled={isWorkbookBusy}>Find</button
        >
        <button class="sheets-button h-8 px-3 text-xs" onclick={clearSearch}
          >Clear</button
        >
        <button
          class="sheets-button h-8 px-3 text-xs"
          onclick={focusPreviousSearchResult}
          disabled={searchResultCount === 0}>Prev</button
        >
        <button
          class="sheets-button h-8 px-3 text-xs"
          onclick={focusNextSearchResult}
          disabled={searchResultCount === 0}>Next</button
        >
      </div>
      <div class="flex flex-1 items-center gap-1">
        <input
          bind:this={filterInputElement}
          bind:value={filterInput}
          class="sheets-input h-8 flex-1 font-mono text-xs"
          placeholder="Filter lines..."
          disabled={!spreadsheetState.hasWorkbook}
          onkeydown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              void applyFilter();
            }
          }}
        />
        <button
          class="sheets-button h-8 px-3 text-xs"
          onclick={() => void applyFilter()}
          disabled={isWorkbookBusy}>Filter</button
        >
        <button
          class="sheets-button h-8 px-3 text-xs"
          onclick={() => void clearViewOptions()}>Reset</button
        >
      </div>
    </div>
  </div>

  {#if spreadsheetState.hasWorkbook}
    <div class="flex min-h-0 flex-1">
      {#if showInspector}
        <aside
          class="sheets-sidebar enterprise-scrollbar hidden w-[20rem] shrink-0 overflow-auto border-r xl:block"
        >
          <div class="sheets-panel-section border-b px-3 py-3">
            <p class="sheets-section-label">Active Sheet</p>
            <h2 class="mt-1 text-sm font-semibold">
              {spreadsheetState.sheetName}
            </h2>
            <div class="mt-3 grid grid-cols-2 gap-2">
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Total Rows</span><span
                  class="sheets-metric-value"
                  >{toHumanCount(spreadsheetState.totalRows)}</span
                >
              </div>
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Columns</span><span
                  class="sheets-metric-value"
                  >{toHumanCount(spreadsheetState.totalCols)}</span
                >
              </div>
            </div>
          </div>

          <div class="sheets-panel-section border-b px-3 py-3">
            <p class="sheets-section-label">Quick Sweep Navigator</p>
            <div class="mt-2 rounded border p-2">
              <span class="text-[11px] text-[var(--sheet-muted)]"
                >Row Slider: {selectedDisplayRowLabel}</span
              >
              <input
                class="sheets-range mt-2 w-full"
                type="range"
                min="1"
                max={Math.max(1, spreadsheetState.totalRows)}
                value={(selectedRow ??
                  spreadsheetState.verticalWindow.visibleStartRow) + 1}
                oninput={scrubToRow}
              />
            </div>
            <div class="mt-2 rounded border p-2">
              <span class="text-[11px] text-[var(--sheet-muted)]"
                >Col Slider: {selectedColumnLabel}</span
              >
              <input
                class="sheets-range mt-2 w-full"
                type="range"
                min="1"
                max={Math.max(1, spreadsheetState.totalCols)}
                value={(selectedCol ??
                  spreadsheetState.horizontalWindow.visibleStartCol) + 1}
                oninput={scrubToColumn}
              />
            </div>
          </div>

          <div class="sheets-panel-section border-b px-3 py-3">
            <p class="sheets-section-label">Active Column Stats</p>
            {#if activeColumnNumericStats}
              <div class="mt-2 grid grid-cols-2 gap-2 text-[11px]">
                <div class="border p-1">
                  Avg: <b>{formatMetricValue(activeColumnNumericStats.avg)}</b>
                </div>
                <div class="border p-1">
                  Sum: <b>{formatMetricValue(activeColumnNumericStats.sum)}</b>
                </div>
                <div class="border p-1">
                  Min: <b>{formatMetricValue(activeColumnNumericStats.min)}</b>
                </div>
                <div class="border p-1">
                  Max: <b>{formatMetricValue(activeColumnNumericStats.max)}</b>
                </div>
              </div>
            {:else}
              <p class="text-[11px] text-gray-500 mt-1">
                No numbers found in focus chunk column.
              </p>
            {/if}
          </div>
        </aside>
      {/if}

      <div
        class="flex min-h-0 min-w-0 flex-1 flex-col relative bg-[var(--sheet-surface)]"
      >
        <div
          bind:this={viewportElement}
          bind:clientHeight={viewportHeight}
          bind:clientWidth={viewportWidth}
          class="absolute inset-0 outline-none overflow-hidden select-none"
          onwheel={handleGridWheel}
          role="grid"
          tabindex="0"
        >
          <canvas
            bind:this={canvasElement}
            class="absolute inset-0 cursor-cell"
            onclick={handleCanvasClick}
            oncontextmenu={handleCanvasContextMenu}
          ></canvas>
        </div>
      </div>
    </div>
  {:else}
    <div class="flex min-h-0 flex-1 items-center justify-center p-4">
      <div class="sheets-empty text-center p-8 max-w-lg border">
        <h2 class="text-lg font-bold">FLEXBOX Pure Canvas Matrix Active</h2>
        <p class="text-sm text-[var(--sheet-muted)] mt-2">
          GPU Acceleration enabled. Template loading streams millions of rows
          seamlessly.
        </p>
        <button
          class="sheets-button-primary mt-4 px-4 py-2 text-xs font-bold"
          onclick={selectExcelFile}>Open Excel Path Picker</button
        >
      </div>
    </div>
  {/if}
</section>

<style>
  .sheets-module {
    --sheet-bg: #f1efe4;
    --sheet-surface: #faf8ef;
    --sheet-surface-strong: #ece8d6;
    --sheet-border: #c9c3ab;
    --sheet-text: #24271d;
    --sheet-muted: #696d58;
    --sheet-accent: #6f7458;
    --sheet-accent-strong: #4f563e;
    background: var(--sheet-bg);
    color: var(--sheet-text);
    border-color: var(--sheet-border);
  }
  .sheets-toolbar,
  .sheets-sidebar {
    background: #faf8ef;
    border-color: var(--sheet-border);
  }
  .sheets-chip {
    display: inline-flex;
    align-items: center;
    padding: 0 0.5rem;
    height: 1.5rem;
    font-size: 11px;
    border: 1px solid var(--sheet-border);
    background: white;
    border-radius: 4px;
  }
  .sheets-chip-strong {
    background: #4f563e;
    color: white;
    font-weight: bold;
  }
  .sheets-chip-danger {
    background: #fde8e8;
    color: #9b1c1c;
    border-color: #f8b4b4;
  }
  .sheets-input {
    border: 1px solid var(--sheet-border);
    padding: 0 0.5rem;
    border-radius: 4px;
    background: white;
  }
  .sheets-input:focus {
    outline: none;
    border-color: var(--sheet-accent-strong);
  }
  .sheets-button,
  .sheets-button-primary {
    border: 1px solid var(--sheet-border);
    padding: 0 0.75rem;
    border-radius: 4px;
    font-size: 12px;
    height: 2rem;
    cursor: pointer;
  }
  .sheets-button-primary {
    background: #6f7458;
    color: white;
    border-color: #4f563e;
  }
  .sheets-button-primary:hover {
    background: #4f563e;
  }
  .sheets-button:hover {
    background: #ece8d6;
  }
  .sheets-name-box {
    border: 1px solid var(--sheet-border);
    display: inline-flex;
    align-items: center;
    justify-content: center;
    background: #ece8d6;
    border-radius: 4px;
  }
  .sheets-formula-box {
    border: 1px solid var(--sheet-border);
    background: white;
    border-radius: 4px;
  }
  .sheets-fx {
    color: #6f7458;
    font-weight: bold;
    font-size: 11px;
    margin-left: 0.5rem;
  }
  .sheets-sidebar {
    padding: 0.5rem;
  }
  .sheets-panel-section {
    padding: 0.5rem 0;
    border-color: var(--sheet-border);
  }
  .sheets-section-label {
    font-size: 10px;
    font-weight: bold;
    color: var(--sheet-muted);
    text-transform: uppercase;
    letter-spacing: 0.05em;
  }
  .sheets-metric-card {
    border: 1px solid var(--sheet-border);
    background: white;
    padding: 0.5rem;
    border-radius: 4px;
    display: flex;
    flex-direction: column;
  }
  .sheets-metric-label {
    font-size: 9px;
    color: var(--sheet-muted);
  }
  .sheets-metric-value {
    font-size: 12px;
    font-family: monospace;
  }
  .sheets-range {
    accent-color: #4f563e;
    width: 100%;
    cursor: pointer;
  }
  .sheets-empty {
    background: white;
    border-radius: 8px;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
  }
</style>
