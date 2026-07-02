<script lang="ts">
  import { onMount, untrack } from "svelte";
  import { isTauri } from "@tauri-apps/api/core";
  import { Menu } from "@tauri-apps/api/menu";
  import { open } from "@tauri-apps/plugin-dialog";
  import { PhysicalPosition } from "@tauri-apps/api/window";
  import {
    spreadsheetState,
    type SpreadsheetCellRef,
    type SpreadsheetColumnValueSummary,
    type SpreadsheetCellUpdate,
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

  type EditHistoryEntry = {
    label: string;
    before: SpreadsheetCellUpdate[];
    after: SpreadsheetCellUpdate[];
  };

  const isDesktopRuntime = isTauri();
  let filePathInput = $state("");
  let searchInput = $state("");
  let filterInput = $state("");
  let addressInput = $state("");
  let jumpToRowInput = $state("");
  let jumpToColumnInput = $state("");
  let viewportElement = $state<HTMLDivElement | null>(null);
  let addressInputElement = $state<HTMLInputElement | null>(null);
  let filePathInputElement = $state<HTMLInputElement | null>(null);
  let searchInputElement = $state<HTMLInputElement | null>(null);
  let filterInputElement = $state<HTMLInputElement | null>(null);
  let jumpToRowInputElement = $state<HTMLInputElement | null>(null);
  let formulaInputElement = $state<HTMLInputElement | null>(null);
  let filterPanelSearchElement = $state<HTMLInputElement | null>(null);
  let viewportHeight = $state(0);
  let viewportWidth = $state(0);
  let showSidebar = $state(true);
  let activeSearchIndex = $state(0);
  let recentSelections = $state<RecentSelection[]>([]);
  let selectionAnchorRow = $state<number | null>(null);
  let selectionAnchorCol = $state<number | null>(null);
  let isEditingCell = $state(false);
  let formulaDraft = $state("");
  let filterPanelCol = $state<number | null>(null);
  let filterPanelQuery = $state("");
  let filterPanelLoading = $state(false);
  let filterPanelValues = $state<SpreadsheetColumnValueSummary[]>([]);
  let undoStack = $state<EditHistoryEntry[]>([]);
  let redoStack = $state<EditHistoryEntry[]>([]);
  let fillDragState = $state<{
    mode: "row" | "col" | null;
    anchor: { startRow: number; endRow: number; startCol: number; endCol: number } | null;
    targetRow: number | null;
    targetCol: number | null;
  }>({
    mode: null,
    anchor: null,
    targetRow: null,
    targetCol: null,
  });
  let resizeDragState = $state<{
    kind: "column" | "row" | null;
    startPointerX: number;
    startPointerY: number;
    startColumnWidth: number;
    startRowHeight: number;
  }>({
    kind: null,
    startPointerX: 0,
    startPointerY: 0,
    startColumnWidth: 0,
    startRowHeight: 0,
  });

  let canvasElement = $state<HTMLCanvasElement | null>(null);
  let canvasContext = $state<CanvasRenderingContext2D | null>(null);
  let viewportResizeObserver: ResizeObserver | null = null;

  function updateViewportMetrics() {
    if (!viewportElement) {
      return;
    }

    const nextHeight = viewportElement.clientHeight;
    const nextWidth = viewportElement.clientWidth;

    if (nextHeight === viewportHeight && nextWidth === viewportWidth) {
      return;
    }

    viewportHeight = nextHeight;
    viewportWidth = nextWidth;
    spreadsheetState.setViewportMetrics(nextHeight, nextWidth);
    requestRender();
  }

  // --- DERIVED STATES ---
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
  const activeColumnFilterLabel = $derived.by(() => {
    if (
      spreadsheetState.columnFilterCol === null ||
      spreadsheetState.activeColumnFilter.length === 0
    ) {
      return "";
    }

    const activeOffset = spreadsheetState.columnFilterCol - activeStartCol;
    const headerLabel =
      activeColumnLabels[activeOffset] ??
      excelColumnLabel(spreadsheetState.columnFilterCol + 1);
    return `${headerLabel} = ${spreadsheetState.columnFilterValue}`;
  });
  const filterPanelColumnLabel = $derived.by(() => {
    if (filterPanelCol === null) return "";
    const activeOffset = filterPanelCol - activeStartCol;
    const headerLabel =
      activeColumnLabels[activeOffset] ??
      excelColumnLabel(filterPanelCol + 1);
    return headerLabel;
  });
  const canUndo = $derived(undoStack.length > 0);
  const canRedo = $derived(redoStack.length > 0);

  const selectionBounds = $derived.by(() => {
    if (
      selectedRow === null ||
      selectedCol === null ||
      selectionAnchorRow === null ||
      selectionAnchorCol === null
    ) {
      return null;
    }

    return {
      startRow: Math.min(selectionAnchorRow, selectedRow),
      endRow: Math.max(selectionAnchorRow, selectedRow),
      startCol: Math.min(selectionAnchorCol, selectedCol),
      endCol: Math.max(selectionAnchorCol, selectedCol),
    };
  });

  const selectionSummary = $derived.by(() => {
    if (!selectionBounds) return "1 cell";
    const rowCount = selectionBounds.endRow - selectionBounds.startRow + 1;
    const colCount = selectionBounds.endCol - selectionBounds.startCol + 1;
    if (rowCount === 1 && colCount === 1) return "1 cell";
    return `${rowCount} x ${colCount} cells`;
  });
  const fillPreviewBounds = $derived.by(() => {
    if (
      !fillDragState.anchor ||
      fillDragState.mode === null
    ) {
      return null;
    }

    if (fillDragState.mode === "row" && fillDragState.targetRow !== null) {
      if (
        fillDragState.targetRow >= fillDragState.anchor.startRow &&
        fillDragState.targetRow <= fillDragState.anchor.endRow
      ) {
        return null;
      }

      if (fillDragState.targetRow < fillDragState.anchor.startRow) {
        return {
          startRow: fillDragState.targetRow,
          endRow: fillDragState.anchor.startRow - 1,
          startCol: fillDragState.anchor.startCol,
          endCol: fillDragState.anchor.endCol,
        };
      }

      return {
        startRow: fillDragState.anchor.endRow + 1,
        endRow: fillDragState.targetRow,
        startCol: fillDragState.anchor.startCol,
        endCol: fillDragState.anchor.endCol,
      };
    }

    if (fillDragState.mode === "col" && fillDragState.targetCol !== null) {
      if (
        fillDragState.targetCol >= fillDragState.anchor.startCol &&
        fillDragState.targetCol <= fillDragState.anchor.endCol
      ) {
        return null;
      }

      if (fillDragState.targetCol < fillDragState.anchor.startCol) {
        return {
          startRow: fillDragState.anchor.startRow,
          endRow: fillDragState.anchor.endRow,
          startCol: fillDragState.targetCol,
          endCol: fillDragState.anchor.startCol - 1,
        };
      }

      return {
        startRow: fillDragState.anchor.startRow,
        endRow: fillDragState.anchor.endRow,
        startCol: fillDragState.anchor.endCol + 1,
        endCol: fillDragState.targetCol,
      };
    }

    return null;
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

  // --- CORE UTILS ---
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

  function syncSelectionInputs() {
    addressInput = selectedAddress;
    jumpToColumnInput =
      selectedCol !== null ? excelColumnLabel(selectedCol + 1) : "";
  }

  function setSelectionAnchor(row: number | null, col: number | null) {
    selectionAnchorRow = row;
    selectionAnchorCol = col;
  }

  function setActiveSelection(
    displayRowIndex: number,
    colIndex: number,
    sourceRowIndex: number | null,
    extendRange = false,
  ) {
    if (!extendRange || selectionAnchorRow === null || selectionAnchorCol === null) {
      setSelectionAnchor(displayRowIndex, colIndex);
    }

    spreadsheetState.setSelectedCell(displayRowIndex, colIndex, sourceRowIndex);
    syncSelectionInputs();
    pushRecentSelection();
    requestRender();
  }

  function syncWorkbookInputs() {
    filePathInput = spreadsheetState.filePath;
    filterInput = spreadsheetState.filterQuery;
    syncSelectionInputs();
  }

  function pushRecentSelection() {
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
  }

  function clearSelectionAnchor() {
    selectionAnchorRow = null;
    selectionAnchorCol = null;
  }

  function syncFormulaDraft() {
    if (isEditingCell) return;
    formulaDraft = selectedValue ?? "";
  }

  function startEditingCell(seedValue?: string) {
    if (
      selectedCol === null ||
      selectedRow === null ||
      selectedSourceRow === null
    ) {
      return;
    }

    isEditingCell = true;
    formulaDraft = seedValue ?? selectedValue ?? "";
    requestAnimationFrame(() => {
      formulaInputElement?.focus();
      formulaInputElement?.select();
    });
  }

  function cancelCellEdit() {
    isEditingCell = false;
    formulaDraft = selectedValue ?? "";
    viewportElement?.focus();
  }

  async function commitCellEdit() {
    if (
      !isEditingCell ||
      selectedCol === null ||
      selectedSourceRow === null
    ) {
      cancelCellEdit();
      return;
    }

    const previousValue = selectedValue ?? "";
    const afterUpdate: SpreadsheetCellUpdate = {
      source_row: selectedSourceRow,
      col_index: selectedCol,
      value: formulaDraft,
    };
    const result = await spreadsheetState.updateCellValue(
      selectedSourceRow,
      selectedCol,
      formulaDraft,
    );

    if (result) {
      pushHistoryEntry({
        label: "Edit cell",
        before: [
          {
            source_row: selectedSourceRow,
            col_index: selectedCol,
            value: previousValue,
          },
        ],
        after: [afterUpdate],
      });
      isEditingCell = false;
      syncFormulaDraft();
      requestRender();
      viewportElement?.focus();
    }
  }

  function pushHistoryEntry(entry: EditHistoryEntry) {
    undoStack = [...undoStack, entry].slice(-40);
    redoStack = [];
  }

  async function createHistoryEntry(
    label: string,
    updates: SpreadsheetCellUpdate[],
  ) {
    const before = await spreadsheetState.getCellValues(
      updates.map<SpreadsheetCellRef>((item) => ({
        source_row: item.source_row,
        col_index: item.col_index,
      })),
    );

    return {
      label,
      before,
      after: updates,
    } satisfies EditHistoryEntry;
  }

  async function applyHistoryEntry(
    entry: EditHistoryEntry,
    direction: "undo" | "redo",
  ) {
    const updates = direction === "undo" ? entry.before : entry.after;
    const result = await spreadsheetState.updateCellValues(updates);
    if (!result) {
      return false;
    }

    const focusTarget =
      updates[updates.length - 1] ?? entry.after[entry.after.length - 1];
    if (focusTarget) {
      spreadsheetState.setSelectedCell(
        spreadsheetState.selectedRow,
        focusTarget.col_index,
        focusTarget.source_row,
      );
    }

    syncFormulaDraft();
    requestRender();
    return true;
  }

  async function undoLastEdit() {
    const entry = undoStack[undoStack.length - 1];
    if (!entry) return;

    const applied = await applyHistoryEntry(entry, "undo");
    if (!applied) return;

    undoStack = undoStack.slice(0, -1);
    redoStack = [...redoStack, entry].slice(-40);
  }

  async function redoLastEdit() {
    const entry = redoStack[redoStack.length - 1];
    if (!entry) return;

    const applied = await applyHistoryEntry(entry, "redo");
    if (!applied) return;

    redoStack = redoStack.slice(0, -1);
    undoStack = [...undoStack, entry].slice(-40);
  }

  function parseClipboardMatrix(rawText: string) {
    return rawText
      .replace(/\r\n/g, "\n")
      .replace(/\r/g, "\n")
      .split("\n")
      .filter((_, index, entries) => !(index === entries.length - 1 && entries[index] === ""))
      .map((line) => line.split("\t"));
  }

  function buildCellUpdatesFromMatrix(matrix: string[][]) {
    if (
      matrix.length === 0 ||
      selectedRow === null ||
      selectedCol === null
    ) {
      return [];
    }

    const updates: SpreadsheetCellUpdate[] = [];

    for (let rowOffset = 0; rowOffset < matrix.length; rowOffset += 1) {
      const displayRowIndex = selectedRow + rowOffset;
      if (displayRowIndex >= spreadsheetState.totalRows) {
        break;
      }

      const sourceRowIndex = spreadsheetState.resolveSourceRow(displayRowIndex);
      if (sourceRowIndex === null) {
        continue;
      }

      const row = matrix[rowOffset] ?? [];
      for (let colOffset = 0; colOffset < row.length; colOffset += 1) {
        const colIndex = selectedCol + colOffset;
        if (colIndex >= spreadsheetState.totalCols) {
          break;
        }

        updates.push({
          source_row: sourceRowIndex,
          col_index: colIndex,
          value: row[colOffset] ?? "",
        });
      }
    }

    return updates;
  }

  async function applyClipboardMatrix(matrix: string[][]) {
    const updates = buildCellUpdatesFromMatrix(matrix);
    if (updates.length === 0) {
      return;
    }

    const historyEntry = await createHistoryEntry("Paste cells", updates);
    const result = await spreadsheetState.updateCellValues(updates);
    if (!result) {
      return;
    }
    pushHistoryEntry(historyEntry);

    const startRow = selectedRow ?? 0;
    const startCol = selectedCol ?? 0;
    const widestRow = Math.max(...matrix.map((row) => row.length), 1);
    const lastRow = Math.min(
      startRow + matrix.length - 1,
      Math.max(spreadsheetState.totalRows - 1, 0),
    );
    const lastCol = Math.min(
      startCol + widestRow - 1,
      Math.max(spreadsheetState.totalCols - 1, 0),
    );

    if (selectedRow !== null && selectedCol !== null) {
      setSelectionAnchor(startRow, startCol);
      spreadsheetState.setSelectedCell(
        lastRow,
        lastCol,
        spreadsheetState.resolveSourceRow(lastRow),
      );
    }

    syncFormulaDraft();
    requestRender();
  }

  async function pasteSelectionFromClipboard() {
    if (!navigator.clipboard || selectedRow === null || selectedCol === null) {
      return;
    }

    try {
      const rawText = await navigator.clipboard.readText();
      const matrix = parseClipboardMatrix(rawText);
      await applyClipboardMatrix(matrix);
    } catch {
      spreadsheetState.loadError = "Clipboard content could not be read.";
    }
  }

  async function clearSelectionValues() {
    if (
      selectedRow === null ||
      selectedCol === null
    ) {
      return;
    }

    const bounds = selectionBounds ?? {
      startRow: selectedRow,
      endRow: selectedRow,
      startCol: selectedCol,
      endCol: selectedCol,
    };

    const updates: SpreadsheetCellUpdate[] = [];
    for (let displayRowIndex = bounds.startRow; displayRowIndex <= bounds.endRow; displayRowIndex += 1) {
      const sourceRowIndex = spreadsheetState.resolveSourceRow(displayRowIndex);
      if (sourceRowIndex === null) {
        continue;
      }

      for (let colIndex = bounds.startCol; colIndex <= bounds.endCol; colIndex += 1) {
        updates.push({
          source_row: sourceRowIndex,
          col_index: colIndex,
          value: "",
        });
      }
    }

    const historyEntry = await createHistoryEntry("Clear cells", updates);
    const result = await spreadsheetState.updateCellValues(updates);
    if (result) {
      pushHistoryEntry(historyEntry);
      syncFormulaDraft();
      requestRender();
    }
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

  function parseAddressInput(rawValue: string) {
    const normalized = rawValue.trim().toUpperCase();
    const match = normalized.match(/^([A-Z]+)(\d+)$/);
    if (!match) return null;

    const [, colPart, rowPart] = match;
    const colIndex = parseColumnInput(colPart);
    const rowIndex = Number.parseInt(rowPart, 10) - 1;

    if (colIndex === null || Number.isNaN(rowIndex) || rowIndex < 0) {
      return null;
    }

    return { rowIndex, colIndex };
  }

  function getSelectionTemplateMatrix() {
    if (!selectionBounds) {
      return null;
    }

    const matrix: string[][] = [];
    for (let row = selectionBounds.startRow; row <= selectionBounds.endRow; row += 1) {
      const rowOffset = row - activeStartRow;
      const activeRow = activeRows[rowOffset];
      if (!activeRow) {
        return null;
      }

      const cells: string[] = [];
      for (let col = selectionBounds.startCol; col <= selectionBounds.endCol; col += 1) {
        const colOffset = col - activeStartCol;
        if (colOffset < 0 || colOffset >= activeColLimit) {
          return null;
        }
        cells.push(activeRow[colOffset] ?? "");
      }
      matrix.push(cells);
    }

    return matrix;
  }

  function isFillHandleHit(
    x: number,
    y: number,
    bounds: { startRow: number; endRow: number; startCol: number; endCol: number } | null,
  ) {
    if (!bounds) return false;

    const { rowHeight, columnWidth, rowHeaderWidth, scrollTop, scrollLeft } =
      spreadsheetState;
    const headerHeight = rowHeight;
    const handleSize = 10;
    const right =
      (bounds.endCol + 1) * columnWidth - scrollLeft + rowHeaderWidth;
    const bottom =
      (bounds.endRow + 1) * rowHeight - scrollTop + headerHeight;

    return (
      x >= right - handleSize &&
      x <= right + 2 &&
      y >= bottom - handleSize &&
      y <= bottom + 2
    );
  }

  function resolveResizeHotspot(
    x: number,
    y: number,
  ): { kind: "column" | "row" } | null {
    const { rowHeight, columnWidth, rowHeaderWidth, scrollTop, scrollLeft } =
      spreadsheetState;
    const headerHeight = rowHeight;
    const threshold = 6;

    if (y >= 0 && y <= headerHeight && x > rowHeaderWidth) {
      const localX = x - rowHeaderWidth + scrollLeft;
      const distanceToBoundary = Math.abs(localX % columnWidth);
      if (distanceToBoundary <= threshold || columnWidth - distanceToBoundary <= threshold) {
        return { kind: "column" };
      }
    }

    if (x >= 0 && x <= rowHeaderWidth && y > headerHeight) {
      const localY = y - headerHeight + scrollTop;
      const distanceToBoundary = Math.abs(localY % rowHeight);
      if (distanceToBoundary <= threshold || rowHeight - distanceToBoundary <= threshold) {
        return { kind: "row" };
      }
    }

    return null;
  }

  async function applyFillDrag() {
    if (!fillPreviewBounds || !fillDragState.anchor) {
      return;
    }

    const templateMatrix = getSelectionTemplateMatrix();
    if (!templateMatrix || templateMatrix.length === 0) {
      spreadsheetState.loadError =
        "Fill source must remain inside the loaded viewport chunk.";
      return;
    }

    const updates: SpreadsheetCellUpdate[] = [];

    if (fillDragState.mode === "row") {
      const rowSpan = templateMatrix.length;
      const fillStartsAbove = fillPreviewBounds.endRow < fillDragState.anchor.startRow;

      for (let row = fillPreviewBounds.startRow; row <= fillPreviewBounds.endRow; row += 1) {
        const sourceRowIndex = spreadsheetState.resolveSourceRow(row);
        if (sourceRowIndex === null) continue;

        const rowIndexInFill = fillStartsAbove
          ? fillPreviewBounds.endRow - row
          : row - fillPreviewBounds.startRow;
        const templateRow =
          templateMatrix[((rowIndexInFill % rowSpan) + rowSpan) % rowSpan] ?? [];

        for (let col = fillPreviewBounds.startCol; col <= fillPreviewBounds.endCol; col += 1) {
          const templateValue =
            templateRow[col - fillPreviewBounds.startCol] ?? "";
          updates.push({
            source_row: sourceRowIndex,
            col_index: col,
            value: templateValue,
          });
        }
      }
    } else if (fillDragState.mode === "col") {
      const fillStartsLeft = fillPreviewBounds.endCol < fillDragState.anchor.startCol;
      for (let row = fillPreviewBounds.startRow; row <= fillPreviewBounds.endRow; row += 1) {
        const sourceRowIndex = spreadsheetState.resolveSourceRow(row);
        if (sourceRowIndex === null) continue;

        const templateRow =
          templateMatrix[(row - fillPreviewBounds.startRow) % templateMatrix.length] ?? [];

        for (let col = fillPreviewBounds.startCol; col <= fillPreviewBounds.endCol; col += 1) {
          const colIndexInFill = fillStartsLeft
            ? fillPreviewBounds.endCol - col
            : col - fillPreviewBounds.startCol;
          const templateValue =
            templateRow[
              ((colIndexInFill % templateRow.length) + templateRow.length) %
                templateRow.length
            ] ?? "";
          updates.push({
            source_row: sourceRowIndex,
            col_index: col,
            value: templateValue,
          });
        }
      }
    }

    if (updates.length === 0) {
      return;
    }

    const historyEntry = await createHistoryEntry("Fill cells", updates);
    const result = await spreadsheetState.updateCellValues(updates);
    if (!result) {
      return;
    }

    pushHistoryEntry(historyEntry);
    spreadsheetState.setSelectedCell(
      fillDragState.mode === "row"
        ? (fillPreviewBounds.startRow < fillDragState.anchor.startRow
            ? fillPreviewBounds.startRow
            : fillPreviewBounds.endRow)
        : fillPreviewBounds.endRow,
      fillDragState.mode === "col"
        ? (fillPreviewBounds.startCol < fillDragState.anchor.startCol
            ? fillPreviewBounds.startCol
            : fillPreviewBounds.endCol)
        : fillPreviewBounds.endCol,
      spreadsheetState.resolveSourceRow(
        fillDragState.mode === "row"
          ? (fillPreviewBounds.startRow < fillDragState.anchor.startRow
              ? fillPreviewBounds.startRow
              : fillPreviewBounds.endRow)
          : fillPreviewBounds.endRow,
      ),
    );
    requestRender();
  }

  // --- CANVAS MOTORU KONTROLCÜLERİ ---
  let renderScheduled = false;
  function requestRender() {
    if (renderScheduled) return;
    renderScheduled = true;
    requestAnimationFrame(() => {
      renderScheduled = false;
      drawCanvas();
    });
  }

  function drawCanvas() {
    untrack(() => {
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

      ctx.fillStyle = "#f7f4ea";
      ctx.fillRect(0, 0, viewportWidth, viewportHeight);

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
      const selection = selectionBounds;
      const fillPreview = fillPreviewBounds;

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
          const isInSelectionRange =
            selection !== null &&
            displayRow >= selection.startRow &&
            displayRow <= selection.endRow &&
            displayCol >= selection.startCol &&
            displayCol <= selection.endCol;
          const isSearchHit = isSearchHighlighted(cell);
          const currentSourceRow = activeSourceRows[rIdx] ?? displayRow;
          const isActiveSearch = isActiveSearchCell(
            currentSourceRow,
            displayCol,
          );

          if (isSelected) {
            ctx.fillStyle = "#eadebe";
          } else if (isInSelectionRange) {
            ctx.fillStyle = "rgba(102, 112, 83, 0.14)";
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
        const isColumnSelected =
          selection !== null &&
          colIndex >= selection.startCol &&
          colIndex <= selection.endCol;
        ctx.fillStyle = isSorted
          ? "#dfdac5"
          : isColumnSelected
            ? "#e1dbc6"
            : "#ede9d9";
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

        const isRowSelected =
          selectedRow === displayRow ||
          (selection !== null &&
            displayRow >= selection.startRow &&
            displayRow <= selection.endRow);
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

      ctx.strokeStyle = "#ddd7c4";
      ctx.lineWidth = 1;

      const visibleBodyWidth = Math.max(0, viewportWidth - rowHeaderWidth);
      const visibleBodyHeight = Math.max(0, viewportHeight - headerHeight);
      const visibleColumnSlots = Math.ceil(visibleBodyWidth / columnWidth) + 1;
      const visibleRowSlots = Math.ceil(visibleBodyHeight / rowHeight) + 1;

      for (let slot = 0; slot <= visibleColumnSlots; slot += 1) {
        const x = rowHeaderWidth + slot * columnWidth - (scrollLeft % columnWidth);
        if (x < rowHeaderWidth - columnWidth || x > viewportWidth + columnWidth) {
          continue;
        }
        ctx.beginPath();
        ctx.moveTo(x + 0.5, headerHeight);
        ctx.lineTo(x + 0.5, viewportHeight);
        ctx.stroke();
      }

      for (let slot = 0; slot <= visibleRowSlots; slot += 1) {
        const y = headerHeight + slot * rowHeight - (scrollTop % rowHeight);
        if (y < headerHeight - rowHeight || y > viewportHeight + rowHeight) {
          continue;
        }
        ctx.beginPath();
        ctx.moveTo(rowHeaderWidth, y + 0.5);
        ctx.lineTo(viewportWidth, y + 0.5);
        ctx.stroke();
      }

      if (fillPreview) {
        const previewX =
          fillPreview.startCol * columnWidth - scrollLeft + rowHeaderWidth;
        const previewY =
          fillPreview.startRow * rowHeight - scrollTop + headerHeight;
        const previewWidth =
          (fillPreview.endCol - fillPreview.startCol + 1) * columnWidth;
        const previewHeight =
          (fillPreview.endRow - fillPreview.startRow + 1) * rowHeight;

        ctx.fillStyle = "rgba(102, 112, 83, 0.08)";
        ctx.fillRect(previewX, previewY, previewWidth, previewHeight);
        ctx.strokeStyle = "#4f563e";
        ctx.setLineDash([5, 3]);
        ctx.strokeRect(previewX, previewY, previewWidth, previewHeight);
        ctx.setLineDash([]);
      }

      if (selection) {
        const handleX =
          (selection.endCol + 1) * columnWidth - scrollLeft + rowHeaderWidth - 5;
        const handleY =
          (selection.endRow + 1) * rowHeight - scrollTop + headerHeight - 5;
        ctx.fillStyle = "#4f563e";
        ctx.fillRect(handleX, handleY, 8, 8);
        ctx.strokeStyle = "#ffffff";
        ctx.strokeRect(handleX, handleY, 8, 8);
      }
    });
  }

  // --- İNTERAKTİVİTE & FARE EVENTLERİ ---
  function resolveEventCoordinates(event: MouseEvent | PointerEvent) {
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
      x,
      y,
      isColHeader,
      isRowHeader,
      isCell,
      clickedDisplayCol,
      clickedDisplayRow,
    };
  }

  function handleCanvasPointerDown(event: PointerEvent) {
    if (!spreadsheetState.hasWorkbook || event.button !== 0) return;
    const coordinates = resolveEventCoordinates(event);
    const resizeHotspot = resolveResizeHotspot(coordinates.x, coordinates.y);

    if (resizeHotspot) {
      resizeDragState = {
        kind: resizeHotspot.kind,
        startPointerX: event.clientX,
        startPointerY: event.clientY,
        startColumnWidth: spreadsheetState.columnWidth,
        startRowHeight: spreadsheetState.rowHeight,
      };
      event.preventDefault();
      return;
    }

    if (isFillHandleHit(coordinates.x, coordinates.y, selectionBounds)) {
      fillDragState = {
        mode: null,
        anchor: selectionBounds
          ? {
              startRow: selectionBounds.startRow,
              endRow: selectionBounds.endRow,
              startCol: selectionBounds.startCol,
              endCol: selectionBounds.endCol,
            }
          : null,
        targetRow: null,
        targetCol: null,
      };
      event.preventDefault();
      return;
    }
  }

  function handleCanvasPointerMove(event: PointerEvent) {
    if (resizeDragState.kind === "column") {
      spreadsheetState.setColumnWidth(
        resizeDragState.startColumnWidth + (event.clientX - resizeDragState.startPointerX),
      );
      requestRender();
      return;
    }

    if (resizeDragState.kind === "row") {
      spreadsheetState.setRowHeight(
        resizeDragState.startRowHeight + (event.clientY - resizeDragState.startPointerY),
      );
      requestRender();
      return;
    }

    if (!fillDragState.anchor) return;

    const { clickedDisplayCol, clickedDisplayRow } = resolveEventCoordinates(event);
    if (clickedDisplayRow === null || clickedDisplayCol === null) {
      return;
    }

    const downwardDistance = clickedDisplayRow - fillDragState.anchor.endRow;
    const upwardDistance = fillDragState.anchor.startRow - clickedDisplayRow;
    const rightDistance = clickedDisplayCol - fillDragState.anchor.endCol;
    const leftDistance = fillDragState.anchor.startCol - clickedDisplayCol;
    const rowMagnitude = Math.max(downwardDistance, upwardDistance, 0);
    const colMagnitude = Math.max(rightDistance, leftDistance, 0);

    if (rowMagnitude === 0 && colMagnitude === 0) {
      fillDragState = { ...fillDragState, mode: null, targetRow: null, targetCol: null };
      requestRender();
      return;
    }

    if (rowMagnitude >= colMagnitude) {
      const targetRow =
        downwardDistance >= upwardDistance
          ? Math.min(clickedDisplayRow, spreadsheetState.totalRows - 1)
          : Math.max(0, clickedDisplayRow);
      fillDragState = {
        ...fillDragState,
        mode: "row",
        targetRow,
        targetCol: null,
      };
    } else {
      const targetCol =
        rightDistance >= leftDistance
          ? Math.min(clickedDisplayCol, spreadsheetState.totalCols - 1)
          : Math.max(0, clickedDisplayCol);
      fillDragState = {
        ...fillDragState,
        mode: "col",
        targetRow: null,
        targetCol,
      };
    }

    requestRender();
  }

  function resetFillDragState() {
    fillDragState = {
      mode: null,
      anchor: null,
      targetRow: null,
      targetCol: null,
    };
  }

  function resetResizeDragState() {
    resizeDragState = {
      kind: null,
      startPointerX: 0,
      startPointerY: 0,
      startColumnWidth: 0,
      startRowHeight: 0,
    };
  }

  async function handleCanvasPointerUp() {
    if (resizeDragState.kind !== null) {
      resetResizeDragState();
      requestRender();
      return;
    }

    if (!fillDragState.anchor) return;
    await applyFillDrag();
    resetFillDragState();
    requestRender();
  }

  function handleCanvasClick(event: MouseEvent) {
    if (!spreadsheetState.hasWorkbook) return;
    if (fillDragState.anchor) return;
    viewportElement?.focus();
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
      setActiveSelection(
        clickedDisplayRow,
        Math.max(selectedCol ?? 0, 0),
        spreadsheetState.resolveSourceRow(clickedDisplayRow),
        event.shiftKey,
      );
    } else if (
      isCell &&
      clickedDisplayRow !== null &&
      clickedDisplayCol !== null
    ) {
      setActiveSelection(
        clickedDisplayRow,
        clickedDisplayCol,
        spreadsheetState.resolveSourceRow(clickedDisplayRow),
        event.shiftKey,
      );
    }
  }

  function handleCanvasDoubleClick(event: MouseEvent) {
    if (!spreadsheetState.hasWorkbook) return;
    const { isCell, clickedDisplayCol, clickedDisplayRow } =
      resolveEventCoordinates(event);

    if (
      !isCell ||
      clickedDisplayRow === null ||
      clickedDisplayCol === null
    ) {
      return;
    }

    focusCell(
      clickedDisplayRow,
      clickedDisplayCol,
      spreadsheetState.resolveSourceRow(clickedDisplayRow),
    );
    startEditingCell();
  }

  function handleGridWheel(event: WheelEvent) {
    if (!spreadsheetState.hasWorkbook) return;
    event.preventDefault();

    const maxTop = Math.max(
      0,
      spreadsheetState.totalContentHeight -
        viewportHeight +
        spreadsheetState.rowHeight,
    );
    const maxLeft = Math.max(
      0,
      spreadsheetState.totalContentWidth -
        viewportWidth +
        spreadsheetState.rowHeaderWidth,
    );

    const nextTop = Math.min(
      Math.max(0, spreadsheetState.scrollTop + event.deltaY),
      maxTop,
    );
    const nextLeft = Math.min(
      Math.max(0, spreadsheetState.scrollLeft + event.deltaX),
      maxLeft,
    );

    spreadsheetState.setScrollOffsets(nextTop, nextLeft);
    requestRender();
  }

  function moveSelectionFromKeyboard(
    rowDelta: number,
    colDelta: number,
    extendRange = false,
  ) {
    const baseRow = selectedRow ?? spreadsheetState.verticalWindow.visibleStartRow;
    const baseCol = selectedCol ?? spreadsheetState.horizontalWindow.visibleStartCol;
    const nextRow = Math.min(
      Math.max(0, baseRow + rowDelta),
      Math.max(spreadsheetState.totalRows - 1, 0),
    );
    const nextCol = Math.min(
      Math.max(0, baseCol + colDelta),
      Math.max(spreadsheetState.totalCols - 1, 0),
    );

    if (!extendRange) {
      setSelectionAnchor(nextRow, nextCol);
    } else if (selectionAnchorRow === null || selectionAnchorCol === null) {
      setSelectionAnchor(baseRow, baseCol);
    }

    spreadsheetState.focusCell(
      nextRow,
      nextCol,
      spreadsheetState.resolveSourceRow(nextRow),
    );
    requestRender();
  }

  function focusEdgeSelection(kind: "row-start" | "row-end" | "grid-start" | "grid-end", extendRange = false) {
    if (!spreadsheetState.hasWorkbook) return;

    if (!extendRange) {
      clearSelectionAnchor();
    } else if (selectionAnchorRow === null || selectionAnchorCol === null) {
      setSelectionAnchor(selectedRow ?? 0, selectedCol ?? 0);
    }

    spreadsheetState.moveSelectionToEdge(kind);
    if (!extendRange) {
      setSelectionAnchor(
        spreadsheetState.selectedRow ?? 0,
        spreadsheetState.selectedCol ?? 0,
      );
    }
    requestRender();
  }

  function handleViewportKeydown(event: KeyboardEvent) {
    if (!spreadsheetState.hasWorkbook) {
      return;
    }

    if (isEditingCell && event.key !== "Escape") {
      return;
    }

    const extendRange = event.shiftKey;
    const ctrl = event.ctrlKey || event.metaKey;

    switch (event.key) {
      case "ArrowUp":
        event.preventDefault();
        moveSelectionFromKeyboard(-1, 0, extendRange);
        break;
      case "ArrowDown":
        event.preventDefault();
        moveSelectionFromKeyboard(1, 0, extendRange);
        break;
      case "ArrowLeft":
        event.preventDefault();
        moveSelectionFromKeyboard(0, -1, extendRange);
        break;
      case "ArrowRight":
        event.preventDefault();
        moveSelectionFromKeyboard(0, 1, extendRange);
        break;
      case "PageUp":
        event.preventDefault();
        moveSelectionFromKeyboard(-(spreadsheetState.verticalWindow.visibleRowCount - 1), 0, extendRange);
        break;
      case "PageDown":
        event.preventDefault();
        moveSelectionFromKeyboard(spreadsheetState.verticalWindow.visibleRowCount - 1, 0, extendRange);
        break;
      case "Home":
        event.preventDefault();
        if (ctrl) {
          focusEdgeSelection("grid-start", extendRange);
        } else {
          focusEdgeSelection("row-start", extendRange);
        }
        break;
      case "End":
        event.preventDefault();
        if (ctrl) {
          focusEdgeSelection("grid-end", extendRange);
        } else {
          focusEdgeSelection("row-end", extendRange);
        }
        break;
      case "Enter":
        event.preventDefault();
        if (isEditingCell) {
          void commitCellEdit();
        } else {
          moveSelectionFromKeyboard(1, 0, extendRange);
        }
        break;
      case "Tab":
        event.preventDefault();
        moveSelectionFromKeyboard(0, event.shiftKey ? -1 : 1, extendRange);
        break;
      case "Backspace":
      case "Delete":
        if (!ctrl && !event.altKey && !event.metaKey) {
          event.preventDefault();
          void clearSelectionValues();
        }
        break;
      case "Escape":
        if (isEditingCell) {
          event.preventDefault();
          cancelCellEdit();
        } else {
          clearSelectionAnchor();
        }
        requestRender();
        break;
      case "a":
      case "A":
        if (ctrl) {
          event.preventDefault();
          setSelectionAnchor(0, 0);
          spreadsheetState.focusCell(
            Math.max(spreadsheetState.totalRows - 1, 0),
            Math.max(spreadsheetState.totalCols - 1, 0),
            spreadsheetState.resolveSourceRow(Math.max(spreadsheetState.totalRows - 1, 0)),
          );
          requestRender();
        }
        break;
      case "f":
      case "F":
        if (ctrl) {
          event.preventDefault();
          focusElement(searchInputElement);
        }
        break;
      case "c":
      case "C":
        if (ctrl) {
          event.preventDefault();
          if (event.shiftKey) {
            void copySelectedColumnValues();
          } else {
            void copySelectionValue();
          }
        }
        break;
      case "z":
      case "Z":
        if (ctrl) {
          event.preventDefault();
          if (event.shiftKey) {
            void redoLastEdit();
          } else {
            void undoLastEdit();
          }
        }
        break;
      case "y":
      case "Y":
        if (ctrl) {
          event.preventDefault();
          void redoLastEdit();
        }
        break;
      case "v":
      case "V":
        if (ctrl) {
          event.preventDefault();
          void pasteSelectionFromClipboard();
        }
        break;
      case "l":
      case "L":
        if (ctrl) {
          event.preventDefault();
          focusElement(filterInputElement);
        }
        break;
      case "g":
      case "G":
        if (ctrl) {
          event.preventDefault();
          focusElement(addressInputElement);
        }
        break;
      case "F2":
        event.preventDefault();
        startEditingCell();
        break;
      default:
        if (
          !ctrl &&
          !event.altKey &&
          event.key.length === 1 &&
          selectedRow !== null &&
          selectedCol !== null
        ) {
          event.preventDefault();
          startEditingCell(event.key);
        }
        break;
    }
  }

  function focusCell(
    displayRowIndex: number,
    colIndex: number,
    sourceRowIndex: number | null,
  ) {
    setSelectionAnchor(displayRowIndex, colIndex);
    setActiveSelection(displayRowIndex, colIndex, sourceRowIndex, false);
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
    focusCell(match.display_row, match.col, match.source_row);
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
    showSidebar = true;
    searchInput = value;
    const results = await spreadsheetState.searchWorkbook(value);
    activeSearchIndex = 0;
    if (results.length > 0) {
      focusMatch(results[0]);
    }
    requestRender();
  }
  async function applyFilterValue(value: string) {
    if (!value.trim()) return;
    showSidebar = true;
    filterInput = value;
    await spreadsheetState.applyFilter(value);
    requestRender();
  }
  async function applySelectedValueColumnFilter(colIndex: number, value: string) {
    if (!value.trim()) return;
    showSidebar = true;
    await spreadsheetState.applyColumnFilter(colIndex, value);
    if (filterPanelCol === colIndex) {
      await loadFilterPanelValues();
    }
    requestRender();
  }
  async function clearViewOptions() {
    filterInput = "";
    await spreadsheetState.clearViewOptions();
    requestRender();
  }
  async function clearColumnFilter() {
    await spreadsheetState.clearColumnFilter();
    if (filterPanelCol !== null) {
      await loadFilterPanelValues();
    }
    requestRender();
  }
  async function loadFilterPanelValues() {
    if (filterPanelCol === null) {
      filterPanelValues = [];
      return;
    }

    filterPanelLoading = true;
    filterPanelValues = await spreadsheetState.getColumnValueSummary(
      filterPanelCol,
      filterPanelQuery,
      80,
    );
    filterPanelLoading = false;
  }
  async function openColumnFilterPanel(colIndex: number) {
    filterPanelCol = colIndex;
    filterPanelQuery = "";
    showSidebar = true;
    await loadFilterPanelValues();
    requestAnimationFrame(() => {
      filterPanelSearchElement?.focus();
      filterPanelSearchElement?.select();
    });
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
    showSidebar = true;
    requestRender();
  }
  function resetViewport() {
    spreadsheetState.resetViewport();
    clearSelectionAnchor();
    requestRender();
  }

  // --- PANOLAMA VE BAĞLAMSAL MENÜLER ---
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
    if (!selectionBounds) {
      await copyText(value);
      return;
    }

    const copiedRows: string[] = [];
    for (let row = selectionBounds.startRow; row <= selectionBounds.endRow; row += 1) {
      const rowOffset = row - activeStartRow;
      const activeRow = activeRows[rowOffset];
      if (!activeRow) {
        continue;
      }
      const cells: string[] = [];
      for (let col = selectionBounds.startCol; col <= selectionBounds.endCol; col += 1) {
        const colOffset = col - activeStartCol;
        cells.push(activeRow[colOffset] ?? "");
      }
      copiedRows.push(cells.join("\t"));
    }

    if (copiedRows.length > 0) {
      await copyText(copiedRows.join("\n"));
      return;
    }

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
          spreadsheetState.activeColumnFilter.length > 0 ||
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
    const selectedValueInColumn =
      selectedCol === colIndex ? (selectedValue ?? "").trim() : "";

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
        text: `Find selected value`,
        enabled: selectedValueInColumn.length > 0,
        action: () => {
          void runSearchForValue(selectedValueInColumn);
        },
      },
      {
        text: `Filter selected value`,
        enabled: selectedValueInColumn.length > 0,
        action: () => {
          void applySelectedValueColumnFilter(colIndex, selectedValueInColumn);
        },
      },
      {
        text: `Open filter panel`,
        action: () => {
          void openColumnFilterPanel(colIndex);
        },
      },
      {
        text: `Clear column filter`,
        enabled: spreadsheetState.activeColumnFilter.length > 0,
        action: () => {
          void clearColumnFilter();
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
        text: `Paste into selection`,
        action: () => {
          void pasteSelectionFromClipboard();
        },
      },
      {
        text: `Clear selection`,
        action: () => {
          void clearSelectionValues();
        },
      },
      {
        text: `Edit cell`,
        action: () => {
          startEditingCell();
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

  function handleViewportPaste(event: ClipboardEvent) {
    if (!spreadsheetState.hasWorkbook || selectedRow === null || selectedCol === null) {
      return;
    }

    const rawText = event.clipboardData?.getData("text/plain") ?? "";
    if (!rawText.trim().length) {
      return;
    }

    event.preventDefault();
    void applyClipboardMatrix(parseClipboardMatrix(rawText));
  }

  function goToAddress() {
    const target = parseAddressInput(addressInput);
    if (!target) {
      return;
    }

    const rowIndex = Math.min(
      target.rowIndex,
      Math.max(spreadsheetState.totalRows - 1, 0),
    );
    const colIndex = Math.min(
      target.colIndex,
      Math.max(spreadsheetState.totalCols - 1, 0),
    );

    focusCell(
      rowIndex,
      colIndex,
      spreadsheetState.resolveSourceRow(rowIndex),
    );
    requestRender();
  }

  // UI İŞLEMLERİ (Menüler ve Butonlar İçin)
  async function loadWorkbook() {
    await spreadsheetState.loadExcelFile(filePathInput);
    showSidebar = true;
    searchInput = "";
    filterInput = spreadsheetState.filterQuery;
    jumpToRowInput = "";
    activeSearchIndex = 0;
    clearSelectionAnchor();
    isEditingCell = false;
    resetFillDragState();
    resetResizeDragState();
    undoStack = [];
    redoStack = [];
    filterPanelCol = null;
    filterPanelQuery = "";
    filterPanelValues = [];
    syncWorkbookInputs();
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
      showSidebar = true;
      searchInput = "";
      filterInput = spreadsheetState.filterQuery;
      activeSearchIndex = 0;
      clearSelectionAnchor();
      isEditingCell = false;
      resetFillDragState();
      resetResizeDragState();
      undoStack = [];
      redoStack = [];
      filterPanelCol = null;
      filterPanelQuery = "";
      filterPanelValues = [];
      syncWorkbookInputs();
      requestRender();
    } catch (e) {
      spreadsheetState.loadError = "Picker error.";
    }
  }
  function toggleInspectorPanel() {
    showSidebar = !showSidebar;
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
      showSidebar = true;
      isEditingCell = false;
      resetFillDragState();
      resetResizeDragState();
      undoStack = [];
      redoStack = [];
      filterPanelCol = null;
      filterPanelQuery = "";
      filterPanelValues = [];
      requestRender();
    }
  }
  async function activateSheetTab(sheetIndex: number) {
    await spreadsheetState.activateSheet(sheetIndex);
    showSidebar = true;
    requestRender();
  }
  async function runSearch() {
    showSidebar = true;
    const results = await spreadsheetState.searchWorkbook(searchInput);
    activeSearchIndex = 0;
    if (results.length > 0) {
      focusMatch(results[0]);
    }
    requestRender();
  }
  function clearSearch() {
    searchInput = "";
    activeSearchIndex = 0;
    spreadsheetState.clearSearchResults();
    requestRender();
  }
  async function applyFilter() {
    showSidebar = true;
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
      setSelectionAnchor(
        spreadsheetState.selectedRow ?? 0,
        spreadsheetState.selectedCol ?? 0,
      );
      requestRender();
    }
  }
  function goToVisibleColumn() {
    const idx = parseColumnInput(jumpToColumnInput);
    if (idx === null) return;
    focusCell(
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
      setSelectionAnchor(
        spreadsheetState.selectedRow ?? 0,
        spreadsheetState.selectedCol ?? 0,
      );
      requestRender();
    }
  }
  function scrubToColumn(event: Event) {
    const t = event.currentTarget as HTMLInputElement | null;
    if (!t) return;
    focusCell(
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

  $effect(() => {
    spreadsheetState.activePacket;
    spreadsheetState.totalRows;
    spreadsheetState.totalCols;
    spreadsheetState.scrollTop;
    spreadsheetState.scrollLeft;
    viewportWidth;
    viewportHeight;
    showSidebar;
    requestRender();
  });

  $effect(() => {
    selectedValue;
    selectedCol;
    selectedRow;
    selectedSourceRow;
    syncFormulaDraft();
  });

  $effect(() => {
    if (!fillDragState.anchor || typeof window === "undefined") {
      return;
    }

    const handlePointerMove = (event: PointerEvent) => {
      handleCanvasPointerMove(event);
    };
    const handlePointerUp = () => {
      void handleCanvasPointerUp();
    };

    window.addEventListener("pointermove", handlePointerMove);
    window.addEventListener("pointerup", handlePointerUp);

    return () => {
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", handlePointerUp);
    };
  });

  $effect(() => {
    if (resizeDragState.kind === null || typeof window === "undefined") {
      return;
    }

    const handlePointerMove = (event: PointerEvent) => {
      handleCanvasPointerMove(event);
    };
    const handlePointerUp = () => {
      void handleCanvasPointerUp();
    };

    window.addEventListener("pointermove", handlePointerMove);
    window.addEventListener("pointerup", handlePointerUp);

    return () => {
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", handlePointerUp);
    };
  });

  // --- REAKTİF TRACKING HAVUZLARI ---
  // HATA ÇÖZÜMÜ: Svelte state'leri izler, ama işlem yaparken (untrack) zinciri kırar
  // --- REACTIVE BRIDGE ---
  onMount(() => {
    if (canvasElement && !canvasContext) {
      canvasContext = canvasElement.getContext("2d", { alpha: false });
    }

    return () => {
      viewportResizeObserver?.disconnect();
      viewportResizeObserver = null;
    };
  });

  $effect(() => {
    const currentViewport = viewportElement;

    viewportResizeObserver?.disconnect();
    viewportResizeObserver = null;

    if (!currentViewport) {
      return;
    }

    updateViewportMetrics();
    viewportResizeObserver = new ResizeObserver(() => {
      updateViewportMetrics();
    });
    viewportResizeObserver.observe(currentViewport);

    return () => {
      viewportResizeObserver?.disconnect();
      viewportResizeObserver = null;
    };
  });
</script>

<section
  class="sheets-module -m-4 flex h-[calc(100%+2rem)] w-[calc(100%+2rem)] min-h-0 flex-col overflow-hidden border"
>
  {#if !isDesktopRuntime}
    <div class="sheets-runtime-banner px-3 py-2 text-[11px]">
      Excel loading requires the native Tauri desktop window.
    </div>
  {/if}

  <div class="sheets-menubar shrink-0 border-b px-4 py-2">
    <div class="flex flex-wrap items-center gap-2">
      <span class="sheets-brand-badge">SS</span>
      <div class="min-w-0 flex-1">
        <div class="flex flex-wrap items-center gap-1.5">
          <span class="sheets-chip sheets-chip-strong">SHEETS</span>
          <span class="sheets-chip font-mono truncate max-w-xs">
            {spreadsheetState.fileName || 'NO FILE'}
          </span>
          <span class="sheets-chip">{spreadsheetState.statusLabel}</span>
          <span class="sheets-chip">Rows {toHumanCount(spreadsheetState.totalRows)}</span>
          <span class="sheets-chip">Cols {toHumanCount(spreadsheetState.totalCols)}</span>
          <span class="sheets-chip">Viewport {visibleRowRange}</span>
          <span class="sheets-chip">Cache {spreadsheetState.cacheEntryCount}</span>
          <span class="sheets-chip">Row H {spreadsheetState.rowHeight}</span>
          <span class="sheets-chip">Col W {spreadsheetState.columnWidth}</span>
          {#if spreadsheetState.activeFilter.length > 0}
            <span class="sheets-chip">Find Filter {spreadsheetState.filterQuery}</span>
          {/if}
          {#if activeColumnFilterLabel}
            <button
              class="sheets-chip"
              onclick={() => void clearColumnFilter()}
              title="Clear column filter"
            >
              Col Filter {activeColumnFilterLabel}
            </button>
          {/if}
          {#if spreadsheetState.loadError}
            <span class="sheets-chip sheets-chip-danger">{spreadsheetState.loadError}</span>
          {/if}
        </div>
      </div>
      <select
        class="sheets-input h-8 min-w-[220px] font-mono text-xs"
        value={spreadsheetState.activeSheetIndex}
        onchange={handleSheetChange}
        disabled={!spreadsheetState.hasWorkbook}
      >
        {#each spreadsheetState.sheets as sheet, index}
          <option value={index}>
            {sheet.name} ({sheet.total_rows.toLocaleString()})
          </option>
        {/each}
      </select>
    </div>
  </div>

  <div class="sheets-toolbar shrink-0 border-b px-3 py-2">
    <div class="flex flex-wrap items-center gap-2">
      <input
        bind:this={filePathInputElement}
        bind:value={filePathInput}
        class="sheets-input h-8 min-w-[260px] flex-[1.5] font-mono text-xs"
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
        >{showSidebar ? "Hide Sidebar" : "Show Sidebar"}</button
      >
      <input
        bind:this={addressInputElement}
        bind:value={addressInput}
        class="sheets-name-box h-8 w-24 px-2 font-mono text-xs font-semibold"
        placeholder="A1"
        onkeydown={(event) => {
          if (event.key === "Enter") {
            event.preventDefault();
            goToAddress();
          }
        }}
      />
      <div class="sheets-formula-box flex h-8 flex-1 items-center gap-2 px-2">
        <span class="sheets-fx">fx</span>
        <input
          bind:this={formulaInputElement}
          bind:value={formulaDraft}
          class="sheets-formula-input h-full min-w-0 flex-1 bg-transparent font-mono text-xs"
          placeholder="Select a cell..."
          disabled={selectedCol === null || selectedSourceRow === null}
          onfocus={() => {
            if (selectedCol !== null && selectedSourceRow !== null) {
              isEditingCell = true;
            }
          }}
          onblur={() => {
            if (isEditingCell) {
              void commitCellEdit();
            }
          }}
          onkeydown={(event) => {
            if (event.key === "Enter") {
              event.preventDefault();
              void commitCellEdit();
            } else if (event.key === "Escape") {
              event.preventDefault();
              cancelCellEdit();
            }
          }}
        />
      </div>
      <button
        class="sheets-button h-8 px-3 text-xs"
        onclick={() => void copySelectionValue()}
        disabled={!selectedValue && !selectionBounds}>Copy</button
      >
      <button
        class="sheets-button h-8 px-3 text-xs"
        onclick={() => void undoLastEdit()}
        disabled={!canUndo}>Undo</button
      >
      <button
        class="sheets-button h-8 px-3 text-xs"
        onclick={() => void redoLastEdit()}
        disabled={!canRedo}>Redo</button
      >
      <button
        class="sheets-button h-8 px-3 text-xs"
        onclick={() => {
          if (isEditingCell) {
            void commitCellEdit();
          } else {
            startEditingCell();
          }
        }}
        disabled={selectedCol === null || selectedSourceRow === null}
        >{isEditingCell ? "Apply" : "Edit"}</button
      >
      <span class="sheets-chip h-8">{selectionSummary}</span>
      <div class="flex items-center gap-1">
        <input
          bind:this={jumpToRowInputElement}
          bind:value={jumpToRowInput}
          class="sheets-input h-8 w-20 font-mono text-xs"
          placeholder="Row"
          inputmode="numeric"
          onkeydown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              goToVisibleRow();
            }
          }}
        />
        <input
          bind:value={jumpToColumnInput}
          class="sheets-input h-8 w-20 font-mono text-xs"
          placeholder="Col"
          onkeydown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              goToVisibleColumn();
            }
          }}
        />
        <button class="sheets-button h-8 px-3 text-xs" onclick={goToVisibleRow}
          >Go Row</button
        >
        <button
          class="sheets-button h-8 px-3 text-xs"
          onclick={goToVisibleColumn}
          >Go Col</button
        >
        <button class="sheets-button h-8 px-3 text-xs" onclick={goToAddress}
          >Go A1</button
        >
      </div>
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
      <div class="flex min-w-[18rem] flex-[1.2] items-center gap-1">
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
      <div class="flex min-w-[18rem] flex-[1.2] items-center gap-1">
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
      <div class="flex items-center gap-1">
        <select
          class="sheets-input h-8 min-w-[8rem] font-mono text-xs"
          value={spreadsheetState.sortCol ?? -1}
          onchange={async (event) => {
            const target = event.currentTarget as HTMLSelectElement;
            const value = Number.parseInt(target.value, 10);
            if (Number.isNaN(value) || value < 0) {
              await spreadsheetState.clearColumnSort();
              requestRender();
            } else {
              await spreadsheetState.setColumnSort(value, spreadsheetState.sortDirection ?? "asc");
              requestRender();
            }
          }}
          disabled={!spreadsheetState.hasWorkbook}
        >
          <option value="-1">Sort: none</option>
          {#each activeColumnLabels as label, index}
            <option value={activeStartCol + index}>
              Sort: {label}
            </option>
          {/each}
        </select>
        <button
          class="sheets-button h-8 px-3 text-xs"
          onclick={async () => {
            if (spreadsheetState.sortCol === null) return;
            await spreadsheetState.cycleColumnSort(spreadsheetState.sortCol);
            requestRender();
          }}
          disabled={spreadsheetState.sortCol === null}
          >Toggle Sort</button
        >
        <select
          class="sheets-input h-8 min-w-[7rem] font-mono text-xs"
          value={spreadsheetState.densityMode}
          onchange={(event) => {
            const target = event.currentTarget as HTMLSelectElement;
            setDensityMode(target.value as SpreadsheetDensityMode);
          }}
        >
          <option value="compact">Compact</option>
          <option value="balanced">Balanced</option>
          <option value="comfortable">Comfortable</option>
        </select>
        <select
          class="sheets-input h-8 min-w-[7rem] font-mono text-xs"
          value={spreadsheetState.columnProfile}
          onchange={(event) => {
            const target = event.currentTarget as HTMLSelectElement;
            setColumnProfile(target.value as SpreadsheetColumnProfile);
          }}
        >
          <option value="narrow">Narrow</option>
          <option value="standard">Standard</option>
          <option value="wide">Wide</option>
        </select>
      </div>
    </div>
  </div>

  {#if spreadsheetState.hasWorkbook}
    <div class="flex min-h-0 flex-1 gap-0">
      <div
        class="sheets-workspace relative flex min-h-0 min-w-0 flex-1 flex-col"
      >
        <div
          bind:this={viewportElement}
          class="sheets-viewport absolute inset-0 outline-none overflow-hidden select-none"
          onkeydown={handleViewportKeydown}
          onpaste={handleViewportPaste}
          onwheel={handleGridWheel}
          role="grid"
          aria-multiselectable="true"
          tabindex="0"
        >
          <canvas
            bind:this={canvasElement}
            class="absolute inset-0 cursor-cell"
            onclick={handleCanvasClick}
            onpointerdown={handleCanvasPointerDown}
            ondblclick={handleCanvasDoubleClick}
            oncontextmenu={handleCanvasContextMenu}
          ></canvas>
        </div>
      </div>

      {#if showSidebar}
        <aside
          class="sheets-sidebar enterprise-scrollbar w-[18rem] shrink-0 overflow-auto border-l lg:w-[20rem] xl:w-[22rem]"
        >
          <div class="sheets-panel-section border-b px-3 py-3">
            <p class="sheets-section-label">Workbook</p>
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
            <p class="sheets-section-label">Selection</p>
            <div class="mt-2 grid gap-2 text-[11px]">
              <div class="sheets-empty-card">
                <span class="sheets-metric-label">Address</span>
                <span class="sheets-metric-value">{selectedAddress}</span>
              </div>
              <div class="grid grid-cols-2 gap-2">
                <div class="sheets-metric-card">
                  <span class="sheets-metric-label">Row</span>
                  <span class="sheets-metric-value">{selectedDisplayRowLabel}</span>
                </div>
                <div class="sheets-metric-card">
                  <span class="sheets-metric-label">Column</span>
                  <span class="sheets-metric-value">{selectedColumnLabel}</span>
                </div>
              </div>
              <div class="sheets-empty-card">
                <span class="sheets-metric-label">Header</span>
                <span class="sheets-metric-value truncate">{selectedColumnHeader}</span>
              </div>
              <div class="sheets-empty-card">
                <span class="sheets-metric-label">Value</span>
                <span class="sheets-metric-value break-words">{selectedValue || "No value selected"}</span>
              </div>
            </div>
          </div>

          <div class="sheets-panel-section border-b px-3 py-3">
            <p class="sheets-section-label">Search</p>
            <div class="mt-2 grid gap-2 text-[11px]">
              <div class="sheets-metric-card">
                <span class="sheets-metric-label">Results</span>
                <span class="sheets-metric-value">{searchResultCount.toLocaleString()}</span>
              </div>
              <div class="sheets-empty-card">
                <span class="sheets-metric-label">Active Match</span>
                <span class="sheets-metric-value">
                  {#if activeSearchMatch}
                    {excelColumnLabel(activeSearchMatch.col + 1)}{activeSearchMatch.display_row + 1}
                  {:else}
                    -
                  {/if}
                </span>
              </div>
              <div class="sheets-empty-card">
                <span class="sheets-metric-label">Query</span>
                <span class="sheets-metric-value truncate">{spreadsheetState.searchQuery || "None"}</span>
              </div>
            </div>
          </div>

          <div class="sheets-panel-section border-b px-3 py-3">
            <p class="sheets-section-label">Navigator</p>
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
            <p class="sheets-section-label">Column Health</p>
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

          <div class="sheets-panel-section border-b px-3 py-3">
            <p class="sheets-section-label">Column Filter</p>
            {#if filterPanelCol !== null}
              <div class="mt-2 grid gap-2 text-[11px]">
                <div class="sheets-empty-card">
                  <span class="sheets-metric-label">Target</span>
                  <span class="sheets-metric-value truncate">
                    {filterPanelColumnLabel} ({excelColumnLabel(filterPanelCol + 1)})
                  </span>
                </div>
                <input
                  bind:this={filterPanelSearchElement}
                  bind:value={filterPanelQuery}
                  class="sheets-input h-8 font-mono text-xs"
                  placeholder="Search values..."
                  onkeydown={(event) => {
                    if (event.key === "Enter") {
                      event.preventDefault();
                      void loadFilterPanelValues();
                    }
                  }}
                />
                <div class="flex gap-2">
                  <button
                    class="sheets-button h-8 flex-1 px-3 text-xs"
                    onclick={() => void loadFilterPanelValues()}
                    disabled={filterPanelLoading}
                  >
                    {filterPanelLoading ? "Loading..." : "Refresh"}
                  </button>
                  <button
                    class="sheets-button h-8 flex-1 px-3 text-xs"
                    onclick={() => void clearColumnFilter()}
                    disabled={spreadsheetState.activeColumnFilter.length === 0}
                  >
                    Clear
                  </button>
                </div>
                <div class="max-h-72 space-y-2 overflow-auto pr-1">
                  {#each filterPanelValues as item (`${item.value}:${item.count}`)}
                    <button
                      class="sheets-empty-card min-h-[3rem] w-full text-left transition hover:border-[var(--sheet-accent-strong)] hover:bg-white"
                      onclick={() => void applySelectedValueColumnFilter(filterPanelCol, item.value)}
                    >
                      <span class="sheets-metric-label">
                        {item.count.toLocaleString()} rows
                      </span>
                      <span class="sheets-metric-value truncate">{item.value}</span>
                    </button>
                  {/each}
                  {#if !filterPanelLoading && filterPanelValues.length === 0}
                    <p class="text-[11px] text-gray-500">
                      No matching values in the current view.
                    </p>
                  {/if}
                </div>
              </div>
            {:else}
              <p class="mt-2 text-[11px] text-gray-500">
                Open a column header context menu and choose "Open filter panel".
              </p>
            {/if}
          </div>

          <div class="sheets-panel-section px-3 py-3">
            <p class="sheets-section-label">Recent</p>
            <div class="mt-2 space-y-2">
              {#each recentSelections as item (item.key)}
                <button
                  class="sheets-empty-card w-full text-left transition hover:border-[var(--sheet-accent-strong)] hover:bg-white"
                  onclick={() => {
                    spreadsheetState.focusCell(item.displayRow, item.colIndex, item.sourceRow);
                    requestRender();
                  }}
                >
                  <span class="sheets-metric-label">{item.address}</span>
                  <span class="sheets-metric-value truncate">{item.value || "Empty"}</span>
                </button>
              {/each}
              {#if recentSelections.length === 0}
                <p class="text-[11px] text-gray-500">No recent selections.</p>
              {/if}
            </div>
          </div>
        </aside>
      {/if}
    </div>
  {:else}
    <div class="flex min-h-0 flex-1 items-center justify-center p-4">
      <div class="sheets-empty text-center p-8 max-w-lg border">
        <h2 class="text-lg font-bold">FLEXBOX Sheets Ready</h2>
        <p class="text-sm text-[var(--sheet-muted)] mt-2">
          Load an Excel workbook to stream only the visible viewport through Rust-backed chunks.
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
    --sheet-bg: #e8e3d2;
    --sheet-surface: #f7f4ea;
    --sheet-surface-strong: #efe9d7;
    --sheet-surface-deep: #e3dcc6;
    --sheet-border: #c7c0a8;
    --sheet-border-strong: #a89f82;
    --sheet-text: #23261d;
    --sheet-muted: #6a6d58;
    --sheet-accent: #667053;
    --sheet-accent-strong: #44503b;
    --sheet-accent-soft: rgba(102, 112, 83, 0.14);
    --sheet-shadow: 0 18px 45px rgba(34, 38, 29, 0.08);
    --sheet-shadow-soft: 0 6px 18px rgba(34, 38, 29, 0.08);
    background:
      radial-gradient(circle at top left, rgba(255, 255, 255, 0.62), transparent 30%),
      radial-gradient(circle at right top, rgba(102, 112, 83, 0.09), transparent 20%),
      linear-gradient(180deg, rgba(255, 255, 255, 0.48), rgba(255, 255, 255, 0)) 0 0 / 100% 88px no-repeat,
      var(--sheet-bg);
    color: var(--sheet-text);
    border-color: var(--sheet-border);
  }

  .sheets-menubar,
  .sheets-toolbar,
  .sheets-sidebar,
  .sheets-empty {
    background: rgba(247, 244, 234, 0.96);
    backdrop-filter: blur(14px);
  }

  .sheets-menubar {
    border-color: var(--sheet-border);
    box-shadow: var(--sheet-shadow-soft);
  }

  .sheets-brand-badge {
    display: inline-flex;
    width: 1.75rem;
    height: 1.75rem;
    align-items: center;
    justify-content: center;
    border-radius: 0.35rem;
    background: linear-gradient(180deg, #5e694f, #3f4935);
    color: #fff;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.1em;
    box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.18);
  }

  .sheets-toolbar {
    border-color: var(--sheet-border);
    box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.5);
  }

  .sheets-chip {
    display: inline-flex;
    align-items: center;
    height: 1.5rem;
    padding: 0 0.5rem;
    border: 1px solid var(--sheet-border);
    border-radius: 0.4rem;
    background: rgba(255, 255, 255, 0.82);
    font-size: 11px;
    color: var(--sheet-muted);
    box-shadow: 0 1px 0 rgba(255, 255, 255, 0.75);
  }

  .sheets-chip-strong {
    background: linear-gradient(180deg, #5e694f, #444f3a);
    border-color: #3d4734;
    color: #fff;
    font-weight: 700;
    letter-spacing: 0.08em;
  }

  .sheets-chip-danger {
    background: rgba(161, 43, 43, 0.1);
    border-color: rgba(161, 43, 43, 0.18);
    color: #8f2f2f;
  }

  .sheets-input {
    border: 1px solid var(--sheet-border);
    border-radius: 0.42rem;
    background: rgba(255, 255, 255, 0.92);
    color: var(--sheet-text);
    box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.9);
    padding: 0 0.65rem;
  }

  .sheets-input::placeholder {
    color: #878a74;
  }

  .sheets-input:focus {
    outline: none;
    border-color: var(--sheet-accent-strong);
    box-shadow: 0 0 0 3px rgba(102, 112, 83, 0.14);
  }

  .sheets-button,
  .sheets-button-primary {
    height: 2rem;
    padding: 0 0.85rem;
    border: 1px solid var(--sheet-border);
    border-radius: 0.42rem;
    font-size: 12px;
    cursor: pointer;
    transition:
      transform 120ms ease,
      background-color 120ms ease,
      border-color 120ms ease,
      box-shadow 120ms ease;
  }

  .sheets-button {
    background: linear-gradient(180deg, rgba(255, 255, 255, 0.92), rgba(239, 234, 218, 0.94));
    color: var(--sheet-text);
  }

  .sheets-button:hover,
  .sheets-button-primary:hover {
    transform: translateY(-1px);
    box-shadow: 0 8px 16px rgba(34, 38, 29, 0.08);
  }

  .sheets-button-primary {
    background: linear-gradient(180deg, #687556, #4d583f);
    border-color: #445037;
    color: white;
  }

  .sheets-name-box,
  .sheets-formula-box {
    border: 1px solid var(--sheet-border);
    border-radius: 0.42rem;
    background: rgba(255, 255, 255, 0.92);
    box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.75);
  }

  .sheets-name-box {
    background: linear-gradient(180deg, #efead8, #e2dcc7);
    color: var(--sheet-accent-strong);
    text-align: center;
  }

  .sheets-fx {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 1.35rem;
    height: 1.35rem;
    margin-left: 0.25rem;
    border-radius: 999px;
    background: var(--sheet-surface-deep);
    color: var(--sheet-accent-strong);
    font-size: 10px;
    font-weight: 700;
  }

  .sheets-formula-input {
    outline: none;
    color: var(--sheet-text);
  }

  .sheets-formula-input::placeholder {
    color: #878a74;
  }

  .sheets-sidebar {
    padding: 0.65rem;
    border-color: var(--sheet-border);
    box-shadow: inset -1px 0 0 rgba(199, 192, 168, 0.55);
  }

  .sheets-panel-section {
    padding: 0.75rem 0;
    border-color: rgba(167, 159, 130, 0.42);
  }

  .sheets-section-label {
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 0.16em;
    text-transform: uppercase;
    color: var(--sheet-muted);
  }

  .sheets-metric-card,
  .sheets-empty-card {
    display: flex;
    min-height: 4rem;
    flex-direction: column;
    justify-content: space-between;
    gap: 0.35rem;
    border: 1px solid var(--sheet-border);
    border-radius: 0.55rem;
    background:
      linear-gradient(180deg, rgba(255, 255, 255, 0.9), rgba(241, 236, 220, 0.88));
    padding: 0.75rem;
    box-shadow: var(--sheet-shadow-soft);
  }

  .sheets-metric-label {
    font-size: 9px;
    font-weight: 700;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: var(--sheet-muted);
  }

  .sheets-metric-value {
    font-family: "IBM Plex Mono", "Cascadia Code", monospace;
    font-size: 13px;
    font-weight: 700;
    color: var(--sheet-text);
  }

  .sheets-range {
    width: 100%;
    accent-color: var(--sheet-accent-strong);
    cursor: pointer;
  }

  .sheets-workspace {
    background:
      radial-gradient(circle at top right, rgba(102, 112, 83, 0.05), transparent 24%),
      linear-gradient(180deg, rgba(255, 255, 255, 0.18), rgba(255, 255, 255, 0)) 0 0 / 100% 100% no-repeat,
      var(--sheet-surface);
  }

  .sheets-viewport {
    background:
      linear-gradient(180deg, rgba(255, 255, 255, 0.6), rgba(255, 255, 255, 0)) 0 0 / 100% 48px no-repeat,
      var(--sheet-surface);
    outline: none;
  }

  .sheets-empty {
    border: 1px solid var(--sheet-border);
    border-radius: 1rem;
    padding: 2rem;
    box-shadow: var(--sheet-shadow);
  }

  .sheets-runtime-banner {
    border-bottom: 1px solid rgba(146, 92, 15, 0.2);
    background: linear-gradient(180deg, rgba(255, 196, 96, 0.2), rgba(255, 196, 96, 0.12));
    color: #8b540f;
  }

  .sheets-toolbar,
  .sheets-sidebar {
    border-color: var(--sheet-border);
  }
</style>

