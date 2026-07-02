#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use calamine::{open_workbook_auto, Data, Range, Reader};
use serde::{de::DeserializeOwned, Deserialize, Serialize};
use std::{
    cmp::Ordering,
    collections::HashMap,
    fs,
    path::{Path, PathBuf},
    process::Command,
    sync::{Arc, Mutex},
    thread,
    time::Duration,
};
use tauri::Manager;
use tauri_plugin_shell::{
    process::{CommandChild, CommandEvent},
    ShellExt,
};

struct BackendSidecarState(Arc<Mutex<BackendSidecarRuntime>>);

#[derive(Default)]
struct BackendSidecarRuntime {
    child: Option<CommandChild>,
    child_pid: Option<u32>,
    shutting_down: bool,
    restart_attempts: u32,
}

struct SpreadsheetState(Arc<Mutex<SpreadsheetRuntime>>);

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
enum SpreadsheetSortDirection {
    Asc,
    Desc,
}

#[derive(Debug, Clone)]
enum SpreadsheetSortKey {
    Empty,
    Number(f64),
    Text(String),
}

#[derive(Debug, Default)]
struct SheetViewRuntime {
    header_row_enabled: bool,
    filter_query: Option<String>,
    column_filter_col: Option<usize>,
    column_filter_value: Option<String>,
    sort_col: Option<usize>,
    sort_direction: Option<SpreadsheetSortDirection>,
    visible_row_count: usize,
    visible_row_indices: Option<Vec<usize>>,
}

struct SheetRuntime {
    name: String,
    range: Range<Data>,
    edits: HashMap<(usize, usize), Data>,
    total_rows: usize,
    total_cols: usize,
    view: SheetViewRuntime,
}

#[derive(Default)]
struct SpreadsheetRuntime {
    file_path: Option<String>,
    file_name: Option<String>,
    sheets: Vec<SheetRuntime>,
    active_sheet_index: usize,
}

#[derive(Debug, Serialize, Deserialize)]
struct ThemePreference {
    theme: String,
}

#[derive(Debug, Serialize, Deserialize)]
struct WorkspaceSessionPayload {
    session_json: String,
}

#[derive(Debug, Serialize)]
struct SpreadsheetSheetSummary {
    name: String,
    total_rows: usize,
    total_cols: usize,
}

#[derive(Debug, Serialize)]
struct SpreadsheetLoadResult {
    file_name: String,
    active_sheet_index: usize,
    sheet_name: String,
    total_rows: usize,
    source_total_rows: usize,
    total_cols: usize,
    header_row_enabled: bool,
    filter_query: String,
    column_filter_col: Option<usize>,
    column_filter_value: String,
    sort_col: Option<usize>,
    sort_direction: Option<SpreadsheetSortDirection>,
    sheets: Vec<SpreadsheetSheetSummary>,
}

#[derive(Debug, Serialize)]
struct SpreadsheetChunkResponse {
    start_row: usize,
    row_limit: usize,
    start_col: usize,
    col_limit: usize,
    total_rows: usize,
    source_total_rows: usize,
    total_cols: usize,
    header_row_enabled: bool,
    headers: Vec<String>,
    source_rows: Vec<usize>,
    rows: Vec<Vec<String>>,
}

#[derive(Debug, Serialize)]
struct SpreadsheetSearchMatch {
    display_row: usize,
    source_row: usize,
    col: usize,
    value: String,
}

#[derive(Debug, Serialize)]
struct SpreadsheetColumnValueSummary {
    value: String,
    count: usize,
}

#[derive(Debug, Serialize, Deserialize)]
struct SpreadsheetCellUpdate {
    source_row: usize,
    col_index: usize,
    value: String,
}

#[derive(Debug, Deserialize)]
struct SpreadsheetCellRef {
    source_row: usize,
    col_index: usize,
}

struct LoadedWorkbook {
    file_name: String,
    sheets: Vec<SheetRuntime>,
}

fn app_config_file_path(app: &tauri::AppHandle, file_name: &str) -> Result<PathBuf, String> {
    let config_dir = app
        .path()
        .app_config_dir()
        .map_err(|error| error.to_string())?;

    fs::create_dir_all(&config_dir).map_err(|error| error.to_string())?;

    Ok(config_dir.join(file_name))
}

fn write_config_json<T: Serialize>(
    app: &tauri::AppHandle,
    file_name: &str,
    payload: &T,
) -> Result<(), String> {
    let path = app_config_file_path(app, file_name)?;
    let serialized = serde_json::to_string_pretty(payload).map_err(|error| error.to_string())?;
    fs::write(path, serialized).map_err(|error| error.to_string())
}

fn read_config_json<T: DeserializeOwned>(
    app: &tauri::AppHandle,
    file_name: &str,
) -> Result<Option<T>, String> {
    let path = app_config_file_path(app, file_name)?;

    if !path.exists() {
        return Ok(None);
    }

    let content = fs::read_to_string(path).map_err(|error| error.to_string())?;
    let payload = serde_json::from_str(&content).map_err(|error| error.to_string())?;
    Ok(Some(payload))
}

#[tauri::command]
fn save_theme_preference(app: tauri::AppHandle, theme: String) -> Result<String, String> {
    let payload = ThemePreference {
        theme: theme.clone(),
    };
    write_config_json(&app, "theme-preference.json", &payload)?;
    Ok(theme)
}

#[tauri::command]
fn load_theme_preference(app: tauri::AppHandle) -> Result<Option<String>, String> {
    let payload = read_config_json::<ThemePreference>(&app, "theme-preference.json")?;
    Ok(payload.map(|item| item.theme))
}

#[tauri::command]
fn save_workspace_session(app: tauri::AppHandle, session_json: String) -> Result<(), String> {
    let payload = WorkspaceSessionPayload { session_json };
    write_config_json(&app, "workspace-session.json", &payload)
}

#[tauri::command]
fn load_workspace_session(app: tauri::AppHandle) -> Result<Option<String>, String> {
    let payload = read_config_json::<WorkspaceSessionPayload>(&app, "workspace-session.json")?;
    Ok(payload.map(|item| item.session_json))
}

#[tauri::command]
fn clear_workspace_session(app: tauri::AppHandle) -> Result<(), String> {
    let path = app_config_file_path(&app, "workspace-session.json")?;
    if path.exists() {
        fs::remove_file(path).map_err(|error| error.to_string())?;
    }
    Ok(())
}

fn file_name_from_path(path: &Path) -> String {
    path.file_name()
        .and_then(|name| name.to_str())
        .map(ToOwned::to_owned)
        .unwrap_or_else(|| "workbook.xlsx".to_string())
}

fn serialize_cell(cell: &Data) -> String {
    match cell {
        Data::Int(value) => value.to_string(),
        Data::Float(value) => value.to_string(),
        Data::String(value) => value.clone(),
        Data::Bool(value) => value.to_string(),
        Data::DateTime(value) => value.to_string(),
        Data::DateTimeIso(value) => value.clone(),
        Data::DurationIso(value) => value.clone(),
        Data::Error(value) => format!("#{:?}", value),
        Data::Empty => String::new(),
    }
}

fn parse_cell_input(value: &str) -> Data {
    let trimmed = value.trim();

    if trimmed.is_empty() {
        return Data::Empty;
    }

    if trimmed.eq_ignore_ascii_case("true") {
        return Data::Bool(true);
    }

    if trimmed.eq_ignore_ascii_case("false") {
        return Data::Bool(false);
    }

    if let Ok(integer) = trimmed.parse::<i64>() {
        return Data::Int(integer);
    }

    if let Ok(float) = trimmed.parse::<f64>() {
        return Data::Float(float);
    }

    Data::String(value.to_string())
}

fn normalize_optional_query(query: Option<String>) -> Option<String> {
    query
        .map(|value| value.trim().to_string())
        .filter(|value| !value.is_empty())
}

fn parse_sort_number(value: &str) -> Option<f64> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        return None;
    }

    let normalized = if trimmed.contains(',') && !trimmed.contains('.') {
        trimmed.replace(',', ".")
    } else {
        trimmed.to_string()
    };

    normalized.parse::<f64>().ok()
}

fn build_sort_key(value: &str) -> SpreadsheetSortKey {
    let trimmed = value.trim();

    if trimmed.is_empty() {
        return SpreadsheetSortKey::Empty;
    }

    if let Some(number) = parse_sort_number(trimmed) {
        return SpreadsheetSortKey::Number(number);
    }

    SpreadsheetSortKey::Text(trimmed.to_lowercase())
}

fn compare_sort_keys(left: &SpreadsheetSortKey, right: &SpreadsheetSortKey) -> Ordering {
    match (left, right) {
        (SpreadsheetSortKey::Empty, SpreadsheetSortKey::Empty) => Ordering::Equal,
        (SpreadsheetSortKey::Empty, _) => Ordering::Greater,
        (_, SpreadsheetSortKey::Empty) => Ordering::Less,
        (SpreadsheetSortKey::Number(left_value), SpreadsheetSortKey::Number(right_value)) => {
            left_value
                .partial_cmp(right_value)
                .unwrap_or(Ordering::Equal)
        }
        (SpreadsheetSortKey::Number(_), SpreadsheetSortKey::Text(_)) => Ordering::Less,
        (SpreadsheetSortKey::Text(_), SpreadsheetSortKey::Number(_)) => Ordering::Greater,
        (SpreadsheetSortKey::Text(left_value), SpreadsheetSortKey::Text(right_value)) => {
            left_value.cmp(right_value)
        }
    }
}

fn cell_value_at(sheet: &SheetRuntime, source_row: usize, col_index: usize) -> String {
    if let Some(value) = sheet.edits.get(&(source_row, col_index)) {
        return serialize_cell(value);
    }

    sheet.range
        .get((source_row, col_index))
        .map(serialize_cell)
        .unwrap_or_default()
}

fn row_matches_query(sheet: &SheetRuntime, source_row: usize, query: &str) -> bool {
    for col_index in 0..sheet.total_cols {
        let value = cell_value_at(sheet, source_row, col_index);

        if value.to_lowercase().contains(query) {
            return true;
        }
    }

    false
}

fn row_matches_column_filter(
    sheet: &SheetRuntime,
    source_row: usize,
    column_filter_col: usize,
    column_filter_value: &str,
) -> bool {
    cell_value_at(sheet, source_row, column_filter_col)
        .to_lowercase()
        .contains(column_filter_value)
}

fn rebuild_sheet_view(sheet: &mut SheetRuntime) {
    let data_start_row = if sheet.view.header_row_enabled && sheet.total_rows > 0 {
        1
    } else {
        0
    };

    if data_start_row >= sheet.total_rows {
        sheet.view.visible_row_count = 0;
        sheet.view.visible_row_indices = None;
        return;
    }

    let normalized_query = sheet
        .view
        .filter_query
        .as_ref()
        .map(|value| value.to_lowercase())
        .filter(|value| !value.is_empty());
    let column_filter_col = sheet
        .view
        .column_filter_col
        .filter(|column| *column < sheet.total_cols);
    let column_filter_value = sheet
        .view
        .column_filter_value
        .as_ref()
        .map(|value| value.to_lowercase())
        .filter(|value| !value.is_empty());
    sheet.view.column_filter_col = column_filter_col;
    let sort_col = sheet
        .view
        .sort_col
        .filter(|column| *column < sheet.total_cols);
    sheet.view.sort_col = sort_col;
    if column_filter_col.is_none() || column_filter_value.is_none() {
        sheet.view.column_filter_value = None;
    }

    let needs_materialized_rows =
        normalized_query.is_some() || sort_col.is_some() || column_filter_col.is_some();
    if !needs_materialized_rows {
        sheet.view.visible_row_count = sheet.total_rows - data_start_row;
        sheet.view.visible_row_indices = None;
        return;
    }

    let mut row_indices: Vec<usize> = (data_start_row..sheet.total_rows).collect();

    if let Some(query) = normalized_query.as_deref() {
        row_indices.retain(|row_index| row_matches_query(sheet, *row_index, query));
    }

    if let (Some(filter_col), Some(filter_value)) = (column_filter_col, column_filter_value.as_deref())
    {
        row_indices.retain(|row_index| {
            row_matches_column_filter(sheet, *row_index, filter_col, filter_value)
        });
    }

    if let Some(sort_column) = sort_col {
        let sort_direction = sheet
            .view
            .sort_direction
            .unwrap_or(SpreadsheetSortDirection::Asc);
        let mut keyed_rows: Vec<(usize, SpreadsheetSortKey)> = row_indices
            .into_iter()
            .map(|row_index| {
                let value = cell_value_at(sheet, row_index, sort_column);
                (row_index, build_sort_key(&value))
            })
            .collect();

        keyed_rows.sort_by(|left, right| compare_sort_keys(&left.1, &right.1));
        if sort_direction == SpreadsheetSortDirection::Desc {
            keyed_rows.reverse();
        }

        row_indices = keyed_rows
            .into_iter()
            .map(|(row_index, _)| row_index)
            .collect();
    }

    sheet.view.visible_row_count = row_indices.len();
    sheet.view.visible_row_indices = Some(row_indices);
}

impl SheetRuntime {
    fn data_start_row(&self) -> usize {
        if self.view.header_row_enabled && self.total_rows > 0 {
            1
        } else {
            0
        }
    }

    fn resolve_source_row(&self, display_row: usize) -> Option<usize> {
        if display_row >= self.view.visible_row_count {
            return None;
        }

        if let Some(visible_row_indices) = self.view.visible_row_indices.as_ref() {
            return visible_row_indices.get(display_row).copied();
        }

        Some(self.data_start_row() + display_row)
    }

    fn visible_headers(&self, start_col: usize, col_limit: usize) -> Vec<String> {
        let end_col = start_col.saturating_add(col_limit).min(self.total_cols);
        let mut headers = Vec::with_capacity(end_col.saturating_sub(start_col));

        for col_index in start_col..end_col {
            let value = if self.view.header_row_enabled && self.total_rows > 0 {
                cell_value_at(self, 0, col_index)
            } else {
                String::new()
            };

            headers.push(value);
        }

        headers
    }
}

impl SpreadsheetRuntime {
    fn active_sheet(&self) -> Option<&SheetRuntime> {
        self.sheets.get(self.active_sheet_index)
    }

    fn active_sheet_mut(&mut self) -> Option<&mut SheetRuntime> {
        self.sheets.get_mut(self.active_sheet_index)
    }
}

fn build_sheet_summaries(sheets: &[SheetRuntime]) -> Vec<SpreadsheetSheetSummary> {
    sheets
        .iter()
        .map(|sheet| SpreadsheetSheetSummary {
            name: sheet.name.clone(),
            total_rows: sheet.total_rows,
            total_cols: sheet.total_cols,
        })
        .collect()
}

fn build_load_result(
    file_name: &str,
    active_sheet_index: usize,
    sheets: &[SheetRuntime],
) -> Result<SpreadsheetLoadResult, String> {
    let active_sheet = sheets
        .get(active_sheet_index)
        .ok_or_else(|| "workbook does not contain an active worksheet".to_string())?;

    Ok(SpreadsheetLoadResult {
        file_name: file_name.to_string(),
        active_sheet_index,
        sheet_name: active_sheet.name.clone(),
        total_rows: active_sheet.view.visible_row_count,
        source_total_rows: active_sheet.total_rows,
        total_cols: active_sheet.total_cols,
        header_row_enabled: active_sheet.view.header_row_enabled,
        filter_query: active_sheet.view.filter_query.clone().unwrap_or_default(),
        column_filter_col: active_sheet.view.column_filter_col,
        column_filter_value: active_sheet.view.column_filter_value.clone().unwrap_or_default(),
        sort_col: active_sheet.view.sort_col,
        sort_direction: active_sheet.view.sort_direction,
        sheets: build_sheet_summaries(sheets),
    })
}

#[tauri::command]
async fn load_excel_file(
    state: tauri::State<'_, SpreadsheetState>,
    path: String,
) -> Result<SpreadsheetLoadResult, String> {
    let workbook_path = PathBuf::from(&path);

    if !workbook_path.exists() {
        return Err("selected Excel file does not exist".to_string());
    }

    let loaded_workbook =
        tauri::async_runtime::spawn_blocking(move || -> Result<LoadedWorkbook, String> {
            let mut workbook = open_workbook_auto(&workbook_path)
                .map_err(|error| format!("failed to open Excel file: {error}"))?;

            let sheet_names = workbook.sheet_names().to_owned();
            if sheet_names.is_empty() {
                return Err("workbook does not contain any worksheets".to_string());
            }

            let mut sheets = Vec::with_capacity(sheet_names.len());

            for (sheet_index, sheet_name) in sheet_names.into_iter().enumerate() {
                let range = workbook
                    .worksheet_range_at(sheet_index)
                    .ok_or_else(|| format!("failed to read worksheet at index {sheet_index}"))?
                    .map_err(|error| format!("failed to parse worksheet '{sheet_name}': {error}"))?;
                let (total_rows, total_cols) = range.get_size();
                let mut sheet = SheetRuntime {
                    name: sheet_name,
                    range,
                    edits: HashMap::new(),
                    total_rows,
                    total_cols,
                    view: SheetViewRuntime::default(),
                };
                rebuild_sheet_view(&mut sheet);
                sheets.push(sheet);
            }

            Ok(LoadedWorkbook {
                file_name: file_name_from_path(&workbook_path),
                sheets,
            })
        })
        .await
        .map_err(|error| format!("failed to join Excel loader: {error}"))??;

    let result = build_load_result(&loaded_workbook.file_name, 0, &loaded_workbook.sheets)?;

    let mut runtime = state
        .0
        .lock()
        .map_err(|_| "failed to acquire spreadsheet state".to_string())?;
    runtime.file_path = Some(path);
    runtime.file_name = Some(loaded_workbook.file_name);
    runtime.sheets = loaded_workbook.sheets;
    runtime.active_sheet_index = 0;

    Ok(result)
}

#[tauri::command]
fn set_active_sheet(
    state: tauri::State<'_, SpreadsheetState>,
    sheet_index: usize,
) -> Result<SpreadsheetLoadResult, String> {
    let mut runtime = state
        .0
        .lock()
        .map_err(|_| "failed to acquire spreadsheet state".to_string())?;

    if runtime.sheets.is_empty() {
        return Err("no Excel workbook has been loaded".to_string());
    }

    if sheet_index >= runtime.sheets.len() {
        return Err("requested worksheet index is out of range".to_string());
    }

    runtime.active_sheet_index = sheet_index;
    let file_name = runtime
        .file_name
        .clone()
        .unwrap_or_else(|| "workbook.xlsx".to_string());

    build_load_result(&file_name, runtime.active_sheet_index, &runtime.sheets)
}

#[tauri::command]
async fn configure_active_sheet_view(
    state: tauri::State<'_, SpreadsheetState>,
    header_row_enabled: bool,
    filter_query: Option<String>,
    column_filter_col: Option<usize>,
    column_filter_value: Option<String>,
    sort_col: Option<usize>,
    sort_direction: Option<SpreadsheetSortDirection>,
) -> Result<SpreadsheetLoadResult, String> {
    let runtime = state.0.clone();

    tauri::async_runtime::spawn_blocking(move || -> Result<SpreadsheetLoadResult, String> {
        let mut runtime = runtime
            .lock()
            .map_err(|_| "failed to acquire spreadsheet state".to_string())?;

        let sheet = runtime
            .active_sheet_mut()
            .ok_or_else(|| "no Excel workbook has been loaded".to_string())?;
        sheet.view.header_row_enabled = header_row_enabled;
        sheet.view.filter_query = normalize_optional_query(filter_query);
        sheet.view.column_filter_col = column_filter_col.filter(|column| *column < sheet.total_cols);
        sheet.view.column_filter_value = normalize_optional_query(column_filter_value);
        if sheet.view.column_filter_col.is_none() || sheet.view.column_filter_value.is_none() {
            sheet.view.column_filter_col = None;
            sheet.view.column_filter_value = None;
        }
        sheet.view.sort_col = sort_col.filter(|column| *column < sheet.total_cols);
        sheet.view.sort_direction = if sheet.view.sort_col.is_some() {
            sort_direction.or(Some(SpreadsheetSortDirection::Asc))
        } else {
            None
        };
        rebuild_sheet_view(sheet);

        let file_name = runtime
            .file_name
            .clone()
            .unwrap_or_else(|| "workbook.xlsx".to_string());
        build_load_result(&file_name, runtime.active_sheet_index, &runtime.sheets)
    })
    .await
    .map_err(|error| format!("failed to configure worksheet view: {error}"))?
}

#[tauri::command]
async fn find_in_active_sheet(
    state: tauri::State<'_, SpreadsheetState>,
    query: String,
    limit: usize,
) -> Result<Vec<SpreadsheetSearchMatch>, String> {
    let normalized_query = query.trim().to_lowercase();

    if normalized_query.is_empty() {
        return Ok(Vec::new());
    }

    let runtime = state.0.clone();

    tauri::async_runtime::spawn_blocking(move || -> Result<Vec<SpreadsheetSearchMatch>, String> {
        let runtime = runtime
            .lock()
            .map_err(|_| "failed to acquire spreadsheet state".to_string())?;
        let sheet = runtime
            .active_sheet()
            .ok_or_else(|| "no Excel workbook has been loaded".to_string())?;
        let result_limit = limit.clamp(1, 200);
        let mut matches = Vec::with_capacity(result_limit);

        for display_row in 0..sheet.view.visible_row_count {
            let Some(source_row) = sheet.resolve_source_row(display_row) else {
                continue;
            };

            for col_index in 0..sheet.total_cols {
                let value = cell_value_at(sheet, source_row, col_index);

                if value.is_empty() || !value.to_lowercase().contains(&normalized_query) {
                    continue;
                }

                matches.push(SpreadsheetSearchMatch {
                    display_row,
                    source_row,
                    col: col_index,
                    value,
                });

                if matches.len() >= result_limit {
                    return Ok(matches);
                }
            }
        }

        Ok(matches)
    })
    .await
    .map_err(|error| format!("failed to search worksheet: {error}"))?
}

#[tauri::command]
async fn get_active_column_value_summary(
    state: tauri::State<'_, SpreadsheetState>,
    col_index: usize,
    query: Option<String>,
    limit: usize,
) -> Result<Vec<SpreadsheetColumnValueSummary>, String> {
    let normalized_query = query
        .map(|value| value.trim().to_lowercase())
        .filter(|value| !value.is_empty());
    let runtime = state.0.clone();

    tauri::async_runtime::spawn_blocking(
        move || -> Result<Vec<SpreadsheetColumnValueSummary>, String> {
            let runtime = runtime
                .lock()
                .map_err(|_| "failed to acquire spreadsheet state".to_string())?;
            let sheet = runtime
                .active_sheet()
                .ok_or_else(|| "no Excel workbook has been loaded".to_string())?;

            if col_index >= sheet.total_cols {
                return Err("requested column is out of range".to_string());
            }

            let result_limit = limit.clamp(1, 200);
            let mut counts: HashMap<String, usize> = HashMap::new();

            for display_row in 0..sheet.view.visible_row_count {
                let Some(source_row) = sheet.resolve_source_row(display_row) else {
                    continue;
                };

                let value = cell_value_at(sheet, source_row, col_index);
                let normalized_value = value.trim();
                if normalized_value.is_empty() {
                    continue;
                }

                if let Some(query_value) = normalized_query.as_deref() {
                    if !normalized_value.to_lowercase().contains(query_value) {
                        continue;
                    }
                }

                *counts.entry(normalized_value.to_string()).or_insert(0) += 1;
            }

            let mut summaries: Vec<SpreadsheetColumnValueSummary> = counts
                .into_iter()
                .map(|(value, count)| SpreadsheetColumnValueSummary { value, count })
                .collect();

            summaries.sort_by(|left, right| {
                right
                    .count
                    .cmp(&left.count)
                    .then_with(|| left.value.cmp(&right.value))
            });
            summaries.truncate(result_limit);

            Ok(summaries)
        },
    )
    .await
    .map_err(|error| format!("failed to load column value summary: {error}"))?
}

#[tauri::command]
async fn get_active_cell_values(
    state: tauri::State<'_, SpreadsheetState>,
    cells: Vec<SpreadsheetCellRef>,
) -> Result<Vec<SpreadsheetCellUpdate>, String> {
    let runtime = state.0.clone();

    tauri::async_runtime::spawn_blocking(move || -> Result<Vec<SpreadsheetCellUpdate>, String> {
        let runtime = runtime
            .lock()
            .map_err(|_| "failed to acquire spreadsheet state".to_string())?;
        let sheet = runtime
            .active_sheet()
            .ok_or_else(|| "no Excel workbook has been loaded".to_string())?;

        let mut values = Vec::with_capacity(cells.len());

        for cell in cells {
            if cell.source_row >= sheet.total_rows || cell.col_index >= sheet.total_cols {
                continue;
            }

            values.push(SpreadsheetCellUpdate {
                source_row: cell.source_row,
                col_index: cell.col_index,
                value: cell_value_at(sheet, cell.source_row, cell.col_index),
            });
        }

        Ok(values)
    })
    .await
    .map_err(|error| format!("failed to read worksheet cells: {error}"))?
}

#[tauri::command]
fn get_data_chunk(
    state: tauri::State<'_, SpreadsheetState>,
    start_row: usize,
    row_limit: usize,
    start_col: usize,
    col_limit: usize,
) -> Result<SpreadsheetChunkResponse, String> {
    let runtime = state
        .0
        .lock()
        .map_err(|_| "failed to acquire spreadsheet state".to_string())?;
    let sheet = runtime
        .active_sheet()
        .ok_or_else(|| "no Excel workbook has been loaded".to_string())?;

    if sheet.view.visible_row_count == 0 || sheet.total_cols == 0 {
        return Ok(SpreadsheetChunkResponse {
            start_row: 0,
            row_limit: 0,
            start_col: 0,
            col_limit: 0,
            total_rows: sheet.view.visible_row_count,
            source_total_rows: sheet.total_rows,
            total_cols: sheet.total_cols,
            header_row_enabled: sheet.view.header_row_enabled,
            headers: Vec::new(),
            source_rows: Vec::new(),
            rows: Vec::new(),
        });
    }

    if sheet.range.is_empty() {
        return Err("no Excel workbook has been loaded".to_string());
    }

    let total_rows = sheet.view.visible_row_count;
    let total_cols = sheet.total_cols;
    let start_row = start_row.min(total_rows);
    let end_row = start_row.saturating_add(row_limit).min(total_rows);
    let start_col = start_col.min(total_cols);
    let end_col = start_col.saturating_add(col_limit).min(total_cols);
    let headers = sheet.visible_headers(start_col, end_col.saturating_sub(start_col));
    let mut source_rows = Vec::with_capacity(end_row.saturating_sub(start_row));
    let mut rows = Vec::with_capacity(end_row.saturating_sub(start_row));

    for display_row in start_row..end_row {
        let Some(source_row) = sheet.resolve_source_row(display_row) else {
            continue;
        };

        source_rows.push(source_row);
        let mut row = Vec::with_capacity(end_col.saturating_sub(start_col));

        for col_index in start_col..end_col {
            let value = cell_value_at(sheet, source_row, col_index);
            row.push(value);
        }

        rows.push(row);
    }

    Ok(SpreadsheetChunkResponse {
        start_row,
        row_limit: rows.len(),
        start_col,
        col_limit: headers.len(),
        total_rows,
        source_total_rows: sheet.total_rows,
        total_cols,
        header_row_enabled: sheet.view.header_row_enabled,
        headers,
        source_rows,
        rows,
    })
}

#[tauri::command]
async fn set_active_cell_value(
    state: tauri::State<'_, SpreadsheetState>,
    source_row: usize,
    col_index: usize,
    value: String,
) -> Result<SpreadsheetLoadResult, String> {
    let runtime = state.0.clone();

    tauri::async_runtime::spawn_blocking(move || -> Result<SpreadsheetLoadResult, String> {
        let mut runtime = runtime
            .lock()
            .map_err(|_| "failed to acquire spreadsheet state".to_string())?;

        let sheet = runtime
            .active_sheet_mut()
            .ok_or_else(|| "no Excel workbook has been loaded".to_string())?;

        if source_row >= sheet.total_rows {
            return Err("requested source row is out of range".to_string());
        }

        if col_index >= sheet.total_cols {
            return Err("requested column is out of range".to_string());
        }

        sheet.edits
            .insert((source_row, col_index), parse_cell_input(&value));
        rebuild_sheet_view(sheet);

        let file_name = runtime
            .file_name
            .clone()
            .unwrap_or_else(|| "workbook.xlsx".to_string());
        build_load_result(&file_name, runtime.active_sheet_index, &runtime.sheets)
    })
    .await
    .map_err(|error| format!("failed to update worksheet cell: {error}"))?
}

#[tauri::command]
async fn set_active_cell_values(
    state: tauri::State<'_, SpreadsheetState>,
    updates: Vec<SpreadsheetCellUpdate>,
) -> Result<SpreadsheetLoadResult, String> {
    let runtime = state.0.clone();

    tauri::async_runtime::spawn_blocking(move || -> Result<SpreadsheetLoadResult, String> {
        let mut runtime = runtime
            .lock()
            .map_err(|_| "failed to acquire spreadsheet state".to_string())?;

        let sheet = runtime
            .active_sheet_mut()
            .ok_or_else(|| "no Excel workbook has been loaded".to_string())?;

        for update in updates {
            if update.source_row >= sheet.total_rows || update.col_index >= sheet.total_cols {
                continue;
            }

            sheet.edits.insert(
                (update.source_row, update.col_index),
                parse_cell_input(&update.value),
            );
        }

        rebuild_sheet_view(sheet);

        let file_name = runtime
            .file_name
            .clone()
            .unwrap_or_else(|| "workbook.xlsx".to_string());
        build_load_result(&file_name, runtime.active_sheet_index, &runtime.sheets)
    })
    .await
    .map_err(|error| format!("failed to update worksheet cells: {error}"))?
}

fn stop_sidecar(sidecar_state: &BackendSidecarState) {
    let (child, child_pid) = if let Ok(mut runtime) = sidecar_state.0.lock() {
        runtime.shutting_down = true;
        (runtime.child.take(), runtime.child_pid.take())
    } else {
        (None, None)
    };

    if let Some(child) = child {
        let _ = child.kill();
    }

    thread::sleep(Duration::from_millis(180));
    terminate_sidecar_process_tree(child_pid);
}

#[cfg(target_os = "windows")]
fn terminate_orphan_sidecars() {
    let _ = Command::new("taskkill")
        .args(["/F", "/T", "/IM", "flexbox-backend.exe"])
        .output();
}

#[cfg(not(target_os = "windows"))]
fn terminate_orphan_sidecars() {}

#[cfg(target_os = "windows")]
fn terminate_sidecar_process_tree(pid: Option<u32>) {
    if let Some(pid) = pid {
        let _ = Command::new("taskkill")
            .args(["/F", "/T", "/PID", &pid.to_string()])
            .output();
        thread::sleep(Duration::from_millis(120));
    }

    terminate_orphan_sidecars();
}

#[cfg(not(target_os = "windows"))]
fn terminate_sidecar_process_tree(_pid: Option<u32>) {}

#[tauri::command]
fn shutdown_application(
    app_handle: tauri::AppHandle,
    sidecar_state: tauri::State<'_, BackendSidecarState>,
) -> Result<(), String> {
    stop_sidecar(&sidecar_state);
    app_handle.exit(0);
    Ok(())
}

fn spawn_backend_sidecar(app_handle: tauri::AppHandle) -> Result<(), String> {
    {
        let state = app_handle.state::<BackendSidecarState>();
        let runtime = state
            .0
            .lock()
            .map_err(|_| "failed to inspect sidecar state".to_string())?;

        if runtime.shutting_down || runtime.child.is_some() {
            return Ok(());
        }
    }

    let sidecar_command = app_handle
        .shell()
        .sidecar("flexbox-backend")
        .map_err(|error| error.to_string())?
        .args(["--host", "127.0.0.1", "--port", "8000"]);

    let (mut rx, child) = sidecar_command
        .spawn()
        .map_err(|error| format!("failed to spawn backend sidecar: {error}"))?;

    {
        let state = app_handle.state::<BackendSidecarState>();
        let mut runtime = state
            .0
            .lock()
            .map_err(|_| "failed to store sidecar".to_string())?;
        runtime.shutting_down = false;
        runtime.child_pid = Some(child.pid());
        runtime.child = Some(child);
        runtime.restart_attempts = 0;
    }

    tauri::async_runtime::spawn(async move {
        while let Some(event) = rx.recv().await {
            match event {
                CommandEvent::Stdout(line) => {
                    println!("[backend] {}", String::from_utf8_lossy(&line));
                }
                CommandEvent::Stderr(line) => {
                    eprintln!("[backend] {}", String::from_utf8_lossy(&line));
                }
                CommandEvent::Terminated(payload) => {
                    eprintln!("[backend] sidecar terminated: {:?}", payload);

                    if let Ok(mut runtime) = app_handle.state::<BackendSidecarState>().0.lock() {
                        runtime.child = None;
                        runtime.child_pid = None;
                    }

                    schedule_sidecar_restart(app_handle.clone());
                    break;
                }
                _ => {}
            }
        }
    });

    Ok(())
}

fn schedule_sidecar_restart(app_handle: tauri::AppHandle) {
    let retry_delay = {
        let state = app_handle.state::<BackendSidecarState>();
        let mut runtime = match state.0.lock() {
            Ok(runtime) => runtime,
            Err(_) => return,
        };

        if runtime.shutting_down {
            return;
        }

        runtime.restart_attempts = runtime.restart_attempts.saturating_add(1);
        let attempt = runtime.restart_attempts.min(5);
        Duration::from_millis(600 * u64::from(attempt))
    };

    thread::spawn(move || {
        thread::sleep(retry_delay);

        if let Err(error) = spawn_backend_sidecar(app_handle.clone()) {
            eprintln!("[backend] restart failed: {error}");
            schedule_sidecar_restart(app_handle);
        }
    });
}

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_shell::init())
        .manage(BackendSidecarState(Arc::new(Mutex::new(
            BackendSidecarRuntime::default(),
        ))))
        .manage(SpreadsheetState(Arc::new(Mutex::new(
            SpreadsheetRuntime::default(),
        ))))
        .setup(|app| {
            terminate_orphan_sidecars();
            spawn_backend_sidecar(app.handle().clone())?;
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            save_theme_preference,
            load_theme_preference,
            save_workspace_session,
            load_workspace_session,
            clear_workspace_session,
            shutdown_application,
            load_excel_file,
            set_active_sheet,
            configure_active_sheet_view,
            set_active_cell_value,
            set_active_cell_values,
            find_in_active_sheet,
            get_active_column_value_summary,
            get_active_cell_values,
            get_data_chunk
        ])
        .build(tauri::generate_context!())
        .expect("failed to build FLEXBOX Tauri shell")
        .run(|app_handle, event| {
            match event {
                tauri::RunEvent::Exit => {
                    let state = app_handle.state::<BackendSidecarState>();
                    stop_sidecar(&state);
                }
                tauri::RunEvent::ExitRequested { .. } => {
                    let state = app_handle.state::<BackendSidecarState>();
                    stop_sidecar(&state);
                }
                tauri::RunEvent::WindowEvent { event, .. } => {
                    if matches!(
                        event,
                        tauri::WindowEvent::CloseRequested { .. } | tauri::WindowEvent::Destroyed
                    ) {
                        let state = app_handle.state::<BackendSidecarState>();
                        stop_sidecar(&state);
                    }
                }
                _ => {}
            }
        });
}
