#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use calamine::{open_workbook_auto, Data, Range, Reader};
use serde::{de::DeserializeOwned, Deserialize, Serialize};
use std::{
    fs,
    path::{Path, PathBuf},
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
    shutting_down: bool,
    restart_attempts: u32,
}

struct SpreadsheetState(Arc<Mutex<SpreadsheetRuntime>>);

struct SheetRuntime {
    name: String,
    range: Range<Data>,
    total_rows: usize,
    total_cols: usize,
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
    total_cols: usize,
    sheets: Vec<SpreadsheetSheetSummary>,
}

#[derive(Debug, Serialize)]
struct SpreadsheetChunkResponse {
    start_row: usize,
    row_limit: usize,
    start_col: usize,
    col_limit: usize,
    total_rows: usize,
    total_cols: usize,
    rows: Vec<Vec<String>>,
}

#[derive(Debug, Serialize)]
struct SpreadsheetSearchMatch {
    row: usize,
    col: usize,
    value: String,
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

impl SpreadsheetRuntime {
    fn active_sheet(&self) -> Option<&SheetRuntime> {
        self.sheets.get(self.active_sheet_index)
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
        total_rows: active_sheet.total_rows,
        total_cols: active_sheet.total_cols,
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

                sheets.push(SheetRuntime {
                    name: sheet_name,
                    range,
                    total_rows,
                    total_cols,
                });
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
fn find_in_active_sheet(
    state: tauri::State<'_, SpreadsheetState>,
    query: String,
    limit: usize,
) -> Result<Vec<SpreadsheetSearchMatch>, String> {
    let normalized_query = query.trim().to_lowercase();

    if normalized_query.is_empty() {
        return Ok(Vec::new());
    }

    let runtime = state
        .0
        .lock()
        .map_err(|_| "failed to acquire spreadsheet state".to_string())?;
    let sheet = runtime
        .active_sheet()
        .ok_or_else(|| "no Excel workbook has been loaded".to_string())?;
    let result_limit = limit.clamp(1, 200);
    let mut matches = Vec::with_capacity(result_limit);

    for row_index in 0..sheet.total_rows {
        for col_index in 0..sheet.total_cols {
            let value = sheet
                .range
                .get((row_index, col_index))
                .map(serialize_cell)
                .unwrap_or_default();

            if value.is_empty() || !value.to_lowercase().contains(&normalized_query) {
                continue;
            }

            matches.push(SpreadsheetSearchMatch {
                row: row_index,
                col: col_index,
                value,
            });

            if matches.len() >= result_limit {
                return Ok(matches);
            }
        }
    }

    Ok(matches)
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

    if sheet.total_rows == 0 || sheet.total_cols == 0 {
        return Ok(SpreadsheetChunkResponse {
            start_row: 0,
            row_limit: 0,
            start_col: 0,
            col_limit: 0,
            total_rows: sheet.total_rows,
            total_cols: sheet.total_cols,
            rows: Vec::new(),
        });
    }

    if sheet.range.is_empty() {
        return Err("no Excel workbook has been loaded".to_string());
    }

    let total_rows = sheet.total_rows;
    let total_cols = sheet.total_cols;
    let start_row = start_row.min(total_rows);
    let end_row = start_row.saturating_add(row_limit).min(total_rows);
    let start_col = start_col.min(total_cols);
    let end_col = start_col.saturating_add(col_limit).min(total_cols);
    let mut rows = Vec::with_capacity(end_row.saturating_sub(start_row));

    for row_index in start_row..end_row {
        let mut row = Vec::with_capacity(end_col.saturating_sub(start_col));

        for col_index in start_col..end_col {
            let value = sheet
                .range
                .get((row_index, col_index))
                .map(serialize_cell)
                .unwrap_or_default();
            row.push(value);
        }

        rows.push(row);
    }

    Ok(SpreadsheetChunkResponse {
        start_row,
        row_limit: end_row.saturating_sub(start_row),
        start_col,
        col_limit: end_col.saturating_sub(start_col),
        total_rows,
        total_cols,
        rows,
    })
}

fn stop_sidecar(sidecar_state: &BackendSidecarState) {
    if let Ok(mut runtime) = sidecar_state.0.lock() {
        runtime.shutting_down = true;
        if let Some(child) = runtime.child.take() {
            let _ = child.kill();
        }
    }
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
            spawn_backend_sidecar(app.handle().clone())?;
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            save_theme_preference,
            load_theme_preference,
            save_workspace_session,
            load_workspace_session,
            clear_workspace_session,
            load_excel_file,
            set_active_sheet,
            find_in_active_sheet,
            get_data_chunk
        ])
        .build(tauri::generate_context!())
        .expect("failed to build FLEXBOX Tauri shell")
        .run(|app_handle, event| {
            if let tauri::RunEvent::Exit = event {
                let state = app_handle.state::<BackendSidecarState>();
                stop_sidecar(&state);
            }
        });
}
