#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use calamine::{open_workbook_auto, Reader};
use serde::{de::DeserializeOwned, Deserialize, Serialize};
use std::{
    fs,
    path::{Path, PathBuf},
    thread,
    sync::{Arc, Mutex},
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

#[derive(Default)]
struct SpreadsheetRuntime {
    file_path: Option<String>,
    file_name: Option<String>,
    sheet_name: Option<String>,
    rows: Vec<Vec<String>>,
    total_rows: usize,
    total_cols: usize,
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
struct SpreadsheetLoadResult {
    file_name: String,
    sheet_name: String,
    total_rows: usize,
    total_cols: usize,
}

#[derive(Debug, Serialize)]
struct SpreadsheetChunkResponse {
    start_row: usize,
    limit: usize,
    total_rows: usize,
    total_cols: usize,
    rows: Vec<Vec<String>>,
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

#[tauri::command]
fn load_excel_file(
    state: tauri::State<SpreadsheetState>,
    path: String,
) -> Result<SpreadsheetLoadResult, String> {
    let workbook_path = PathBuf::from(&path);

    if !workbook_path.exists() {
        return Err("selected Excel file does not exist".to_string());
    }

    let mut workbook =
        open_workbook_auto(&workbook_path).map_err(|error| format!("failed to open Excel file: {error}"))?;

    let sheet_name = workbook
        .sheet_names()
        .first()
        .cloned()
        .ok_or_else(|| "workbook does not contain any worksheets".to_string())?;

    let range = workbook
        .worksheet_range_at(0)
        .ok_or_else(|| "failed to read first worksheet".to_string())?
        .map_err(|error| format!("failed to parse worksheet: {error}"))?;

    let (total_rows, total_cols) = range.get_size();
    let rows = range
        .rows()
        .map(|row| row.iter().map(|cell| cell.to_string()).collect::<Vec<_>>())
        .collect::<Vec<_>>();

    let file_name = file_name_from_path(&workbook_path);
    let result = SpreadsheetLoadResult {
        file_name: file_name.clone(),
        sheet_name: sheet_name.clone(),
        total_rows,
        total_cols,
    };

    let mut runtime = state
        .0
        .lock()
        .map_err(|_| "failed to acquire spreadsheet state".to_string())?;
    runtime.file_path = Some(workbook_path.to_string_lossy().to_string());
    runtime.file_name = Some(file_name);
    runtime.sheet_name = Some(sheet_name);
    runtime.rows = rows;
    runtime.total_rows = total_rows;
    runtime.total_cols = total_cols;

    Ok(result)
}

#[tauri::command]
fn get_data_chunk(
    state: tauri::State<SpreadsheetState>,
    start_row: usize,
    limit: usize,
) -> Result<SpreadsheetChunkResponse, String> {
    let runtime = state
        .0
        .lock()
        .map_err(|_| "failed to acquire spreadsheet state".to_string())?;

    if runtime.rows.is_empty() {
        return Err("no Excel workbook has been loaded".to_string());
    }

    let total_rows = runtime.total_rows.min(runtime.rows.len());
    let total_cols = runtime.total_cols;
    let start_row = start_row.min(total_rows);
    let end_row = start_row.saturating_add(limit).min(total_rows);
    let rows = runtime.rows[start_row..end_row].to_vec();

    Ok(SpreadsheetChunkResponse {
        start_row,
        limit: end_row.saturating_sub(start_row),
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
        .sidecar("uniframe-backend")
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
            get_data_chunk
        ])
        .build(tauri::generate_context!())
        .expect("failed to build UniFrame Tauri shell")
        .run(|app_handle, event| {
            if let tauri::RunEvent::Exit = event {
                let state = app_handle.state::<BackendSidecarState>();
                stop_sidecar(&state);
            }
        });
}
