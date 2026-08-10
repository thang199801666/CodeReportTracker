use std::{fs, path::{Path, PathBuf}, process::Command};

use tauri_plugin_dialog::DialogExt;

#[derive(serde::Serialize)]
struct ExcelFile {
    path: String,
    data: Vec<u8>,
}

fn safe_name(value: &str) -> String {
    let cleaned: String = value
        .chars()
        .map(|character| if r#"<>:/\\|?*"#.contains(character) { '_' } else { character })
        .collect();
    let trimmed = cleaned.trim();
    if trimmed.is_empty() { "New Tab".to_string() } else { trimmed.to_string() }
}

fn tab_path(root: &str, tab: &str) -> PathBuf {
    Path::new(root).join(safe_name(tab))
}

#[tauri::command]
async fn pick_directory(app: tauri::AppHandle) -> Result<Option<String>, String> {
    Ok(app.dialog().file().blocking_pick_folder().map(|path| path.to_string()))
}

#[tauri::command]
fn write_pdf(root: String, tab: String, file_name: String, data: Vec<u8>) -> Result<(), String> {
    let directory = tab_path(&root, &tab);
    fs::create_dir_all(&directory).map_err(|error| error.to_string())?;
    fs::write(directory.join(safe_name(&file_name)), data).map_err(|error| error.to_string())
}

#[tauri::command]
fn read_pdf(root: String, tab: String, file_name: String) -> Result<Vec<u8>, String> {
    fs::read(tab_path(&root, &tab).join(safe_name(&file_name))).map_err(|error| error.to_string())
}

#[tauri::command]
fn delete_pdfs(root: String, tab: String) -> Result<u32, String> {
    let directory = tab_path(&root, &tab);
    if !directory.exists() { return Ok(0); }
    let mut deleted = 0;
    for entry in fs::read_dir(directory).map_err(|error| error.to_string())? {
        let path = entry.map_err(|error| error.to_string())?.path();
        if path.extension().is_some_and(|extension| extension.eq_ignore_ascii_case("pdf")) {
            fs::remove_file(path).map_err(|error| error.to_string())?;
            deleted += 1;
        }
    }
    Ok(deleted)
}

#[tauri::command]
fn open_folder(root: String, tab: String) -> Result<(), String> {
    let directory = tab_path(&root, &tab);
    fs::create_dir_all(&directory).map_err(|error| error.to_string())?;
    Command::new("explorer.exe")
        .arg(directory)
        .spawn()
        .map(|_| ())
        .map_err(|error| error.to_string())
}

#[tauri::command]
fn save_crp(app: tauri::AppHandle, data: Vec<u8>) -> Result<Option<String>, String> {
    let Some(path) = app
        .dialog()
        .file()
        .add_filter("Code Report", &["crp"])
        .set_file_name("report.crp")
        .blocking_save_file()
    else {
        return Ok(None);
    };
    let path = path.to_string();
    fs::write(&path, data).map_err(|error| error.to_string())?;
    Ok(Some(path))
}

#[tauri::command]
fn save_excel(app: tauri::AppHandle, data: Vec<u8>) -> Result<Option<String>, String> {
    let Some(path) = app
        .dialog()
        .file()
        .add_filter("Excel Workbook", &["xlsx", "xls"])
        .set_file_name("updated-report.xlsx")
        .blocking_save_file()
    else {
        return Ok(None);
    };
    let path = path.to_string();
    fs::write(&path, data).map_err(|error| error.to_string())?;
    Ok(Some(path))
}

#[tauri::command]
fn pick_excel_file(app: tauri::AppHandle) -> Result<Option<ExcelFile>, String> {
    let Some(path) = app
        .dialog()
        .file()
        .add_filter("Excel Workbook", &["xlsx", "xls"])
        .blocking_pick_file()
    else {
        return Ok(None);
    };
    let path = path.to_string();
    let data = fs::read(&path).map_err(|error| error.to_string())?;
    Ok(Some(ExcelFile { path, data }))
}

#[tauri::command]
fn overwrite_file(path: String, data: Vec<u8>) -> Result<(), String> {
    fs::write(path, data).map_err(|error| error.to_string())
}

pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![pick_directory, write_pdf, read_pdf, delete_pdfs, open_folder, save_crp, save_excel, pick_excel_file, overwrite_file])
        .run(tauri::generate_context!())
        .expect("error while running Code Report Tracker");
}
