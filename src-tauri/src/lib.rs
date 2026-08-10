use std::{fs, path::{Path, PathBuf}, process::Command};

use tauri_plugin_dialog::DialogExt;

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

pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![pick_directory, write_pdf, read_pdf, delete_pdfs, open_folder])
        .run(tauri::generate_context!())
        .expect("error while running Code Report Tracker");
}
