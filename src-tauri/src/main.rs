// Prevents additional console window on Windows in release, DO NOT REMOVE!!
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::fs;
use std::path::PathBuf;
use rust_xlsxwriter::{workbook, Workbook, XlsxError};
use tauri::command;



// Learn more about Tauri commands at https://tauri.app/v1/guides/features/command
#[tauri::command]
fn greet(name: &str) -> String {
    format!("Hello, {}! You've been greeted from Jihyun!", name)
}

// C# 파일 Import
#[tauri::command]
fn load_file() -> Result<Vec<String>, String> {
    // 경로 가져오기
    let mut path = dirs::desktop_dir().ok_or("Could not find desktop directory")?;

    // C# 파일 이름으로 경로 추가
    path.push("TC/SeatHeatVent_ThirdRL.cs");

    // 파일 읽기
    let file_content = fs::read_to_string(&path)
        .map_err(|e| format!("Failed to read the file: {}", e))?;


    // 파일 내용을 줄 단위로 벡터에 저장
    let lines: Vec<String> = file_content.lines().map(|line| line.to_string()).collect();

    Ok(lines)
}


// 불러온 C# 코드를 엑셀 파일로 Export
#[command]
fn write_to_excel(lines: Vec<String>) -> Result<(), String> {
    // 엑셀 파일 생성
    let mut path = dirs::desktop_dir().ok_or("Could not find desktop directory")?;
    println!("Excel file path: {:?}", path);

    path.push("/result.xlsx");
    println!("Excel file path before setting file name: {:?}", path);

    // Workbook 생성
    let mut workbook = Workbook::new();

    // 새로운 시트 추가
    let worksheet = workbook.add_worksheet();


    // 벡터에 저장한 코드 한 줄씩 엑셀 파일에 작성
    for (i, line) in lines.iter().enumerate() {
        worksheet.write_string(i as u32, 0, line);
    }

    // 엑셀 파일 저장
    workbook.save(&path.to_str().unwrap()).map_err(|e| format!("Failed to save workbook: {}", e))?;


    Ok(())
}



fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![greet])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
