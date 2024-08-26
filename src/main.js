const { invoke } = window.__TAURI__.tauri;

let greetInputEl;
let greetMsgEl;

async function greet() {
  // Learn more about Tauri commands at https://tauri.app/v1/guides/features/command
  greetMsgEl.textContent = await invoke("greet", { name: greetInputEl.value });
}


async function load_File() {
  try {
    const lines = await invoke("load_file");
    console.log('Loaded Lines:', lines);  // 파일 내용 확인용 로그
    statusE1.textContent = "C # File Loaded Successfully!";
    return lines;  // Export 버튼에서 사용할 수 있도록 반환
  } catch (error) {
    console.error('Failed to load file:', error);
    statusE1.textContent = 'Failed to Load C# File';
    return [];
  }
}



async function write_to_excel() {
  try{
    const lines = await loadFile();  // Load the file first
    if(lines.length > 0){
      await invoke("write_to_excel", {lines});
      statusE1.textContent = 'Exported to Excel Successfully!';
    } else {
      statusE1.textContent = 'No data to export.';
    }
  } catch (error) {
    console.error('Failed to export to Excel:', error);
    statusE1.textContent = 'Failed to Export to Excel';
  }
}




window.addEventListener("DOMContentLoaded", () => {
  greetInputEl = document.querySelector("#greet-input");
  greetMsgEl = document.querySelector("#greet-msg");
  statusE1 = document.querySelector("#status");


  document.querySelector("#greet-form").addEventListener("submit", (e) => {
    e.preventDefault();
    greet();
  });

  document.querySelector("#load-file").addEventListener("click", loadFile);
  document.querySelector("#export-excel").addEventListener("click", exportToExcel);
});
