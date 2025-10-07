const btnBrowseExcel = document.getElementById("btnBrowseExcel");
const excelPathInput = document.getElementById("excelPath");

const btnBrowseOutput = document.getElementById("btnBrowseOutput");
const outputPathInput = document.getElementById("outputPath");

const btnGenerate = document.getElementById("btnGenerate");
const courseNameInput = document.getElementById("courseName");
const courseDurationInput = document.getElementById("courseDuration");
const dateStrInput = document.getElementById("dateStr");
const statusDiv = document.getElementById("status");

// Browse Excel
btnBrowseExcel.addEventListener("click", async () => {
  const res = await window.electronAPI.openFile();
  if (!res.canceled && res.filePaths.length) {
    excelPathInput.value = res.filePaths[0];
  }
});

// Browse Output
btnBrowseOutput.addEventListener("click", async () => {
  const res = await window.electronAPI.openFolder();
  if (!res.canceled && res.filePaths.length) {
    outputPathInput.value = res.filePaths[0];
  }
});

// Generate Marklists
btnGenerate.addEventListener("click", async () => {
  const excelPath = excelPathInput.value.trim();
  const outputFolder = outputPathInput.value.trim();
  const courseName = courseNameInput.value.trim();
  const courseDuration = courseDurationInput.value.trim();
  const dateStr = dateStrInput.value.trim();

  if (!excelPath || !outputFolder || !courseName || !courseDuration || !dateStr) {
    alert("Fill in all fields and select files/folder.");
    return;
  }

  statusDiv.textContent = "Generating...";
  const result = await window.electronAPI.generateMarklists({
    excelPath,
    outputFolder,
    courseName,
    courseDuration,
    dateStr
  });

  if (result.success) {
    statusDiv.textContent = `✅ Generated ${result.count} marklists`;
  } else {
    statusDiv.textContent = `❌ Error: ${result.error}`;
  }
});
