const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");
const fs = require("fs");
const xlsx = require("xlsx");

// Libraries for template-based docx generation
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

function createWindow() {
  const win = new BrowserWindow({
    width: 820,
    height: 520,
    title: "NSDC Marklist Generator",
    icon: path.join(__dirname, "assets", "logo.ico"),
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  win.loadFile(path.join(__dirname, "renderer", "index.html"));
}

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

// ✅ Open file dialog
ipcMain.handle("dialog:openFile", async () => {
  return await dialog.showOpenDialog({
    title: "Select Excel file",
    properties: ["openFile"],
    filters: [{ name: "Excel", extensions: ["xlsx", "xls"] }]
  });
});

// ✅ Open folder dialog
ipcMain.handle("dialog:openFolder", async () => {
  return await dialog.showOpenDialog({
    title: "Select output folder",
    properties: ["openDirectory"]
  });
});

// ✅ Helper: Fill template.docx
function generateFromTemplate(templatePath, outputPath, data) {
  const content = fs.readFileSync(templatePath, "binary");
  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true
  });

  try {
    // 👇 Try rendering placeholders with data
    doc.render(data);
  } catch (error) {
    // 👇 Debug block: shows exactly which tag caused the issue
    console.log(
      JSON.stringify(
        {
          name: error.name,
          message: error.message,
          properties: error.properties
        },
        null,
        2
      )
    );
    throw error; // Re-throw so the app still shows the error
  }

  const buf = doc.getZip().generate({ type: "nodebuffer" });
  fs.writeFileSync(outputPath, buf);
}


// ✅ Generate marklists from Excel
ipcMain.handle("generate-marklists", async (event, { excelPath, outputFolder, courseName, courseDuration, dateStr }) => {
  try {
    const workbook = xlsx.readFile(excelPath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    if (!rows.length) throw new Error("Excel file is empty.");

    const requiredCols = [
      "Certificate Number",
      "Institute Roll No",
      "Name Of Candidate",
      "Module Test",
      "Module Practical",
      "Course Viva",
      "Assignment",
      "Mini Project",
      "Main Project",
      "Main Exam",
      "Total"
    ];

    const headers = Object.keys(rows[0]);
    const missing = requiredCols.filter(col => !headers.includes(col));
    if (missing.length) throw new Error("Missing columns: " + missing.join(", "));

    // Template path
    const templatePath = path.join(__dirname, "template.docx");
    if (!fs.existsSync(templatePath)) {
      throw new Error("Missing template.docx in project folder.");
    }

    let count = 0;
    for (const row of rows) {
      const data = {
        CertificateNo: row["Certificate Number"],
        Name: row["Name Of Candidate"],
        RollNo: row["Institute Roll No"],
        Duration: courseDuration,
        CourseName: courseName,
        Date: dateStr,
        ModuleTest: `${Math.round(row["Module Test"])}`,
        ModulePractical: `${Math.round(row["Module Practical"])}`,
        CourseViva: `${Math.round(row["Course Viva"])}`,
        Assignment: `${Math.round(row["Assignment"])}`,
        MiniProject: `${Math.round(row["Mini Project"])}`,
        MainProject: `${Math.round(row["Main Project"])}`,
        MainExam: `${Math.round(row["Main Exam"])}`,
        Total: `${Math.round(row["Total"])}`
      };

      const safeName = `${row["Certificate Number"]}_${row["Name Of Candidate"]}`.replace(/[\\/:*?"<>|]/g, "_");
      const outPath = path.join(outputFolder, `${safeName}.docx`);

      generateFromTemplate(templatePath, outPath, data);
      count++;
    }

    return { success: true, count };
  } catch (err) {
    return { success: false, error: err.message };
  }
});
