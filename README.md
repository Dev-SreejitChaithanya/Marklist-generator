
# 🧾 Marklist Generator (Electron App)

## Overview
**Marklist Generator** is a desktop application built with **Electron + Node.js** that automatically generates formatted Word marklists (`.docx`) for students from an Excel file using a Word template.  
Designed for training/educational institutes, it simplifies certificate mark generation and maintains consistent document formatting.

---

## ✨ Features
- Imports student data directly from Excel (`.xlsx / .xls`)
- Fills a Word template while preserving design
- Custom fields for **Course Name**, **Duration**, and **Date**
- Generates one `.docx` per candidate automatically
- Simple Electron-based UI
- Customizable app name and logo

---

## ⚙️ Setup Instructions

```bash
# 1. Install dependencies
npm install

# 2. Run app in dev mode
npm run dev

# 3. Build executable
npm run build
```

---

## 📘 Excel Format

Required columns:
```
Certificate Number, Institute Roll No, Name Of Candidate, Module Test, 
Module Practical, Course Viva, Assignment, Mini Project, Main Project, 
Main Exam, Total
```

---

## 🧩 Template Placeholders

Use the following tags in your Word template  (`template.docx`,name the document template.docx only),check `template_copy.docx`:

```
{{CertificateNo}}, {{Date}}, {{Name}}, {{RollNo}}, {{Duration}}, {{CourseName}}, 
{{ModuleTest}}, {{ModulePractical}}, {{CourseViva}}, {{Assignment}}, 
{{MiniProject}}, {{MainProject}}, {{MainExam}}, {{Total}}
```

> 💡 Make sure each placeholder is typed manually (no copy-paste). Word sometimes breaks tags across multiple XML runs.

---

## 🎨 Customization

- Change app name in `package.json`
- Update window title in `main.js`
- Add your custom icons in `/assets` folder
- Use `.ico` for Windows, `.icns` for macOS

---

## 🧰 Troubleshooting

| Issue | Solution |
|-------|-----------|
| Duplicate close tag | Retype `{{tags}}` manually in Word |
| Default Electron icon | Ensure `icon: "assets/icon.ico"` in `BrowserWindow` |
| Excel header mismatch | Verify exact column names |

---
