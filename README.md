# 📊 Excel File Consolidator

Merge multiple Excel / CSV files of **any format** into one clean, formatted spreadsheet.  
A **Source File** column is automatically added to every row so you always know where each record came from.

---

## ✅ Supported File Formats

| Format | Extension | Notes |
|--------|-----------|-------|
| Binary XLS | `.xls` | Classic Excel 97–2003 format |
| HTML-as-XLS | `.xls` | Web-exported tables saved as `.xls` |
| Excel Workbook | `.xlsx`, `.xlsm` | Modern Excel format |
| CSV | `.csv` | Auto-detects encoding (UTF-8, Latin-1, CP1252) |

---

## 🚀 Quick Start

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the app
```bash
streamlit run app.py
```

The app will open at `http://localhost:8501`

---

## 🖥️ How to Use

1. **Upload** — Drag and drop one or more files (any mix of formats)
2. **Sheet selection** — For multi-sheet files, choose which sheet to use
3. **Parse** — Click "Parse All Files" to read and preview each file
4. **Consolidate** — Click "Consolidate All Files" to merge everything
5. **Download** — Export as `.xlsx` (formatted) or `.csv`

---

## ⚙️ Sidebar Options

| Option | Description |
|--------|-------------|
| Output filename | Name of the downloaded file |
| Add File Summary sheet | Adds a second sheet with per-file row counts |
| Mismatched columns | **Union** keeps all columns; **Intersection** keeps only shared columns |
| Add global row number | Inserts a sequential `Row No.` column |

---

## 📤 Output Structure

### Sheet 1 — Consolidated Data
- **Source File** column (leftmost, highlighted in gold) — shows the originating filename for every row
- All data rows from every uploaded file
- Alternating row colours, frozen header row

### Sheet 2 — File Summary *(optional)*
- One row per source file showing row count

---

## 📦 Dependencies

```
streamlit       — web UI framework
pandas          — data manipulation
openpyxl        — read/write .xlsx files
xlrd            — read binary .xls files
beautifulsoup4  — parse HTML-as-XLS files
lxml            — fast HTML parser (used by bs4)
numpy           — numeric support
```
