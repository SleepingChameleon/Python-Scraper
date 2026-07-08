# ⚡ ProductScraper

A Python desktop utility with a drag-and-drop UI for converting **PDF → Excel**, **Excel → JSON**, and **Image → JSON** (via Tesseract OCR). Built with Tkinter and designed for offline use.

---

## Features

| Tab | Input | Output | Engine |
|-----|-------|--------|--------|
| 📄 PDF → Excel | `.pdf` | `.xlsx` | pdfplumber |
| 📊 Excel → JSON | `.xlsx` / `.xls` | `.json` | pandas |
| 🖼️ Image → JSON | `.jpg` `.png` `.webp` `.gif` | `.json` | Tesseract OCR |

- Drag & drop or click-to-browse for all file types
- Syntax-highlighted JSON preview with one-click copy
- Activity log for every action
- Output files are saved automatically next to the source file
- Fully offline — no internet required

---

## Requirements

### Python Packages

```bash
pip install pdfplumber pandas openpyxl tkinterdnd2 pytesseract Pillow
```

### Tesseract OCR (for Image → JSON tab only)

1. Download and install from: https://github.com/UB-Mannheim/tesseract/wiki
2. Add Tesseract to your system PATH:
   - Default install path: `C:\Program Files\Tesseract-OCR`
   - Windows: Search **Environment Variables** → Edit `Path` → Add the path above
   - Restart your terminal after adding it

### Python Version

Python 3.10 or higher recommended.

---

## Installation

```bash
# 1. Clone or download the project
git clone https://github.com/SleepingChameleon/ConverterKit.git
cd ConverterKit

# 2. Create and activate a virtual environment (recommended)
python -m venv venv

# Windows
venv\Scripts\activate

# Mac/Linux
source venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the app
py app.py
```

---

## Usage

### 📄 PDF → Excel

1. Go to the **PDF → Excel** tab
2. Drop a `.pdf` file onto the drop zone (or click to browse)
3. Click **Convert to Excel**
4. Output saved as `yourfile_converted.xlsx` in the same folder

> Sheet names follow the format `P1_T1` meaning **Page 1, Table 1**. If no tables are detected, raw text is extracted as a single column.

---

### 📊 Excel → JSON

1. Go to the **Excel → JSON** tab
2. Drop an `.xlsx` or `.xls` file onto the drop zone
3. Adjust the **Preview rows** count if needed
4. Click **Convert to JSON**
5. Output saved as `yourfile_body.json` in the same folder

> Each sheet in the workbook becomes a key in the JSON object, with its rows as an array of objects.

**Example output:**
```json
{
  "Sheet1": [
    { "Name": "Rice", "Price": "45.00", "Unit": "kg" },
    { "Name": "Corn", "Price": "32.00", "Unit": "kg" }
  ]
}
```

---

### 🖼️ Image → JSON

1. Go to the **Image → JSON** tab
2. Drop an image (`.jpg`, `.png`, `.webp`, `.gif`) onto the drop zone
3. Select an **OCR mode** from the dropdown:
   - `PSM 6` — Best for standard uniform tables (default)
   - `PSM 4` — Best for single-column layouts
   - `PSM 11` — Best for mixed or sparse layouts
4. Click **Extract with Tesseract OCR**
5. Output saved as `yourimage_extracted.json` in the same folder

> For best results, use clear, high-resolution images with dark text on a light background. The first row of the detected table is used as the JSON keys (column headers).

**Example output:**
```json
{
  "table": [
    { "Commodity": "Bangus", "Retail Price": "180.00", "Unit": "kg" },
    { "Commodity": "Tilapia", "Retail Price": "120.00", "Unit": "kg" }
  ]
}
```

---

## Project Structure

```
ConverterKit/
├── app.py              # Main application
├── requirements.txt    # Python dependencies
└── README.md           # This file
```

---

## Dependencies

| Package | Purpose |
|---------|---------|
| `tkinter` | GUI framework (built into Python) |
| `tkinterdnd2` | Drag-and-drop support |
| `pdfplumber` | PDF table extraction |
| `pandas` | Excel parsing and data handling |
| `openpyxl` | Excel file writing |
| `pytesseract` | Python wrapper for Tesseract OCR |
| `Pillow` | Image processing and thumbnail preview |

---

## Troubleshooting

**`ModuleNotFoundError: No module named 'tkinterdnd2'`**
```bash
pip install tkinterdnd2
```

**`ModuleNotFoundError: No module named 'pdfplumber'`**
```bash
pip install pdfplumber
```

**`pytesseract.pytesseract.TesseractNotFoundError`**
- Tesseract is not installed or not in PATH
- Download from: https://github.com/UB-Mannheim/tesseract/wiki
- Add `C:\Program Files\Tesseract-OCR` to your system PATH and restart the terminal

**Image OCR returns garbled or missing text**
- Try a different PSM mode from the dropdown
- Use a higher resolution image
- Make sure the table has clear borders and dark text on a light background

**No tables found in PDF**
- The PDF may contain scanned images instead of real text
- Try converting the PDF to an image first, then use the Image → JSON tab

---

## Author

**Rey** — [@SleepingChameleon](https://github.com/SleepingChameleon)

---

## License

MIT License — free to use, modify, and distribute.
