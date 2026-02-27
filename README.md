# PESU Academy Resource Downloader

Download course materials from PESU Academy with automatic conversion and merging.

## Quick Start

### 1. Clone the Repository
```bash
git clone https://github.com/ilb225112/pesu_course_downloader.git
cd pesu_course_downloader
```

### Create a Virtual Environment (Recommended)
> **Required:** Python 3.11 or 3.12 (3.13+ not supported due to `windows-curses`)

**Windows:**
```bash
py -3.12 -m venv venv
venv\Scripts\activate
```

**Linux/Mac:**
```bash
python3.12 -m venv venv
source venv/bin/activate
```


### 2. Install Dependencies
```bash
pip install -r requirements.txt
```


### 3. Setup Credentials
To `.env` add your credentials:


### 4. Run the Downloader
```bash
python interactive_download.py
```

### 🐧 Linux / Ubuntu Users
```bash
python3 interactive_download.py
```

---

## 📄 What Each Script Does

### **`interactive_download.py`**  (Main Script) :<br>

  Complete interactive workflow: <br>
  -  Login → Select Course → Select Units → Download → Convert PPTX/DOCX to PDF → Detect & Remove Duplicates → Merge PDFs → Cleanup.
  -  Includes automatic corruption repair for damaged files. **This is the only file you need to run.**
---

### **`pdf_dedup.py`** (Auto-loaded by main script) :
Detects and removes duplicate PDFs after conversion.
- Skips files with different sizes instantly (zero cost)
- Uses perceptual hashing (pHash) only on same-size candidates
- Prompts before deleting — or pass `auto_delete=True` for batch use

**Requires:** `pip install pymupdf`

##  Notes

- **Windows users:** PowerPoint COM provides best conversion quality (requires MS Office installed)
- **Cross-platform:** Use Aspose.Slides or LibreOffice as fallback
- Files are numbered sequentially within each unit for easy merging
- Empty files and temporary data are automatically cleaned up

---

## 🤝 Contributing

Contributions, issues, and feature requests are welcome! Feel free to fork and submit a PR.

