# PESU Academy Resource Downloader

Download course materials from PESU Academy with automatic conversion and merging.

## Quick Start

### 1. Clone the Repository
```bash
git clone https://github.com/ilb225112/pesu_course_downloader.git
cd pesu_course_downloader
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

**OR RUN** 
```bash
pip install -r requirements.txt 
```

### 3. Setup Credentials
Copy `.env.example` to `.env` and add your credentials:
```bash
cp .env.example .env
```
Then edit `.env` with your PESU Academy SRN and password.

### 4. Run the Downloader
```bash
python interactive_download.py
```

---

## 📄 What Each Script Does

### **`interactive_download.py`**  (Main Script) :<br>

  Complete interactive workflow: <br>
  - Login → Select Course → Select Units → Download → Convert PPTX/DOCX to PDF → Merge PDFs → Cleanup.
  -  Includes automatic corruption repair for damaged files. **This is the only file you need to run.**
---


##  Notes

- **Windows users:** PowerPoint COM provides best conversion quality (requires MS Office installed)
- **Cross-platform:** Use Aspose.Slides or LibreOffice as fallback
- Files are numbered sequentially within each unit for easy merging
- Empty files and temporary data are automatically cleaned up

---

## 🤝 Contributing

Contributions, issues, and feature requests are welcome! Feel free to fork and submit a PR.

