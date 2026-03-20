# ID-RSS — Intelligent Document Retrieval & Structuring System

> **HackIndia 2026 · Team ARISE · KCC Institute**

🔗 **Live Demo:** https://Ayushh-Sharmaa.github.io/ID-RSS-fixed

---

## What it does

ID-RSS reads `.docx` student admission files, lets you type which fields to extract (e.g. `Name`, `Serial No.`), pulls them out of all files using regex, shows results in a live animated table, then exports to Excel.

## Tech Stack

| Layer | Tech |
|---|---|
| Backend | Python + Flask |
| .docx parsing | python-docx |
| Excel export | openpyxl |
| Field extraction | regex / re |
| Frontend | HTML + CSS + JS (no React) |
| GitHub Pages | JSZip + SheetJS (browser version) |

## Project Structure

```
ID-RSS-fixed/
├── index.html          ← GitHub Pages demo (browser-only, no server needed)
├── App.py              ← Flask server + routes
├── Extractor.py        ← reads .docx, regex extraction
├── Exporter.py         ← writes styled 3-sheet Excel
├── gen_demo.py         ← generates 30 demo student files
├── templates/
│   └── index.html      ← Flask template (same UI, uses /extract API)
└── demo_files/
    └── student_0001.docx … student_0030.docx
```

## Run Locally (Full Flask Version)

```bash
# 1. Clone
git clone https://github.com/Ayushh-Sharmaa/ID-RSS-fixed.git
cd ID-RSS-fixed

# 2. Install dependencies
pip install flask python-docx openpyxl

# 3. Generate demo files (optional, already included)
python gen_demo.py

# 4. Start server
python App.py

# 5. Open browser
# http://localhost:5000
```

## How Extraction Works

Each `.docx` file contains fields in this exact format anywhere in the document:
```
Name- Tanishk Bansal
Serial No.- STU-2024-0001
```

The regex pattern used:
```python
r"Label\s*-\s*(.+?)(?:\r?\n|$)"
```

## Team ARISE

| Role | Responsibility |
|---|---|
| Member 1 — Extraction Engineer | `Extractor.py` |
| Member 2 — Data Structure Engineer | `Exporter.py` + UI |
| Member 3 — Full Stack Engineer | `App.py` + `templates/index.html` |

---

*Built at HackIndia 2026, KCC Institute NCR East Region*
