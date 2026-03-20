# ID-RSS — Intelligent Document Retrieval & Structuring System

**HackIndia 2026 · Team ARISE · KCC Institute**

🔗 **Live Demo:** https://Ayushh-Sharmaa.github.io/ID-RSS-fixed

---

## What it does

Upload `.docx` (or any document), type field names like `Name` or `Serial No.`, and ID-RSS extracts them from every file using regex — showing results live row by row, then exporting to Excel.

## Run locally (Python/Flask)

```bash
git clone https://github.com/Ayushh-Sharmaa/ID-RSS-fixed.git
cd ID-RSS-fixed
pip install flask python-docx openpyxl
python App.py
# Open http://localhost:5000
```

## Structure

```
ID-RSS-fixed/
├── index.html        ← GitHub Pages (works in browser, no server)
├── App.py            ← Flask server
├── Extractor.py      ← regex field extraction
├── Exporter.py       ← styled 3-sheet Excel export
├── gen_demo.py       ← generates 30 demo .docx files
├── templates/
│   └── index.html    ← Flask template
└── demo_files/       ← 30 sample student admission files
```

## Team ARISE

| Member | Role | File |
|--------|------|------|
| Member 1 | Extraction Engineer | `Extractor.py` |
| Member 2 | Data Structure Engineer | `Exporter.py` + UI |
| Member 3 | Full Stack Engineer | `App.py` + `templates/` |

---
*Built at HackIndia 2026 · NCR East Region*
