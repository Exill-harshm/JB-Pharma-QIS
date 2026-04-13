# QIS Automation — CTD Module 3 DOCX Generator

## QIS v2 (Table-Driven)

A newer, generic table-driven pipeline is included under `qis_v2/`.

- Entry point: `qis_v2/run.py`
- Core package: `qis_v2/src/qis_api/`
- Uses direct Module 1 QIS table extraction with editable DOCX table fill.

Automatically populates a **Quality Information Summary (QIS)** DOCX template with content extracted from CTD Module 3 source PDFs.

---

## How It Works

```
config.yaml
    │
    ▼
section_mapper.py   ← Scans source PDF folder, maps section numbers to PDF paths
    │
    ▼
docx_builder.py     ← Finds "Refer Section X.X.X" placeholders in the template
    │
    ▼
pdf_extractor.py    ← For each section: extracts text + tables + images from PDF
    │                  Auto-detects and strips repeating headers/footers
    ▼
Output_QIS.docx     ← Fully populated QIS document
```

---

## Project Structure

```
QIS_Automation/
├── main.py             # Entry point — runs the full pipeline
├── config.yaml         # All paths and per-section settings
├── config_loader.py    # Reads and validates config.yaml
├── section_mapper.py   # Maps CTD section numbers → source PDF paths
├── pdf_extractor.py    # Extracts content from each PDF
├── docx_builder.py     # Injects content into the QIS template DOCX
└── logger_setup.py     # Rotating file logger
```

---

## Setup

### Prerequisites

```bash
pip install pymupdf pdf2docx python-docx pyyaml
```

> **Note:** `pdf2docx` will use **Tesseract OCR** automatically for scanned PDFs (image-only pages). Install Tesseract and add it to your PATH if your source PDFs are scanned.

---

## Configuration (`config.yaml`)

```yaml
# Path to the blank QIS template DOCX
template_docx_path:    "D:\\...\\QUALITY INFORMATION SUMMARY.docx"

# Optional: path to a .docx or .pdf listing required sections (leave "" to skip)
mapping_logic_pdf_path: ""

# Folder containing all source CTD Module 3 PDFs (scanned recursively)
source_pdf_folder:     "D:\\...\\Module 3\\32-Body-Data"

# Where to save the populated output DOCX
output_docx_path:      "D:\\...\\Output_QIS.docx"

# Where to write log files
log_folder:            "D:\\...\\Cardiolek"

# Pages to SKIP at the start of specific PDFs (cover pages, TOC, etc.)
section_start_pages:
  "3.2.P.2":   2    # Skip cover + TOC, start at page 3
  "3.2.P.3.5": 2
  "3.2.P.7":   1

# Max content pages to extract per section (after skipping start pages)
section_page_limits:
  "3.2.P.2":   5
  "3.2.P.3.5": 3
  "3.2.P.7":   2
  "3.2.S.6":   3
  "3.2.S.7":   3
```

### Key Rules
- `section_start_pages` and `section_page_limits` are **optional**. Omit a section to extract all pages.
- The source folder is scanned **recursively** — subfolders are fine.
- Section numbers are detected from **filenames** (e.g. `3.2.P.3.1-Manufacturer.pdf` → `3.2.P.3.1`).

---

## Running

```bash
cd D:\QIS_Automation
python main.py
```

---

## Output & Logs

| Item | Location |
|------|----------|
| Populated DOCX | `output_docx_path` in config |
| Log file | `log_folder/qis_generation_YYYYMMDD_HHMMSS.log` |

### Summary printed at end:
```
==================================================
      FINAL GENERATION SUMMARY
==================================================
  Sections successfully filled : 22
  Warnings generated           : 3
  Total failures               : 3
  Output DOCX                  : D:\...\Output_QIS.docx
==================================================
```

---

## Template Placeholder Format

The QIS template must contain placeholders in this format:

```
Refer Section 3.2.P.3.1
Refer the section 3.2.S.4.1
```

- Matching is **case-insensitive**
- "the" is optional
- Sections listed in `MANUAL_ENTRY_SECTIONS` (`1.2`, `1.3`, `1.4`, `1.5`, `1.5.1`, `1.5.2`, `1.6`) are skipped — they require manual data entry

---

## Auto Header/Footer Removal

No hardcoded company names. The pipeline **automatically detects** repeating headers and footers by:

1. Scanning the top 12% and bottom 10% of every page in the source PDF
2. Any text appearing on **3 or more pages** in those margins is added to a noise blocklist
3. Matching paragraphs and table rows are removed before injection

This works for **any company's PDFs** with no configuration needed.

---

## Reusing for a Different Medicine / Client

1. Update `source_pdf_folder` to point to the new CTD Module 3 data folder
2. Update `template_docx_path` to the new QIS template
3. Update `output_docx_path` and `log_folder`
4. Adjust `section_start_pages` and `section_page_limits` as needed for the new PDFs
5. Run `python main.py`

No code changes required.

---

## Troubleshooting

| Problem | Cause | Fix |
|---------|-------|-----|
| `[WARNING]` red text in output DOCX | No PDF found for that section number | Check filename contains the correct section number (e.g. `3.2.P.4.1`) |
| Cover/TOC pages appearing in output | `section_start_pages` not set for that section | Add the section to `section_start_pages` in config |
| OCR messages in console | Source PDF is scanned (no text layer) — pdf2docx uses Tesseract automatically | Install Tesseract if not already installed |
| `Template DOCX not found` on startup | Wrong path in `template_docx_path` | Fix the path in `config.yaml` |
| Duplicate section warning in log | Two PDFs in the source folder match the same section number | Rename or remove the duplicate file |
