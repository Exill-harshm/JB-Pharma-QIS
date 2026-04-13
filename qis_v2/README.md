# QIS Automation

This module fills QIS template tables from dossier PDFs as editable DOCX content.

## What it fills

- `1.4.2` QIS summary table from Module 1 (table-driven extraction, not image paste)
- Related dossiers table
- API section (`2.3.S`) rows including API name/manufacturer and option checkbox
- `2.3.S.2.1 Manufacturer(s)` section table
- `2.3.P.3.1 Manufacturer(s)` section table

## Source mapping

- Summary source: Module 1 PDF under dossier root (prefers files containing QIS heading)
- API/manufacture source: primary `3.2.S.2.1`, fallback to parent CTD sections
- P3.1 source: `3.2.P.3.1`, fallback via section mapper

## Run

```powershell
cd "D:\QOS Automation\QIS"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python run.py \
  --template "D:\path\to\QUALITY INFORMATION SUMMARY.docx" \
  --dossier-root "D:\path\to\Medicine" \
  --output "D:\path\to\Quality Information Summary_QIS_Auto.docx"
```
