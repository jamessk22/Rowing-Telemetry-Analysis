# Rowing-Telemetry-Analysis
Storing and analysis of crew telemetry data

## XLSX to CSV conversion

Bulk converter for telemetry workbooks:

[scripts/convert_xlsx_to_csv.py](scripts/convert_xlsx_to_csv.py)

It scans a folder recursively for `.xlsx` files and converts them to `.csv`. The default workflow is designed for the rowing data layout already in this repo: each seat has its own folder, and the generated CSV is written into that same seat folder beside the original workbook.

Full usage and option details live at [convert_xlsx_to_csv.md](convert_xlsx_to_csv.md).

### What it does

- Recursively searches the folder you give it for `.xlsx` files.
- Converts the first worksheet from each workbook by default.
- Writes each CSV next to its source workbook unless you choose a different output root.
- Preserves the relative folder structure when `--output-dir` is used.
- Skips existing CSVs unless you explicitly ask to overwrite them.

### Usage

Convert every workbook under a day folder and write each CSV next to its source workbook, preserving seat folders:

```bash
python3 scripts/convert_xlsx_to_csv.py 427MC-C/data
```

Convert the same folder but place the CSV outputs under a separate destination while keeping the seat folder structure:

```bash
python3 scripts/convert_xlsx_to_csv.py 427MC-C/data --output-dir 427MC-C/converted_data
```

Overwrite previously generated CSVs in place:

```bash
python3 scripts/convert_xlsx_to_csv.py 427MC-C/data --overwrite
```

Export every worksheet from each workbook instead of only the first sheet:

```bash
python3 scripts/convert_xlsx_to_csv.py 427MC-C/data --all-sheets
```

Delete each `.xlsx` file after a successful conversion:

```bash
python3 scripts/convert_xlsx_to_csv.py 427MC-C/data --overwrite --delete-xlsx
```
