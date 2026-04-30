# Rowing-Telemetry-Analysis
Storing and analysis of crew telemetry data

## XLSX to CSV conversion

This repo now includes a bulk converter for telemetry workbooks:

[scripts/convert_xlsx_to_csv.py](/Users/jamesskowronek/Coding/Rowing-Telemetry-Analysis/scripts/convert_xlsx_to_csv.py)

It scans a folder recursively for `.xlsx` files and converts them to `.csv`. The default workflow is designed for the rowing data layout already in this repo: each seat has its own folder, and the generated CSV is written into that same seat folder beside the original workbook.

### What it does

- Recursively searches the folder you give it for `.xlsx` files.
- Converts the first worksheet from each workbook by default.
- Writes each CSV next to its source workbook unless you choose a different output root.
- Preserves the relative folder structure when `--output-dir` is used.
- Skips existing CSVs unless you explicitly ask to overwrite them.

### Common usage

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

### Arguments

`input_dir`

- Required positional argument.
- This is the top-level folder the script will scan recursively.
- Example: `427MC-C/data`

`-o, --output-dir`

- Optional destination root for CSV output.
- If omitted, CSVs are written next to the source `.xlsx` files.
- If provided, the script keeps the same relative folder structure under the new destination.
- Example: `--output-dir 427MC-C/converted_data`

`--all-sheets`

- Optional flag.
- By default, the script exports only the first worksheet in each workbook, which matches the telemetry files currently in this repo.
- If a workbook has multiple worksheets and you use `--all-sheets`, the script exports all of them.
- When needed, extra worksheets are named by appending the worksheet name to the CSV filename so files do not collide.

`--overwrite`

- Optional flag.
- By default, if a target CSV already exists, the script skips it and leaves it untouched.
- With `--overwrite`, existing CSVs are regenerated.
- This is useful after new exports arrive or after changing the conversion script.

`--delete-xlsx`

- Optional flag.
- After a workbook has been converted successfully, the original `.xlsx` file is deleted.
- Deletion only happens after at least one CSV has been written for that workbook.
- This is the most destructive option in the script, so it is best used together with `--overwrite` once you have confirmed the CSV output looks correct.

### Output behavior

- Default behavior writes each CSV in the same seat folder as its source workbook.
- If you pass `--output-dir`, the script preserves the relative seat-folder structure under that destination.
- By default, existing CSV files are skipped. Use `--overwrite` to regenerate them.
- By default, only the first worksheet is exported, which matches the telemetry files currently in `427MC-C/data`.
- With `--all-sheets`, extra worksheet exports get the sheet name appended to the CSV filename.
- With `--delete-xlsx`, the source workbook is removed only after successful CSV export.

### Suggested workflow

1. Drop new `.xlsx` files into the seat folders for the day.
2. Run the converter without deletion first:

```bash
python3 scripts/convert_xlsx_to_csv.py 427MC-C/data --overwrite
```

3. Verify the CSVs look right in the same seat folders.
4. If you want to keep only CSVs going forward, rerun with deletion:

```bash
python3 scripts/convert_xlsx_to_csv.py 427MC-C/data --overwrite --delete-xlsx
```
