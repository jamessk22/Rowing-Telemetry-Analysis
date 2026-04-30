# XLSX to CSV Conversion

Detailed reference for [scripts/convert_xlsx_to_csv.py](scripts/convert_xlsx_to_csv.py).

## Arguments

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

## Output Behavior

- Default behavior writes each CSV in the same seat folder as its source workbook.
- If you pass `--output-dir`, the script preserves the relative seat-folder structure under that destination.
- By default, existing CSV files are skipped. Use `--overwrite` to regenerate them.
- By default, only the first worksheet is exported, which matches the telemetry files currently in `427MC-C/data`.
- With `--all-sheets`, extra worksheet exports get the sheet name appended to the CSV filename.
- With `--delete-xlsx`, the source workbook is removed only after successful CSV export.

## Suggested Workflow

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
