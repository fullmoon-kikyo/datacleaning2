# datacleaning2

Python scripts for cleaning and transforming S4 mold/material Excel data.

## Scripts

- `process_mold_data.py`: legacy process component analysis script. It reads 591E material data, material list data, and nitriding detail data, then produces an analysis workbook.
- `clean_mold_data_20260414.py`: newer mold data cleaning script. It adds material freeze status, BOM status checks, process suffix handling, nitriding matching, and formatted Excel output.
- `clean_bom_transform.py`: consolidates EA/ZM BOM conversion relationships and marks missing or abnormal conversion records.
- `extract_mold_versions.py`: extracts and summarizes mold version history from processed analysis results.
- `extract_mold_versions_subborder.py`: variant of mold version extraction for sub-border style mold numbers.

## Environment

Install dependencies in your Python or Conda environment:

```powershell
pip install pandas openpyxl tqdm
```

## Usage

Run a script directly from this directory:

```powershell
python clean_mold_data_20260414.py
python clean_bom_transform.py
python process_mold_data.py
```

Input and output Excel files are intentionally not committed to Git. Put the required source files in the project directory before running the scripts.

## Git Notes

The repository ignores Excel/CSV data files, generated reports, Python cache files, virtual environments, and local editor settings.
