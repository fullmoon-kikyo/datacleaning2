# datacleaning2

Python scripts for cleaning and transforming S4 mold/material Excel data.

## Scripts

- `M1-mold_process_component_analysis.py`: standard mold process component analysis. It keeps the simpler M1 result structure, preserves `基本视图状态` and `工厂视图状态` as two-character text values such as `09`, and writes output files named `M1-分析结果-D物料过程组件分析-YYYYMMDD-HHMM.xlsx`.
- `M1.1-mold_process_component_analysis_enhancement.py`: enhanced mold process component analysis. It adds material freeze status, missing BOM status, `MT` nitriding suffix support, dated output naming, and formatted Excel output.
- `M2-mold_drawing_version_extract.py`: extracts and summarizes mold drawing version history from processed analysis results. Use `--subborder` to add subgroup border formatting.
- `M3-bom_transition_relationship.py`: consolidates EA/ZM BOM transition relationships and marks missing or abnormal conversion records.

## Environment

Install dependencies in your Python or Conda environment:

```powershell
pip install pandas openpyxl tqdm
```

## Usage

Run a script directly from this directory:

```powershell
python M1-mold_process_component_analysis.py
python M1.1-mold_process_component_analysis_enhancement.py
python M2-mold_drawing_version_extract.py
python M2-mold_drawing_version_extract.py --subborder
python M3-bom_transition_relationship.py
```

Input and output Excel files are intentionally not committed to Git. Put the required source files in the project directory before running the scripts.

## Git Notes

The repository ignores Excel/CSV data files, generated reports, Python cache files, virtual environments, and local editor settings.
