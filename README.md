# datacleaning2

Python scripts for cleaning and transforming S4 mold/material Excel data.

## Scripts

- `M1-mold_outsource_summary.py`: consolidates all Excel files in `外委清单`, creates `1-汇总`, `2-汇总（去重）`, `3-外委明细`, and `4-氮化明细`, and writes `M1-模具外委数据汇总-YYYYMMDD-HHMMSSxx.xlsx`.
- `M2-mold_process_component_analysis.py`: standard mold process component analysis. It reads outsource records from the latest `M1-模具外委数据汇总-*.xlsx` or compatible legacy files, keeps the simpler M2 result structure, preserves `基本视图状态` and `工厂视图状态` as two-character text values such as `09`, and writes output files named `M2-过程组件分析结果-YYYYMMDD-HHMMSSxx.xlsx`.
- `M2.1-mold_process_component_analysis_enhancement.py`: enhanced mold process component analysis. It reads outsource records from the latest `M1-模具外委数据汇总-*.xlsx` or compatible legacy files, adds material freeze status, missing BOM status, `MT` nitriding suffix support, dated output naming, and formatted Excel output files named `M2.1-处理后模具数据_YYYYMMDD-HHMMSSxx.xlsx`.
- `M3-mold_drawing_version_extract.py`: extracts and summarizes mold drawing version history from processed analysis results. Use `--subborder` to add subgroup border formatting.
- `M4-bom_transition_relationship.py`: consolidates EA/ZM BOM transition relationships and marks missing or abnormal conversion records.

## Environment

Install dependencies in your Python or Conda environment:

```powershell
pip install pandas openpyxl tqdm xlrd XlsxWriter
```

## Usage

Run a script directly from this directory:

```powershell
python M1-mold_outsource_summary.py
python M2-mold_process_component_analysis.py
python M2.1-mold_process_component_analysis_enhancement.py
python M3-mold_drawing_version_extract.py
python M3-mold_drawing_version_extract.py --subborder
python M4-bom_transition_relationship.py
```

Input and output Excel files are intentionally not committed to Git. Put the required source files in the project directory before running the scripts.

## Git Notes

The repository ignores Excel/CSV data files, generated reports, Python cache files, virtual environments, and local editor settings.
