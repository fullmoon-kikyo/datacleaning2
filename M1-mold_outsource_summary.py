# -*- coding: utf-8 -*-
from __future__ import annotations

import importlib.util
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

try:
    from tqdm import tqdm
except ImportError:

    def tqdm(iterable=None, *args, **kwargs):
        return iterable


INPUT_DIR = Path("外委清单")
OUTPUT_PREFIX = "2-模具外委数据汇总"
SUMMARY_SHEET = "1-汇总"
DETAIL_SHEET = "2-汇总（去重）"
OUTSOURCE_DETAIL_SHEET = "3-外委明细"
NITRIDING_SHEET = "4-氮化明细"
REQUIRED_COLUMNS = ["模具号", "状态", "工序", "返厂时间"]
DETAIL_COLUMNS = ["模具号", "模具号-修正", "状态", "工序", "返厂时间"]
EXCEL_SUFFIXES = {".xls", ".xlsx", ".xlsm"}
INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
CHINESE_CHARS = re.compile(r"[\u4e00-\u9fff]")


def configure_console_encoding() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def log(message: str) -> None:
    print(message, flush=True)


def progress(iterable, **kwargs):
    return tqdm(iterable, file=sys.stdout, **kwargs)


def build_output_path() -> Path:
    now = datetime.now()
    timestamp = now.strftime("%Y%m%d-%H%M%S") + f"{now.microsecond // 10000:02d}"
    return Path(f"{OUTPUT_PREFIX}-{timestamp}.xlsx")


def list_excel_files(input_dir: Path) -> tuple[list[Path], list[Path]]:
    excel_files: list[Path] = []
    skipped_files: list[Path] = []

    for path in sorted(input_dir.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file():
            continue
        if path.name.startswith("~$"):
            skipped_files.append(path)
            continue
        if path.suffix.lower() in EXCEL_SUFFIXES:
            excel_files.append(path)
        else:
            skipped_files.append(path)

    return excel_files, skipped_files


def check_dependencies(excel_files: list[Path]) -> None:
    if any(path.suffix.lower() == ".xls" for path in excel_files):
        if importlib.util.find_spec("xlrd") is None:
            raise RuntimeError(
                "检测到 .xls 文件，但当前环境缺少 xlrd。\n"
                "请先执行：pip install xlrd\n"
                "安装后再运行：python M1-mold_outsource_summary.py"
            )


def get_writer_engine() -> str:
    if importlib.util.find_spec("xlsxwriter") is not None:
        return "xlsxwriter"
    return "openpyxl"


def clean_sheet_name(raw_name: str, used_names: set[str]) -> str:
    name = INVALID_SHEET_CHARS.sub("_", raw_name).strip()
    name = name.strip("'")
    if not name:
        name = "Sheet"

    base = name[:31]
    candidate = base
    index = 2
    while candidate in used_names:
        suffix = f"_{index}"
        candidate = f"{base[:31 - len(suffix)]}{suffix}"
        index += 1

    used_names.add(candidate)
    return candidate


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    return df


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def build_corrected_mold_no(value: object) -> str:
    mold_no = clean_text(value)
    if not mold_no:
        return ""
    if mold_no.upper().startswith("D"):
        return mold_no
    return f"D{mold_no}"


def contains_chinese(value: object) -> bool:
    return bool(CHINESE_CHARS.search(clean_text(value)))


def read_first_sheet(path: Path) -> pd.DataFrame:
    return clean_columns(pd.read_excel(path, sheet_name=0, dtype=object))


def ensure_required_columns(df: pd.DataFrame) -> None:
    missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_columns:
        missing_text = "、".join(missing_columns)
        existing_text = "、".join(map(str, df.columns))
        raise KeyError(
            f"{SUMMARY_SHEET} 缺少必需列：{missing_text}\n"
            f"当前可用列：{existing_text}"
        )


def build_detail(summary_df: pd.DataFrame) -> pd.DataFrame:
    detail = summary_df.loc[:, REQUIRED_COLUMNS].copy()
    detail["_读取顺序"] = range(len(detail))
    detail["_模具号_清洗"] = detail["模具号"].map(clean_text)
    detail["_工序_清洗"] = detail["工序"].map(clean_text)
    detail["_氮化优先级"] = detail["_工序_清洗"].ne("氮化").astype(int)

    detail = detail.sort_values(
        by=["_模具号_清洗", "_氮化优先级", "_读取顺序"],
        kind="mergesort",
    )
    detail = detail.drop_duplicates(subset=["_模具号_清洗"], keep="first")
    detail = detail.sort_values(by="_读取顺序", kind="mergesort")
    detail["模具号-修正"] = detail["模具号"].map(build_corrected_mold_no)
    detail = detail.loc[:, DETAIL_COLUMNS].reset_index(drop=True)
    return detail


def build_outsource_detail(detail_df: pd.DataFrame) -> pd.DataFrame:
    mask = ~detail_df["模具号"].map(contains_chinese)
    return detail_df.loc[mask, DETAIL_COLUMNS].reset_index(drop=True)


def build_nitriding_detail(outsource_detail_df: pd.DataFrame) -> pd.DataFrame:
    mask = outsource_detail_df["工序"].map(clean_text).eq("氮化")
    return outsource_detail_df.loc[mask, DETAIL_COLUMNS].reset_index(drop=True)


def write_workbook(
    output_path: Path,
    source_tables: list[tuple[str, pd.DataFrame]],
    summary_df: pd.DataFrame,
    detail_df: pd.DataFrame,
    outsource_detail_df: pd.DataFrame,
    nitriding_df: pd.DataFrame,
) -> None:
    used_sheet_names = {
        SUMMARY_SHEET,
        DETAIL_SHEET,
        OUTSOURCE_DETAIL_SHEET,
        NITRIDING_SHEET,
    }
    engine = get_writer_engine()
    log(f"写入引擎：{engine}")

    tables_to_write: list[tuple[str, pd.DataFrame]] = [
        (SUMMARY_SHEET, summary_df),
        (DETAIL_SHEET, detail_df),
        (OUTSOURCE_DETAIL_SHEET, outsource_detail_df),
        (NITRIDING_SHEET, nitriding_df),
    ]

    for source_name, df in source_tables:
        sheet_name = clean_sheet_name(source_name, used_sheet_names)
        tables_to_write.append((sheet_name, df))

    with pd.ExcelWriter(output_path, engine=engine) as writer:
        for index, (sheet_name, df) in enumerate(tables_to_write, start=1):
            log(
                f"写入 sheet {index}/{len(tables_to_write)}：{sheet_name}，"
                f"{len(df)} 行，{len(df.columns)} 列。"
            )
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            log(f"完成 sheet：{sheet_name}")


def main() -> None:
    configure_console_encoding()
    log("开始处理模具外委数据。")
    log(f"输入目录：{INPUT_DIR.resolve()}")

    if not INPUT_DIR.exists() or not INPUT_DIR.is_dir():
        raise FileNotFoundError(f"输入目录不存在：{INPUT_DIR.resolve()}")

    excel_files, skipped_files = list_excel_files(INPUT_DIR)
    for path in skipped_files:
        log(f"跳过非目标文件：{path.name}")

    if not excel_files:
        raise FileNotFoundError(f"{INPUT_DIR.resolve()} 中未找到 Excel 文件。")

    log(f"发现 Excel 文件 {len(excel_files)} 个。")
    check_dependencies(excel_files)

    source_tables: list[tuple[str, pd.DataFrame]] = []
    summary_parts: list[pd.DataFrame] = []

    for path in progress(excel_files, desc="读取文件", unit="个", dynamic_ncols=True):
        log(f"正在读取：{path.name}")
        df = read_first_sheet(path)
        source_tables.append((path.stem, df))

        summary_part = df.copy()
        summary_part.insert(0, "来源文件", path.name)
        summary_parts.append(summary_part)
        log(f"读取完成：{path.name}，{len(df)} 行，{len(df.columns)} 列。")

    summary_df = pd.concat(summary_parts, ignore_index=True)
    log(f"{SUMMARY_SHEET} 汇总完成，共 {len(summary_df)} 行。")

    ensure_required_columns(summary_df)

    before_dedupe_rows = len(summary_df)
    detail_df = build_detail(summary_df)
    log(
        f"{DETAIL_SHEET} 去重完成：去重前 {before_dedupe_rows} 行，"
        f"去重后 {len(detail_df)} 行。"
    )

    outsource_detail_df = build_outsource_detail(detail_df)
    log(
        f"{OUTSOURCE_DETAIL_SHEET} 筛选完成：筛选前 {len(detail_df)} 行，"
        f"筛选后 {len(outsource_detail_df)} 行。"
    )

    nitriding_df = build_nitriding_detail(outsource_detail_df)
    log(f"{NITRIDING_SHEET} 提取完成，共 {len(nitriding_df)} 行。")

    output_path = build_output_path()
    log(f"正在写入输出文件：{output_path.resolve()}")
    write_workbook(
        output_path,
        source_tables,
        summary_df,
        detail_df,
        outsource_detail_df,
        nitriding_df,
    )
    log("处理完成。")
    log(f"输出文件：{output_path.resolve()}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        log(f"处理失败：{exc}")
        sys.exit(1)
