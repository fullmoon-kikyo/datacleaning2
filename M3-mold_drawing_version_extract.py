# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

try:
    from tqdm import tqdm
except ImportError:
    def tqdm(iterable=None, *args, **kwargs):
        return iterable


INPUT_FILE_NAME = "1510-过程组件分析结果-20260422.xlsx"
INPUT_FILE_GLOB = "1510-*20260422.xlsx"
INPUT_SHEET_NAME = "分析结果"
CHILD_COLUMN = "子件号"

PACKAGE_MOLD_COLUMN = "成套模具号"
INITIAL_MOLD_COLUMN = "初始模具号"
COLOR_GROUP_COLUMN = "分组色号"

BASE_NEW_COLUMNS = [PACKAGE_MOLD_COLUMN, INITIAL_MOLD_COLUMN]
VERSION_SUMMARY_COLUMNS = ["变更履历", "当前版本", "版本数量"]
DEFAULT_VERSION_COLUMNS = ["0版"] + [f"{letter}版" for letter in "ABCDEF"]

GROUP_FILL_COLORS = [
    "B7DEE8",
    "C6E0B4",
    "FFD966",
    "F4B084",
    "D9B3E6",
    "9FE2BF",
    "A9C4E8",
    "F8CBAD",
]

SUBGROUP_KEY_COLUMN = PACKAGE_MOLD_COLUMN
SUBGROUP_COLUMNS = [PACKAGE_MOLD_COLUMN, "是否无用物料", "是否冻结物料", "不维护过程组件"]
BORDER_COLOR = "305496"


def configure_console_encoding() -> None:
    """尽量保证 Windows 控制台中文提示正常显示。"""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def log(message: str) -> None:
    print(message, flush=True)


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def resolve_input_file(input_name: str) -> Path:
    exact_path = Path(input_name)
    if exact_path.exists():
        return exact_path

    if exact_path.suffix:
        pattern = f"{exact_path.stem}*{exact_path.suffix}"
    else:
        pattern = f"{input_name}*.xls*"
    matches = sorted(Path(".").glob(pattern))
    if matches:
        return matches[0]

    if input_name == INPUT_FILE_NAME:
        matches = sorted(Path(".").glob(INPUT_FILE_GLOB))
        if matches:
            return matches[0]

    raise FileNotFoundError(f"未找到输入文件: {input_name}")


def build_output_path(input_path: Path, output_name: str | None = None) -> Path:
    if output_name:
        return Path(output_name)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return input_path.with_name(f"{input_path.stem}【处理后】{timestamp}{input_path.suffix}")


def split_child_no(value: object) -> tuple[str, str, str]:
    package_mold_no = ""
    initial_mold_no = ""
    version = ""

    text = clean_text(value)
    if not text:
        return package_mold_no, initial_mold_no, version

    if "/" in text:
        package_mold_no = text.split("/", 1)[0].strip()
    else:
        package_mold_no = text

    if package_mold_no and package_mold_no[-1].isalpha():
        initial_mold_no = package_mold_no[:-1]
        version = package_mold_no[-1].upper()
    else:
        initial_mold_no = package_mold_no
        version = "0"

    return package_mold_no, initial_mold_no, version


def assign_group_color_ids(df: pd.DataFrame) -> None:
    color_by_initial: dict[str, int] = {}
    palette_index = 0
    previous_color_id = 0
    initial_values = [clean_text(value) for value in df[INITIAL_MOLD_COLUMN].tolist()]
    df[COLOR_GROUP_COLUMN] = ""

    run_start = 0
    while run_start < len(initial_values):
        initial_mold_no = initial_values[run_start]
        run_end = run_start
        while run_end + 1 < len(initial_values) and initial_values[run_end + 1] == initial_mold_no:
            run_end += 1

        version_count_text = clean_text(df.at[df.index[run_start], "版本数量"]) if initial_mold_no else ""
        try:
            version_count = int(float(version_count_text))
        except ValueError:
            version_count = 0

        if initial_mold_no and version_count >= 2:
            color_id = color_by_initial.get(initial_mold_no)
            if color_id is None:
                color_id = palette_index % len(GROUP_FILL_COLORS) + 1
                palette_index += 1
                if color_id == previous_color_id:
                    color_id = palette_index % len(GROUP_FILL_COLORS) + 1
                    palette_index += 1
                color_by_initial[initial_mold_no] = color_id
            df.iloc[run_start : run_end + 1, df.columns.get_loc(COLOR_GROUP_COLUMN)] = color_id
            previous_color_id = color_id

        run_start = run_end + 1


def apply_group_row_colors(workbook, ws, df: pd.DataFrame) -> None:
    color_group_col = df.columns.get_loc(COLOR_GROUP_COLUMN) + 1
    color_group_col_zero_based = color_group_col - 1
    formats = {
        color_idx: workbook.add_format({"bg_color": f"#{color}"})
        for color_idx, color in enumerate(GROUP_FILL_COLORS, start=1)
    }

    color_ids = pd.to_numeric(df[COLOR_GROUP_COLUMN], errors="coerce").fillna(0).astype(int)
    visible_col_count = df.columns.get_loc(COLOR_GROUP_COLUMN)
    for row_idx, color_id in enumerate(color_ids, start=1):
        if color_id > 0:
            row_values = df.iloc[row_idx - 1, :visible_col_count].tolist()
            for col_idx, value in enumerate(row_values):
                if pd.isna(value):
                    ws.write_blank(row_idx, col_idx, None, formats[color_id])
                else:
                    ws.write(row_idx, col_idx, value, formats[color_id])

    ws.set_column(color_group_col_zero_based, color_group_col_zero_based, None, None, {"hidden": True})
    colored_rows = int(color_ids.gt(0).sum())
    colored_groups = df.loc[color_ids.gt(0), INITIAL_MOLD_COLUMN].nunique()
    log(f"着色完成: 共覆盖 {colored_groups} 个多版本分组，{colored_rows} 行。")


def apply_child_subgroup_borders(workbook, ws, df: pd.DataFrame) -> None:
    missing_cols = [col for col in SUBGROUP_COLUMNS if col not in df.columns]
    if missing_cols:
        log(f"提示: 缺少子组边框列 {missing_cols}，将自动跳过这些列。")

    target_columns = [col for col in SUBGROUP_COLUMNS if col in df.columns]
    if not target_columns:
        log("提示: 没有可用于子组边框的列，跳过子组边框步骤。")
        return

    target_col_indexes = [df.columns.get_loc(col) for col in target_columns]
    group_formats: dict[tuple[int, bool, bool, bool, bool], object] = {}
    colored_mask = pd.to_numeric(df[COLOR_GROUP_COLUMN], errors="coerce").fillna(0).astype(int).gt(0)
    initial_values = [clean_text(value) for value in df[INITIAL_MOLD_COLUMN].tolist()]
    child_values = [clean_text(value) for value in df[SUBGROUP_KEY_COLUMN].tolist()]

    colored_subgroups = 0
    colored_rows = 0

    def get_format(color_id: int, top: bool, bottom: bool, left: bool, right: bool):
        key = (color_id, top, bottom, left, right)
        if key not in group_formats:
            fmt_props = {"bg_color": f"#{GROUP_FILL_COLORS[color_id - 1]}"}
            if top:
                fmt_props["top"] = 2
                fmt_props["top_color"] = f"#{BORDER_COLOR}"
            if bottom:
                fmt_props["bottom"] = 2
                fmt_props["bottom_color"] = f"#{BORDER_COLOR}"
            if left:
                fmt_props["left"] = 2
                fmt_props["left_color"] = f"#{BORDER_COLOR}"
            if right:
                fmt_props["right"] = 2
                fmt_props["right_color"] = f"#{BORDER_COLOR}"
            group_formats[key] = workbook.add_format(fmt_props)
        return group_formats[key]

    main_start = 0
    while main_start < len(df):
        initial_mold_no = initial_values[main_start]
        main_end = main_start
        while main_end + 1 < len(df) and initial_values[main_end + 1] == initial_mold_no:
            main_end += 1

        if not initial_mold_no or not bool(colored_mask.iloc[main_start]):
            main_start = main_end + 1
            continue

        child_start = main_start
        while child_start <= main_end:
            child_no = child_values[child_start]
            child_end = child_start
            while child_end + 1 <= main_end and child_values[child_end + 1] == child_no:
                child_end += 1

            color_id = int(pd.to_numeric(df.at[df.index[child_start], COLOR_GROUP_COLUMN], errors="coerce"))
            for row_idx in range(child_start, child_end + 1):
                for position, col_idx in enumerate(target_col_indexes):
                    value = df.iat[row_idx, col_idx]
                    fmt = get_format(
                        color_id=color_id,
                        top=row_idx == child_start,
                        bottom=row_idx == child_end,
                        left=position == 0,
                        right=position == len(target_col_indexes) - 1,
                    )
                    if pd.isna(value):
                        ws.write_blank(row_idx + 1, col_idx, None, fmt)
                    else:
                        ws.write(row_idx + 1, col_idx, value, fmt)

            colored_subgroups += 1
            colored_rows += child_end - child_start + 1
            child_start = child_end + 1

        main_start = main_end + 1

    log(f"子组边框完成: 共覆盖 {colored_subgroups} 个子组，{colored_rows} 行。")


def build_version_columns(found_versions: set[str]) -> list[str]:
    extra_versions = sorted(found_versions - set("0ABCDEF"))
    return DEFAULT_VERSION_COLUMNS + [f"{letter}版" for letter in extra_versions]


def add_version_columns(df: pd.DataFrame, version_columns: list[str]) -> None:
    for col in BASE_NEW_COLUMNS + VERSION_SUMMARY_COLUMNS + version_columns:
        df[col] = ""


def fill_parsed_version_columns(
    df: pd.DataFrame,
    parsed_rows: list[tuple[str, str, str]],
) -> None:
    for row_idx, (package_mold_no, initial_mold_no, version) in tqdm(
        zip(df.index, parsed_rows),
        total=len(parsed_rows),
        desc="回填版本列",
        unit="行",
        file=sys.stdout,
    ):
        df.at[row_idx, PACKAGE_MOLD_COLUMN] = package_mold_no
        df.at[row_idx, INITIAL_MOLD_COLUMN] = initial_mold_no
        if version:
            df.at[row_idx, f"{version}版"] = version


def fill_version_summary(df: pd.DataFrame, version_columns: list[str]) -> None:
    versions_by_initial: dict[str, set[str]] = {}
    for row_idx in tqdm(df.index, total=len(df), desc="建立版本分组", unit="行", file=sys.stdout):
        initial_mold_no = clean_text(df.at[row_idx, INITIAL_MOLD_COLUMN])
        if not initial_mold_no:
            continue

        versions_by_initial.setdefault(initial_mold_no, set())
        for col in version_columns:
            version = clean_text(df.at[row_idx, col])
            if version:
                versions_by_initial[initial_mold_no].add(version)

    version_order = [col.removesuffix("版") for col in version_columns]
    version_summary_by_initial: dict[str, tuple[str, str, int]] = {}
    for initial_mold_no, versions in versions_by_initial.items():
        ordered_versions = [version for version in version_order if version in versions]
        if ordered_versions:
            history = "" if ordered_versions == ["0"] else f"【{''.join(ordered_versions)}】"
            current_version = f"-{ordered_versions[-1]}-"
            version_summary_by_initial[initial_mold_no] = (history, current_version, len(ordered_versions))

    for row_idx in tqdm(df.index, total=len(df), desc="回填分组版本", unit="行", file=sys.stdout):
        initial_mold_no = clean_text(df.at[row_idx, INITIAL_MOLD_COLUMN])
        summary = version_summary_by_initial.get(initial_mold_no)
        if summary:
            df.at[row_idx, "变更履历"] = summary[0]
            df.at[row_idx, "当前版本"] = summary[1]
            df.at[row_idx, "版本数量"] = summary[2]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="提取模具版本并按初始模具号汇总履历")
    parser.add_argument("--input", default=INPUT_FILE_NAME, help=f"输入 Excel 文件，默认: {INPUT_FILE_NAME}")
    parser.add_argument("--sheet", default=INPUT_SHEET_NAME, help=f"目标工作表，默认: {INPUT_SHEET_NAME}")
    parser.add_argument("--output", default=None, help="输出 Excel 文件，默认在输入文件名后追加处理时间")
    parser.add_argument("--subborder", action="store_true", help="对已染色分组内的成套模具号子组添加边框标识")
    return parser.parse_args()


def main() -> None:
    configure_console_encoding()
    args = parse_args()

    log("开始处理分析结果模具版本列。")
    input_path = resolve_input_file(args.input)
    sheet_name = args.sheet
    log(f"读取文件: {input_path.resolve()}")
    log(f"读取工作表: {sheet_name}")

    try:
        df1 = pd.read_excel(input_path, sheet_name=sheet_name)
    except ValueError as exc:
        raise ValueError(f"未找到工作表: {sheet_name}") from exc

    if CHILD_COLUMN not in df1.columns:
        raise KeyError(f"分析结果缺少必要列: {CHILD_COLUMN}")

    log(f"读取完成: df1 共 {len(df1)} 行，{len(df1.columns)} 列。")

    log("步骤1/5: 解析子件号，生成成套模具号、初始模具号和版本字母。")
    parsed_rows: list[tuple[str, str, str]] = []
    found_versions: set[str] = set()

    for row_idx in tqdm(df1.index, total=len(df1), desc="解析子件号", unit="行", file=sys.stdout):
        package_mold_no, initial_mold_no, version = split_child_no(df1.at[row_idx, CHILD_COLUMN])
        parsed_rows.append((package_mold_no, initial_mold_no, version))
        if version:
            found_versions.add(version)

    version_columns = build_version_columns(found_versions)

    log("步骤2/5: 追加空白列。")
    add_version_columns(df1, version_columns)
    log("已新增变更履历、当前版本、版本数量和版本列。")

    if found_versions:
        detected = "、".join(sorted(found_versions))
        version_cols_text = "、".join(version_columns)
        log(f"检测到版本字母: {detected}")
        log(f"将写入版本列: {version_cols_text}")
    else:
        log("未检测到额外版本，仅保留默认 0版、A版-F版 空白列。")

    log("步骤3/5: 回填成套模具号、初始模具号和版本列。")
    fill_parsed_version_columns(df1, parsed_rows)

    log("步骤4/5: 按初始模具号合并变更履历并计算当前版本。")
    fill_version_summary(df1, version_columns)

    output_path = build_output_path(input_path, args.output)
    assign_group_color_ids(df1)
    log("步骤5/6: 写出处理后的分析结果。")
    log(f"输出文件: {output_path.resolve()}")
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df1.to_excel(writer, sheet_name=sheet_name, index=False)
        log("步骤6/6: 仅对版本数量>=2的分组着色。")
        apply_group_row_colors(writer.book, writer.sheets[sheet_name], df1)
        if args.subborder:
            log("额外步骤: 对已染色分组内的成套模具号子组进行边框标识。")
            apply_child_subgroup_borders(writer.book, writer.sheets[sheet_name], df1)

    log("处理完成。")


if __name__ == "__main__":
    main()
