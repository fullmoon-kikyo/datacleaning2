# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

try:
    from tqdm import tqdm
except ImportError:  # 兼容未安装 tqdm 的环境。
    def tqdm(iterable=None, *args, **kwargs):
        return iterable


INPUT_591E = Path("591E20260409.xlsx")
INPUT_MATERIAL_LIST = Path("2-物料列表.xlsx")
INPUT_NITRIDING_DETAIL = Path("3-氮化明细.xlsx")
OUTPUT_FILE = Path("1510-过程组件分析结果.xlsx")

PROCESS_SUFFIXES = {"MC", "MR", "MN"}
PROCESS_COLUMNS = {
    "MC": "粗加工",
    "MR": "热处理",
    "MN": "氮化",
}


def configure_console_encoding() -> None:
    """尽量保证 Windows 控制台中文提示正常显示。"""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def log(message: str) -> None:
    print(message, flush=True)


def progress(
    iterable,
    desc: str,
    total: int | None = None,
    colour: str = "green",
):
    return tqdm(
        iterable,
        desc=desc,
        total=total,
        unit="行",
        dynamic_ncols=True,
        file=sys.stdout,
        colour=colour,
    )


def to_clean_text(value: object) -> str:
    """将 Excel 单元格内容转成去掉首尾空格的字符串，空值保留为空字符串。"""
    if pd.isna(value):
        return ""
    return str(value).strip()


def find_column(df: pd.DataFrame, candidates: list[str], table_name: str) -> str:
    for col in candidates:
        if col in df.columns:
            return col
    raise KeyError(f"{table_name} 缺少必要列，候选列名: {candidates}")


def split_material_code(value: object) -> tuple[str, str]:
    text = to_clean_text(value)
    tail = text[-2:].upper() if len(text) >= 2 else ""
    if tail in PROCESS_SUFFIXES:
        return text[:-2], tail
    return text, ""


def ensure_blank_column(df: pd.DataFrame, col_name: str) -> None:
    df[col_name] = ""


def append_new_materials(df2: pd.DataFrame, df1a: pd.DataFrame) -> pd.DataFrame:
    existing: set[str] = set()
    for value in progress(df2["子件号"], desc="建立 df2 子件号索引", colour="blue"):
        existing.add(to_clean_text(value))

    rows_to_append: list[dict[str, str]] = []
    for value in progress(df1a["S4物料号"], desc="比对新增物料", colour="yellow"):
        s4_material = to_clean_text(value)
        if s4_material and s4_material not in existing:
            rows_to_append.append(
                {
                    "子件号": s4_material,
                    "成品物料号": "",
                    "是否为新增物料": "是",
                    "粗加工": "",
                    "热处理": "",
                    "氮化": "",
                }
            )
            existing.add(s4_material)

    if rows_to_append:
        log(f"发现新增物料 {len(rows_to_append)} 条，正在追加到 df2。")
        df2 = pd.concat([df2, pd.DataFrame(rows_to_append)], ignore_index=True)
    else:
        log("未发现需要追加的新增物料。")
    return df2


def main() -> None:
    configure_console_encoding()

    log("开始处理 Excel 数据。")
    log(f"输入文件 1: {INPUT_591E.resolve()}")
    log(f"输入文件 2: {INPUT_MATERIAL_LIST.resolve()}")
    log(f"输入文件 3: {INPUT_NITRIDING_DETAIL.resolve()}")

    log("步骤 1/8：读取 591E 数据。")
    df1 = pd.read_excel(INPUT_591E)
    material_col = find_column(df1, ["物料号", "物料"], "591E20260409.xlsx")
    log(f"读取完成：df1 共 {len(df1)} 行，使用物料列“{material_col}”。")

    log("步骤 2/8：拆分 df1 物料号，生成“子件号”和“工序”。")
    ensure_blank_column(df1, "子件号")
    ensure_blank_column(df1, "工序")

    AA = ""
    BB = ""
    for i in progress(df1.index, desc="拆分 df1 物料号", total=len(df1), colour="cyan"):
        AA, BB = split_material_code(df1.at[i, material_col])
        df1.at[i, "子件号"] = AA
        df1.at[i, "工序"] = BB

    log("步骤 3/8：提取 df1 子件号去重结果，生成 df2。")
    df2 = df1[["子件号"]].drop_duplicates().reset_index(drop=True)
    for col in ["成品物料号", "是否为新增物料", "粗加工", "热处理", "氮化"]:
        ensure_blank_column(df2, col)
    log(f"df2 初始去重后共 {len(df2)} 行。")

    log("步骤 4/8：读取 2-物料列表.xlsx，生成 S4物料号并比对新增物料。")
    df1A = pd.read_excel(INPUT_MATERIAL_LIST, dtype={"模具号": str})
    mold_col_1a = find_column(df1A, ["模具号"], "2-物料列表.xlsx")
    df1A["S4物料号"] = ""
    for i in progress(df1A.index, desc="生成 df1A S4物料号", total=len(df1A), colour="magenta"):
        mold_no = to_clean_text(df1A.at[i, mold_col_1a])
        df1A.at[i, "S4物料号"] = f"S-{mold_no}" if mold_no else ""

    df2 = append_new_materials(df2, df1A)
    log(f"新增物料比对完成：df2 当前共 {len(df2)} 行。")

    log("步骤 5/8：根据 df1 工序后缀，回填 df2 的成品物料号、粗加工、热处理、氮化列。")
    row_index_by_child: dict[str, int] = {}
    for i in progress(df2.index, desc="建立 df2 行号索引", total=len(df2), colour="blue"):
        child_no = to_clean_text(df2.at[i, "子件号"])
        if child_no:
            row_index_by_child[child_no] = i

    CC = ""
    DD = ""
    for i in progress(df1.index, desc="回填 df2 工序信息", total=len(df1), colour="green"):
        CC = ""
        DD = ""
        CC, DD = split_material_code(df1.at[i, material_col])
        if not CC or CC not in row_index_by_child:
            continue

        t = row_index_by_child[CC]
        if DD in PROCESS_COLUMNS:
            df2.at[t, PROCESS_COLUMNS[DD]] = DD
        else:
            df2.at[t, "成品物料号"] = CC

    log("步骤 6/8：读取 3-氮化明细.xlsx，生成 S4物料号。")
    ensure_blank_column(df2, "是否存在氮化记录")

    df3 = pd.read_excel(INPUT_NITRIDING_DETAIL, dtype={"模具号": str})
    mold_col_3 = find_column(df3, ["模具号"], "3-氮化明细.xlsx")
    df3["S4物料号"] = ""
    for i in progress(df3.index, desc="生成 df3 S4物料号", total=len(df3), colour="magenta"):
        mold_no = to_clean_text(df3.at[i, mold_col_3])
        df3.at[i, "S4物料号"] = f"S-{mold_no}" if mold_no else ""

    log("步骤 7/8：比对 df2 子件号和氮化明细，回填“是否存在氮化记录”。")
    nitriding_index_by_s4: dict[str, int] = {}
    for i in progress(df3.index, desc="建立氮化明细索引", total=len(df3), colour="blue"):
        s4_no = to_clean_text(df3.at[i, "S4物料号"])
        if s4_no:
            nitriding_index_by_s4[s4_no] = i

    FF = ""
    matched_nitriding_count = 0
    for i in progress(df2.index, desc="匹配氮化记录", total=len(df2), colour="yellow"):
        FF = to_clean_text(df2.at[i, "子件号"])
        if FF in nitriding_index_by_s4:
            h = nitriding_index_by_s4[FF]
            df2.at[i, "是否存在氮化记录"] = df3.at[h, "S4物料号"]
            matched_nitriding_count += 1
    log(f"氮化记录匹配完成：共匹配 {matched_nitriding_count} 条。")

    log("步骤 8/8：写入处理后模具数据.xlsx。")
    actual_output = OUTPUT_FILE
    try:
        with pd.ExcelWriter(actual_output, engine="openpyxl") as writer:
            log("正在写入 sheet：Data")
            df1.to_excel(writer, sheet_name="Data", index=False)
            log("正在写入 sheet：外委明细")
            df1A.to_excel(writer, sheet_name="外委明细", index=False)
            log("正在写入 sheet：氮化明细")
            df3.to_excel(writer, sheet_name="氮化明细", index=False)
            log("正在写入 sheet：分析结果")
            df2.to_excel(writer, sheet_name="分析结果", index=False)
    except PermissionError:
        actual_output = OUTPUT_FILE.with_name(
            f"{OUTPUT_FILE.stem}_{datetime.now():%Y%m%d_%H%M%S}{OUTPUT_FILE.suffix}"
        )
        log(f"无法覆盖 {OUTPUT_FILE}，文件可能正在打开。")
        log(f"正在改为另存到: {actual_output.resolve()}")
        with pd.ExcelWriter(actual_output, engine="openpyxl") as writer:
            log("正在写入 sheet：Data")
            df1.to_excel(writer, sheet_name="Data", index=False)
            log("正在写入 sheet：外委明细")
            df1A.to_excel(writer, sheet_name="外委明细", index=False)
            log("正在写入 sheet：氮化明细")
            df3.to_excel(writer, sheet_name="氮化明细", index=False)
            log("正在写入 sheet：分析结果")
            df2.to_excel(writer, sheet_name="分析结果", index=False)

    log("处理完成。")
    log(f"输出文件: {actual_output.resolve()}")
    log(f"Data 行数: {len(df1)}")
    log(f"外委明细 行数: {len(df1A)}")
    log(f"氮化明细 行数: {len(df3)}")
    log(f"分析结果 行数: {len(df2)}")


if __name__ == "__main__":
    main()
