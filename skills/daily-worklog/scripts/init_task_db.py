from pathlib import Path

ASSETS_DIR = Path(__file__).parent.parent / "assets"
import json
import os

import pandas as pd

EXCEL_DIR = r"E:\Downloads"
OUTPUT_FILE = str(ASSETS_DIR / "task_db.json")

def find_source_files(directory: str) -> list[str]:
    files = []
    for f in os.listdir(directory):
        if f.endswith(".xlsx") and "工时报表" in f and "(" in f:
            files.append(os.path.join(directory, f))
    return sorted(files)

def build_search_text(row: pd.Series) -> str:
    parts = []
    for col in ["工作项1", "工作项2", "工作项3", "工作日志", "项目名称", "业务线"]:
        val = row.get(col)
        if pd.notna(val) and str(val).strip():
            parts.append(str(val).strip())
    return " ".join(parts)

def main():
    files = find_source_files(EXCEL_DIR)
    if not files:
        raise FileNotFoundError(f"未在 {EXCEL_DIR} 找到符合条件的 Excel 文件")

    dfs = []
    for f in files:
        df = pd.read_excel(f)
        dfs.append(df)

    combined = pd.concat(dfs, ignore_index=True)

    # 去重：基于核心字段
    dedup_cols = ["申报类型", "申报日期", "业务集团", "业务线", "工作项1",
                  "工作项2", "工作项3", "项目名称", "工作日志", "投入时间(H)"]
    available_cols = [c for c in dedup_cols if c in combined.columns]
    combined = combined.drop_duplicates(subset=available_cols, keep="first")

    # 处理空值
    combined = combined.where(combined.notna(), None)

    # 建立 search_text
    combined["search_text"] = combined.apply(build_search_text, axis=1)

    # 输出
    records = combined.to_dict(orient="records")
    with open(OUTPUT_FILE, "w", encoding="utf-8") as fh:
        json.dump(records, fh, ensure_ascii=False, indent=2)

    print(f"共读取 {len(dfs)} 个文件，合并后去重共 {len(records)} 条记录")
    print(f"已输出到 {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
