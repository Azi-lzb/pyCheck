import os
import glob
from pathlib import Path
import pandas as pd

# 参照表在项目根下的 excel-data/data
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
DATA_DIR = PROJECT_ROOT / "excel-data" / "data"
files = [str(DATA_DIR / "五篇大文章与国民经济行业分类对应参照表20260312.xlsx")]
files += sorted(glob.glob(str(DATA_DIR / "五篇大文章*_2026*.xlsx")))
files = [f for f in files if os.path.isfile(f)]

for fpath in files:
    try:
        sheets = pd.read_excel(fpath, sheet_name=None, header=0, dtype=str)
    except Exception as e:
        print(f"  读取失败: {e}")
        continue

    print(f"\n=== {fpath} ===")
    for sn, df in sheets.items():
        cols = list(df.columns)
        if len(cols) < 13:
            continue
        # col1=产业类型, col2=产业大类编码, col10=门类, col11=大类, col13=小类
        c1, c2, c10, c11, c13 = cols[0], cols[1], cols[9], cols[10], cols[12]
        mask = df[c13].str.upper().str.contains("7517", na=False)
        rows = df[mask]
        if not rows.empty:
            print(f"  [{sn}] {len(rows)} 行含7517:")
            for _, row in rows.iterrows():
                print(f"    产业类型={row[c1]}  产业大类={row[c2]}  门类={row[c10]}  大类={row[c11]}  小类={row[c13]}")
