import re
import glob
from pathlib import Path
import pandas as pd

DATA_DIR = Path(__file__).resolve().parent.parent.parent / "data"
files = sorted(glob.glob(str(DATA_DIR / "五篇大文章*_2026*.xlsx")))
f = files[-1]
print("验证文件:", f)

# 只读相关列 (0-based: col9=门类, col10=大类, col11=中类, col12=小类, col13=门类名, col14=大类名)
USE_COLS = [9, 10, 11, 12, 13, 14, 16]
COL_NAMES = ["门类码", "大类码", "中类码", "小类码", "门类名", "大类名", "小类名"]

all_sheets = pd.read_excel(f, sheet_name=None, header=0, dtype=str,
                           usecols=USE_COLS)

print()
for sn, raw_df in all_sheets.items():
    raw_df.columns = COL_NAMES
    df = raw_df.dropna(subset=["小类码"])
    if df.empty:
        continue

    # 检查1: 84大类
    mask_84 = df["大类码"].str.upper().str.contains(r"[A-Z]84$", na=False, regex=True)
    rows_84 = df[mask_84]
    if not rows_84.empty:
        print(f"[{sn}] 84大类行 ({len(rows_84)}条):")
        for _, row in rows_84.iterrows():
            print(f"  门={row['门类码']} 大={row['大类码']} 中={row['中类码']} 小={row['小类码']} 大名={row['大类名']}")

    # 检查2: P门类下8xxx
    mask_p8 = (df["门类码"].str.upper() == "P") & \
              df["小类码"].str.upper().str.contains(r"[A-Z]8\d{3}", na=False, regex=True)
    rows_p8 = df[mask_p8]
    if not rows_p8.empty:
        print(f"[{sn}] P门类下8xxx ({len(rows_p8)}条) <- 可能有问题:")
        for _, row in rows_p8.iterrows():
            print(f"  门=P 大={row['大类码']} 小={row['小类码']}")

    # 检查3: K门类下7xxx（K是金融业，7xxx是房地产）
    mask_k7 = (df["门类码"].str.upper() == "K") & \
              df["小类码"].str.upper().str.contains(r"[A-Z]7\d{3}", na=False, regex=True)
    rows_k7 = df[mask_k7]
    if not rows_k7.empty:
        print(f"[{sn}] K门类下7xxx ({len(rows_k7)}条) <- 可能有问题:")
        for _, row in rows_k7.iterrows():
            print(f"  门=K 大={row['大类码']} 小={row['小类码']}")

    # 检查4: 小类码字母前缀与门类码不一致
    def sml_prefix(code):
        if pd.isna(code): return ""
        m = re.match(r"([A-Z])", str(code).strip().upper())
        return m.group(1) if m else ""

    df2 = df.copy()
    df2["sml_pref"] = df2["小类码"].apply(sml_prefix)
    df2["sec_pref"] = df2["门类码"].str.strip().str.upper()
    mask_prefix = (df2["sml_pref"] != "") & (df2["sec_pref"] != "") & \
                  (df2["sml_pref"] != df2["sec_pref"])
    rows_prefix = df2[mask_prefix]
    if not rows_prefix.empty:
        print(f"[{sn}] 小类码前缀与门类码不一致 ({len(rows_prefix)}条):")
        for _, row in rows_prefix.head(5).iterrows():
            print(f"  门={row['sec_pref']} != 小类前缀={row['sml_pref']} | 大={row['大类码']} 小={row['小类码']}")
        if len(rows_prefix) > 5:
            print(f"  ... 还有 {len(rows_prefix)-5} 条")

print("\n=== 验证完成 ===")
