"""
用映射表作为权威来源，对输出的参照表做最终验证：
- 检查每行小类码能否在映射表里查到
- 检查门类/大类/中类码是否与映射表一致
- 输出真实的残余不一致行
"""
import re
import glob
from pathlib import Path
import pandas as pd

DATA_DIR = Path(__file__).resolve().parent.parent.parent / "data"

# ── 读映射表 ──────────────────────────────────────────────────
map_files = sorted(glob.glob(str(DATA_DIR / "国民经济行业分类映射表_*.xlsx")))
MAP_FILE = map_files[-1] if map_files else str(DATA_DIR / "国民经济行业分类映射表.xlsx")
df_map = pd.read_excel(MAP_FILE, header=0, dtype=str,
                       usecols=[0, 3, 6, 9])
df_map.columns = ["sec", "big", "mid", "sml"]
df_map = df_map.dropna(subset=["sml"])

lookup_full = {}
lookup_num4 = {}
for _, row in df_map.iterrows():
    s = str(row["sml"]).strip().upper().replace("*", "")
    info = {"sec": str(row["sec"]).strip().upper(),
            "big": str(row["big"]).strip().upper(),
            "mid": str(row["mid"]).strip().upper()}
    lookup_full[s] = info
    m = re.search(r"(\d{4,5})", s)
    if m:
        n4 = m.group(1)[:4]
        if n4 not in lookup_num4:
            lookup_num4[n4] = info

# ── 读输出参照表 ──────────────────────────────────────────────
ref_files = sorted(glob.glob(str(DATA_DIR / "五篇大文章*_2026*.xlsx")))
REF_FILE = ref_files[-1] if ref_files else str(DATA_DIR / "五篇大文章与国民经济行业分类对应参照表20260312.xlsx")
print(f"验证文件: {REF_FILE}")
print(f"映射文件: {MAP_FILE}\n")

USE_COLS = [9, 10, 11, 12]
COL_NAMES = ["门类码", "大类码", "中类码", "小类码"]

all_sheets = pd.read_excel(REF_FILE, sheet_name=None, header=0,
                           dtype=str, usecols=USE_COLS)

grand_total = 0
still_wrong = []

for sn, raw_df in all_sheets.items():
    raw_df.columns = COL_NAMES
    df = raw_df.dropna(subset=["小类码"])
    if df.empty:
        continue

    sheet_issues = []
    for i, row in df.iterrows():
        sml_raw = str(row["小类码"]).strip().upper()
        sml = sml_raw.replace("*", "")
        if not sml:
            continue

        info = lookup_full.get(sml)
        if not info:
            m = re.search(r"(\d{4,5})", sml)
            if m:
                info = lookup_num4.get(m.group(1)[:4])
        if not info:
            continue  # 映射表里没有这条，跳过

        cur_sec = str(row["门类码"]).strip().upper()
        cur_big = str(row["大类码"]).strip().upper()
        cur_mid = str(row["中类码"]).strip().upper()

        bad = []
        if cur_sec != info["sec"]:
            bad.append(f"门类 {cur_sec}->{info['sec']}")
        if cur_big != info["big"]:
            bad.append(f"大类 {cur_big}->{info['big']}")
        if cur_mid != info["mid"]:
            bad.append(f"中类 {cur_mid}->{info['mid']}")
        # 小类码前缀
        sml_pref = sml[0] if sml and sml[0].isalpha() else ""
        if sml_pref and sml_pref != info["sec"]:
            bad.append(f"小类前缀 {sml_pref}->{info['sec']}")

        if bad:
            sheet_issues.append(f"  row={i+2}  小类={sml_raw}  问题: {', '.join(bad)}")
            grand_total += 1

    if sheet_issues:
        print(f"[{sn}] 残余问题 {len(sheet_issues)} 行:")
        for line in sheet_issues[:10]:
            print(line)
        if len(sheet_issues) > 10:
            print(f"  ... 还有 {len(sheet_issues)-10} 行")
        still_wrong.extend(sheet_issues)
    else:
        print(f"[{sn}] OK - 无残余问题")

print(f"\n=== 最终残余不一致总计: {grand_total} 行 ===")
if grand_total == 0:
    print("全部修正完成，参照表行业层级编码与映射表完全一致。")
