import xlrd
import re
from pathlib import Path

DATA_DIR = Path(__file__).resolve().parent.parent.parent / "data"
wb = xlrd.open_workbook(str(DATA_DIR / "2017国民经济行业分类注释（网络版）.xls"))
ws = wb.sheet_by_name("Sheet1")

def norm(v):
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        return s[:-2]
    return s

cur_sec = ""
print("=== 参考 XLS 门类/大类对应（75-90大类） ===")
for r in range(1, ws.nrows):
    c0 = norm(ws.cell_value(r, 0)).upper()
    name = str(ws.cell_value(r, 3)).strip()
    if re.fullmatch(r"[A-Z]", c0):
        cur_sec = c0
    elif re.fullmatch(r"\d{2}", c0):
        n = int(c0)
        if 75 <= n <= 90:
            print(f"  大类 {c0:3s} -> 门类={cur_sec}  名称: {name[:20] if name else '(空)'}")
        if n > 92:
            break
