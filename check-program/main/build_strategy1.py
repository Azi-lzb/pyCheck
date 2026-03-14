import os
import re
from collections import defaultdict
from datetime import datetime

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.exceptions import InvalidFileException

# 可选：用于文件选择对话框（留空输入时弹出，Windows/带图形环境可用）
try:
    import tkinter as _tk
    from tkinter import filedialog as _filedialog
except ImportError:
    _tk = None
    _filedialog = None


# 程序所在目录（check-program/main）；config、结果文件及相对路径均基于此
ROOT = os.path.dirname(os.path.abspath(__file__))
# 项目根目录（pyCheck，含 excel-data、check-program），文件选择器默认从此处打开
PROJECT_ROOT = os.path.dirname(os.path.dirname(ROOT))
OUTPUT_FILE = os.path.join(ROOT, "策略一核查结果.xlsx")
CONFIG_FILE = os.path.join(ROOT, "config.xlsx")
# 术语：参照表 = 五篇大文章与国民经济行业分类对应参照表（config mapping 表 A2）；映射表 = 国民经济行业分类映射表（config mapping 表 B2）
# 映射表（国民经济行业分类映射表）：用于按行业小类取描述；路径优先从 config mapping B2 读取，否则用此处默认（可改为 ../excel-data/data/国民经济行业分类映射表.xlsx）
INDUSTRY_DESC_FILE = os.path.join(ROOT, "国民经济行业分类映射表.xlsx")
# 映射表内列：小类代码列（1-based J=10）、小类描述列（1-based L=12）
INDUSTRY_MAP_CODE_COL = 10
INDUSTRY_MAP_DESC_COL = 12


def _ask_open_file(title, filetypes=None, initialdir=None):
    """弹出系统文件选择对话框选择单个文件。无图形环境时返回 None。"""
    if _tk is None or _filedialog is None:
        return None
    if filetypes is None:
        filetypes = (
            ("Excel 文件", "*.xlsx *.xlsm *.xltx *.xltm"),
            ("所有文件", "*.*"),
        )
    root = _tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = _filedialog.askopenfilename(
        title=title,
        filetypes=filetypes,
        initialdir=initialdir or PROJECT_ROOT,
    )
    try:
        root.destroy()
    except Exception:
        pass
    return path if path else None


def _ask_open_files(title, filetypes=None, initialdir=None):
    """弹出系统文件选择对话框选择多个文件。无图形环境时返回 ()。"""
    if _tk is None or _filedialog is None:
        return ()
    if filetypes is None:
        filetypes = (
            ("Excel 文件", "*.xlsx *.xlsm *.xltx *.xltm"),
            ("所有文件", "*.*"),
        )
    root = _tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    paths = _filedialog.askopenfilenames(
        title=title,
        filetypes=filetypes,
        initialdir=initialdir or PROJECT_ROOT,
    )
    try:
        root.destroy()
    except Exception:
        pass
    return tuple(paths) if paths else ()

CONFIG_HEADERS = [
    "输出工作表名称",
    "来源台账文件",
    "台账类型",
    "表头行号",
    "数据起始行号",
    "贷款投向行业列序号",
    "机构报送产业分类列序号",
    "报送列-类别映射",
    "参照表工作表序号",
    "参照表产业分类代码列序号",
    "参照表行业4位码列序号",
    "参照表星标列序号",
    "参照表原始映射列序号",
]

TYPE_LABEL_TO_KEY = {"科技": "tech", "数字": "digital", "养老": "elder"}
TYPE_KEY_TO_LABEL = {"tech": "科技", "digital": "数字", "elder": "养老"}

# 重置 config 时若无源文件可推断，则使用以下默认配置（与 config 表实际数据一致，防止 config 丢失需重新配置）
DEFAULT_CONFIG_ROWS = [
    {
        "输出工作表名称": "科技产业贷款明细",
        "来源台账文件": r"伪数据\附科技",
        "台账类型": "科技",
        "表头行号": 3,
        "数据起始行号": 4,
        "贷款投向行业列序号": 15,
        "机构报送产业分类列序号": [17, 20, 23, 26],
        "报送列-类别映射": [(17, "HTP"), (20, "HTS"), (23, "SE"), (26, "PA")],
        "参照表工作表序号": 4,
        "参照表产业分类代码列序号": 2,
        "参照表行业4位码列序号": 13,
        "参照表星标列序号": 19,
        "参照表原始映射列序号": 21,
    },
    {
        "输出工作表名称": "数字经济产业贷款明细",
        "来源台账文件": r"伪数据\附数字",
        "台账类型": "数字",
        "表头行号": 4,
        "数据起始行号": 5,
        "贷款投向行业列序号": 15,
        "机构报送产业分类列序号": [16],
        "报送列-类别映射": [(16, "DE")],
        "参照表工作表序号": 2,
        "参照表产业分类代码列序号": 2,
        "参照表行业4位码列序号": 13,
        "参照表星标列序号": 19,
        "参照表原始映射列序号": 21,
    },
    {
        "输出工作表名称": "养老产业贷款明细",
        "来源台账文件": r"伪数据\附养老",
        "台账类型": "养老",
        "表头行号": 3,
        "数据起始行号": 4,
        "贷款投向行业列序号": 15,
        "机构报送产业分类列序号": [16],
        "报送列-类别映射": [(16, "EC")],
        "参照表工作表序号": 3,
        "参照表产业分类代码列序号": 2,
        "参照表行业4位码列序号": 13,
        "参照表星标列序号": 19,
        "参照表原始映射列序号": 21,
    },
]

CATEGORY_NAME_MAP = {
    "HTP": "高技术制造业",
    "HTS": "高技术服务业",
    "SE": "战略性新兴产业",
    "PA": "知识产权密集型产业",
    "DE": "数字经济产业",
    "EC": "养老产业",
}

# 线索规则表（config.xlsx 的 clue 工作表）：按条件匹配得到主标签、副标签、是否线索、是否疑似线索，通用六大产业
# 表头共 9 列（见 使用说明.md 第七节）
CLUE_HEADERS = [
    "编号",
    "匹配到的产业数量",
    "行业是否含星",
    "机构报送产业为空?",
    "命中情况",
    "主标签",
    "是否线索",
    "是否疑似线索",
    "备注/副标签",
]
# 默认 12 行规则（与线索默认规则表一致）；匹配顺序从上到下
# 表中「是否机构报送」= 是 → 机构有报送 → 本表「机构报送产业为空?」= 否；「是否机构报送」= 否 → 本表「机构报送产业为空?」= 是
CLUE_DEFAULT_ROWS = [
    [1, "0", "-", "是", "-", "正确", "否", "否", "行业未匹配到任何产业且机构报送为空，正确"],
    [2, "0", "-", "否", "-", "多报", "是", "是", "行业未匹配到任何产业且机构报送为不为空，多报"],
    [3, "1", "否", "否", "是", "正确", "否", "否", "行业匹配到的产业与报送产业一一对应，正确"],
    [4, "1", "是", "否", "是", "疑似正确", "否", "否", "行业*匹配到的产业与报送产业一一对应，疑似正确"],
    [5, "1", "否", "是", "-", "漏报", "是", "是", "行业匹配到产业但是机构未报送，漏报"],
    [6, "1", "是", "是", "-", "疑似漏报", "是", "是", "行业*匹配到产业但是机构未报送，疑似漏报"],
    [7, "1", "-", "否", "否", "错报", "是", "是", "行业（*）匹配到产业与机构报送产业不同，错报"],
    [8, ">1", "否", "否", "是", "正确", "否", "否", "行业匹配到的产业包含了机构报送的产业，正确"],
    [9, ">1", "是", "否", "是", "疑似正确", "否", "否", "行业*匹配到的产业包含了机构报送的产业，疑似正确"],
    [10, ">1", "否", "是", "-", "漏报", "是", "是", "行业匹配到多个产业但是机构未报送，漏报"],
    [11, ">1", "是", "是", "-", "疑似漏报", "是", "是", "行业*匹配到多个产业但是机构未报送，疑似漏报"],
    [12, ">1", "-", "否", "否", "错报", "是", "是", "行业（*）匹配到多个产业与机构报送产业不同，错报"],
]


def extract_industry4(text):
    if text is None:
        return ""
    s = str(text).upper()
    m = re.search(r"[A-Z]\d{4}", s)
    return m.group(0) if m else ""


def get_industry_desc_file_from_config():
    """从 config.xlsx 的 mapping 工作表读取 B2，即映射表（国民经济行业分类映射表）的文件路径。"""
    if not os.path.exists(CONFIG_FILE):
        return None
    wb = openpyxl.load_workbook(CONFIG_FILE, read_only=True, data_only=True)
    if "mapping" not in wb.sheetnames:
        wb.close()
        return None
    ws = wb["mapping"]
    path_cell = ws.cell(2, 2).value  # B2 = 映射表路径
    wb.close()
    if not path_cell:
        return None
    raw = str(path_cell).strip()
    if not raw:
        return None
    if os.path.isabs(raw):
        p = raw
    else:
        p = os.path.join(ROOT, raw)
    return p if os.path.exists(p) else None


def load_industry_desc_map():
    """从映射表（国民经济行业分类映射表）读取 小类代码->(显示用代码, 小类描述)，描述去掉空格和换行。

    例如 J 列为 `C2345*`，L 列为 `某某行业`，则字典中存储为：
    key='C2345*'，value=('C2345*', '某某行业')
    最终“行业小类描述”列会展示为 `C2345*：某某行业`。
    路径优先从 config 的 mapping 工作表 B2 读取。
    """
    file_path = get_industry_desc_file_from_config() or INDUSTRY_DESC_FILE
    if not os.path.exists(file_path):
        return {}
    d = {}
    code_idx = INDUSTRY_MAP_CODE_COL - 1
    desc_idx = INDUSTRY_MAP_DESC_COL - 1
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row is None or max(code_idx, desc_idx) >= len(row):
                continue
            code_raw = row[code_idx]
            desc_raw = row[desc_idx]
            if code_raw is None:
                continue
            display_code = str(code_raw).strip()  # 保留原样（含*等）
            code = display_code.upper()
            if not code:
                continue
            desc = "" if desc_raw is None else str(desc_raw)
            desc = re.sub(r"[\r\n\t]+", " ", desc)
            desc = re.sub(r"\s+", " ", desc).strip()
            value = (display_code, desc)
            d[code] = value
            # 若是类似 A0111，则同时支持用 0111 作为 key 查询
            if len(code) == 5 and code[0].isalpha() and code[1:].isdigit():
                d[code[1:]] = value
        wb.close()
    except Exception:
        pass
    return d


def extract_codes(text):
    if text is None:
        return []
    s = str(text).upper()
    codes = re.findall(r"[A-Z]{2,4}\d{2,6}", s)
    out = []
    seen = set()
    for code in codes:
        if code not in seen:
            seen.add(code)
            out.append(code)
    return out


def is_star_value(v):
    if v is None:
        return False
    s = str(v).strip().upper()
    if not s:
        return False
    if s in {"是", "Y", "YES", "TRUE", "1", "*"}:
        return True
    if s in {"否", "N", "NO", "FALSE", "0"}:
        return False
    return False


def parse_int(v, default=0):
    try:
        return int(str(v).strip())
    except Exception:
        return default


def parse_idx_list(v):
    if v is None:
        return []
    parts = re.split(r"[，,、\s]+", str(v).strip())
    out = []
    for p in parts:
        if p.isdigit():
            out.append(int(p))
    return out


def parse_col_category_map(v):
    """
    解析示例: 17:HTP,20:HTS,23:SE,26:PA
    返回: [(17,'HTP'), (20,'HTS'), ...]
    """
    if v is None:
        return []
    s = str(v).strip()
    if not s:
        return []
    out = []
    for part in re.split(r"[，,、\s]+", s):
        if not part or ":" not in part:
            continue
        col_s, cat_s = part.split(":", 1)
        col_s = col_s.strip()
        cat = cat_s.strip().upper()
        if col_s.isdigit() and cat:
            out.append((int(col_s), cat))
    return out


def parse_sheet_selector(v, default=1):
    """
    解析参照表工作表序号：既支持数字（序号），也支持字符串（sheet名称）。
    """
    if v is None:
        return default
    s = str(v).strip()
    if not s:
        return default
    if s.isdigit():
        return int(s)
    return s


def dedup_pairs_keep_order(pairs):
    out = []
    seen = set()
    for col, cat in pairs:
        key = (col, cat)
        if key not in seen:
            seen.add(key)
            out.append(key)
    return out


def dedup_keep_order(items):
    out = []
    seen = set()
    for x in items:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


def find_mapping_file():
    # 1. 优先从config.xlsx的mapping面板读取用户指定路径
    cfg_path = get_mapping_file_from_config()
    if cfg_path:
        return cfg_path

    # 2. 否则在根目录自动识别（工作表最多的xlsx）
    candidate = None
    max_sheets = -1
    for name in os.listdir(ROOT):
        if not name.lower().endswith(".xlsx"):
            continue
        if name.startswith("~$"):
            continue
        if name in {os.path.basename(OUTPUT_FILE), os.path.basename(CONFIG_FILE)}:
            continue
        path = os.path.join(ROOT, name)
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        sheet_count = len(wb.sheetnames)
        wb.close()
        if sheet_count > max_sheets:
            max_sheets = sheet_count
            candidate = path
    if not candidate:
        raise FileNotFoundError("未找到匹配表文件。")
    return candidate


def find_ledger_files():
    files = []
    for dp, _, fns in os.walk(ROOT):
        for fn in fns:
            if not fn.lower().endswith(".xlsx"):
                continue
            if fn.startswith("~$"):
                continue
            if fn in {os.path.basename(OUTPUT_FILE), os.path.basename(CONFIG_FILE)}:
                continue
            full = os.path.join(dp, fn)
            rel = os.path.relpath(full, ROOT)
            if rel.count("\\") == 0:
                continue
            if fn.lower() == "22.xlsx":
                continue
            wb = openpyxl.load_workbook(full, read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]
            mc, mr = ws.max_column, ws.max_row
            wb.close()
            if mc in (20, 21, 29) and mr >= 90:
                files.append(full)
    return sorted(files)


def detect_header_row(ws):
    best_row = 1
    best_count = -1
    for r in range(1, min(15, ws.max_row) + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        cnt = sum(v is not None and str(v).strip() != "" for v in vals)
        if cnt > best_count:
            best_count = cnt
            best_row = r
    return best_row


def classify_file_type(ws):
    mc = ws.max_column
    if mc >= 29:
        return "tech", [17, 20, 23, 26], 15
    if mc == 21:
        # 养老：只用第17列的小类编码作为产业分类列
        return "elder", [17], 15
    if mc == 20:
        # 数字经济：只用第17列的小类编码作为产业分类列
        return "digital", [17], 15
    raise ValueError(f"无法识别台账类型，列数={mc}")


def copy_sheet(src_ws, dst_ws):
    for r_idx, row in enumerate(src_ws.iter_rows(values_only=True), start=1):
        for c_idx, v in enumerate(row, start=1):
            dst_ws.cell(r_idx, c_idx).value = v


def ensure_unique_sheet_name(name, used):
    name = (name or "Sheet").strip()[:31] or "Sheet"
    if name not in used:
        used.add(name)
        return name
    base = name[:25]
    i = 2
    while f"{base}_{i}" in used:
        i += 1
    out = f"{base}_{i}"
    used.add(out)
    return out


def _default_col_category_map(file_type, reported_cols):
    if file_type == "tech":
        defaults = ["HTP", "HTS", "SE", "PA"]
        return [(col, defaults[i] if i < len(defaults) else "HTP") for i, col in enumerate(reported_cols)]
    if file_type == "digital":
        return [(col, "DE") for col in reported_cols]
    if file_type == "elder":
        return [(col, "EC") for col in reported_cols]
    return [(col, "ALL") for col in reported_cols]


def category_display_name(cat):
    return CATEGORY_NAME_MAP.get(cat.upper(), cat.upper())


# 已知主/副标签名，用于从备注解析副标签
CLUE_MARK_NAMES = ("正确", "疑似正确", "漏报", "疑似漏报", "多报", "错报", "疑似错报")


def _parse_sub_marks(主标签, 备注):
    """从 备注/副标签 列解析出副标签列表（不含主标签，去重）。"""
    if not 备注:
        return []
    s = str(备注).strip()
    out = []
    for name in CLUE_MARK_NAMES:
        if name != 主标签 and name in s:
            out.append(name)
    return out


def load_clue_rules():
    """
    从 config.xlsx 的 clue 工作表读取规则表。
    返回 list[dict]，每项: n_match, star, empty, hit, need_extra(从 Excel 加载时恒为 False), 主标签, 副标签(从 Excel 加载时恒为 [])。
    第 9 列「备注/副标签」仅供用户说明，不参与逻辑判断。
    """
    if not os.path.exists(CONFIG_FILE):
        return []
    wb = openpyxl.load_workbook(CONFIG_FILE, read_only=True, data_only=True)
    if "clue" not in wb.sheetnames:
        wb.close()
        return []
    ws = wb["clue"]
    # 表头顺序(9列): 编号(1), 匹配数量(2), 含星(3), 报送空(4), 命中(5), 主标签(6), 是否线索(7), 是否疑似线索(8), 备注/副标签(9，仅用户说明、不参与逻辑，但会写入结果表「备注-产业」列)
    key_n, key_star, key_empty, key_hit, key_main, key_clue, key_suspect, key_note = 2, 3, 4, 5, 6, 7, 8, 9
    rules = []
    for row_idx in range(2, ws.max_row + 1):
        n_val = ws.cell(row_idx, key_n).value
        star_val = ws.cell(row_idx, key_star).value
        empty_val = ws.cell(row_idx, key_empty).value
        hit_val = ws.cell(row_idx, key_hit).value
        main_val = ws.cell(row_idx, key_main).value
        clue_val = ws.cell(row_idx, key_clue).value
        suspect_val = ws.cell(row_idx, key_suspect).value
        note_val = ws.cell(row_idx, key_note).value if key_note else None
        if main_val is None or str(main_val).strip() == "":
            continue
        n_str = str(n_val).strip() if n_val is not None else ""
        star_str = str(star_val).strip() if star_val is not None else "-"
        empty_str = str(empty_val).strip() if empty_val is not None else "-"
        hit_str = str(hit_val).strip() if hit_val is not None else "-"
        main_str = str(main_val).strip()
        is_clue_str = str(clue_val).strip() if clue_val is not None else ""
        is_suspect_str = str(suspect_val).strip() if suspect_val is not None else ""
        note_str = str(note_val).strip() if note_val else ""
        # 备注/副标签(I列)不参与逻辑：不解析副标签，不根据「混报」设 need_extra；仅用于结果表「备注-产业」列展示
        need_extra = False
        sub_list = []
        rules.append({
            "n_match": n_str,
            "star": star_str,
            "empty": empty_str,
            "hit": hit_str,
            "need_extra": need_extra,
            "主标签": main_str,
            "副标签": sub_list,
            "是否线索": is_clue_str,
            "是否疑似线索": is_suspect_str,
            "备注": note_str,
        })
    wb.close()
    if rules:
        print("已从 config 加载 clue 规则：{} 条（文件：{}）".format(len(rules), CONFIG_FILE))
    else:
        print("已从 config 加载 clue 规则：0 条，使用内置默认规则（请检查 clue 表第6列主标签是否有内容、列顺序是否为 9 列）")
        rules = _default_clue_rules()
    return rules


def _default_clue_rules():
    """无 clue 表或解析失败时返回内置默认规则（由 CLUE_DEFAULT_ROWS 生成，与 clue 表实际数据一致）。"""
    rules = []
    for row in CLUE_DEFAULT_ROWS:
        # row: [编号, 匹配数量, 行业是否含星, 机构报送产业为空?, 命中情况, 主标签, 是否线索, 是否疑似线索, 备注]
        if len(row) < 9:
            continue
        rules.append({
            "n_match": str(row[1]),
            "star": str(row[2]) if row[2] is not None else "-",
            "empty": str(row[3]) if row[3] is not None else "-",
            "hit": str(row[4]) if row[4] is not None else "-",
            "need_extra": False,
            "主标签": str(row[5]).strip(),
            "副标签": [],
            "是否线索": str(row[6]).strip() if row[6] is not None else "",
            "是否疑似线索": str(row[7]).strip() if row[7] is not None else "",
            "备注": str(row[8]).strip() if row[8] is not None else "",
        })
    return rules


def apply_clue_rules(n_candidates, has_star, reported_empty, hit_non_empty, extra_non_empty, rules):
    """
    按顺序匹配规则：主标签/是否线索等取第一条命中；若有多条命中则备注拼接展示。
    返回 (主标签, 副标签list, 是否线索, 是否疑似线索, 触发列表)。
    触发列表 = [(规则编号1-based, 备注), ...]，所有命中的规则都会列入，便于在备注列拼接「触发规则编号 1：备注1；触发规则编号 2：备注2」。
    """
    n_str = "0" if n_candidates == 0 else ("1" if n_candidates == 1 else ">1")
    empty_str = "是" if reported_empty else "否"
    hit_str = "是" if hit_non_empty else "否"
    star_ok = lambda rule: rule["star"] in ("-", "是" if has_star else "否", "否/是")
    触发列表 = []
    主标签, 副标签, 是否线索, 是否疑似线索 = "", [], "", ""
    for idx, r in enumerate(rules):
        if r["n_match"] != n_str:
            continue
        if r["empty"] != "-" and r["empty"] != empty_str:
            continue
        if r["hit"] != "-" and r["hit"] != hit_str:
            continue
        if not star_ok(r):
            continue
        if r["need_extra"] and not extra_non_empty:
            continue
        if r["need_extra"] is False and extra_non_empty and r["hit"] == "是" and r["empty"] == "否" and r["n_match"] == ">1":
            continue
        备注 = (r.get("备注") or "").strip()
        触发列表.append((idx + 1, 备注))
        if len(触发列表) == 1:
            是否线索 = (r.get("是否线索") or "").strip()
            是否疑似线索 = (r.get("是否疑似线索") or "").strip()
            主标签, 副标签 = r["主标签"], list(r["副标签"])
    return (主标签, 副标签, 是否线索, 是否疑似线索, 触发列表)


def _create_empty_config_with_schema():
    """新建 config.xlsx：含 config 表头、mapping 全表（A1/B1、A2/B2）、log 表头。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "config"
    ws.append(CONFIG_HEADERS)
    for c in range(1, len(CONFIG_HEADERS) + 1):
        ws.cell(1, c).font = Font(name="Arial", bold=True)

    mapping_ws = wb.create_sheet("mapping")
    mapping_ws.append(["参照表路径", ""])
    mapping_ws.append(["映射表路径", ""])

    LOG_HEADERS = [
        "运行时间", "参照表路径", "参照表sheet数量", "参照表sheet列表",
        "结果文件", "明细sheet名称", "来源台账文件", "台账类型",
        "使用参照表sheet序号", "使用参照表sheet名称",
        "错报数量", "漏报数量", "疑似漏报数量",
    ]
    log_ws = wb.create_sheet("log")
    log_ws.append(LOG_HEADERS)
    for c in range(1, len(LOG_HEADERS) + 1):
        log_ws.cell(1, c).font = Font(name="Arial", bold=True)

    clue_ws = wb.create_sheet("clue")
    clue_ws.append(CLUE_HEADERS)
    for c in range(1, len(CLUE_HEADERS) + 1):
        clue_ws.cell(1, c).font = Font(name="Arial", bold=True)
    for row in CLUE_DEFAULT_ROWS:
        clue_ws.append(row)

    wb.save(CONFIG_FILE)
    wb.close()


def init_config_file_if_missing():
    if os.path.exists(CONFIG_FILE):
        return

    print("未找到 config.xlsx（程序根目录：{}）".format(ROOT))
    try:
        choice = input("是否新建 config 并初始化？(y/n)：").strip().lower()
    except EOFError:
        choice = "n"
    if choice not in ("y", "yes"):
        raise FileNotFoundError("未找到 config.xlsx。请先新建并初始化，或放入程序根目录。")

    # 先创建空 config（config 表头 + mapping 全表 + log 表头）
    _create_empty_config_with_schema()

    ledger_files = find_ledger_files()
    if not ledger_files:
        print("已创建空的 config.xlsx，请手动在 config 表填写台账配置，或在 mapping 表 B1/B2 填写参照表、映射表路径。")
        return

    # 根据台账文件补全 config 表数据行
    wb = openpyxl.load_workbook(CONFIG_FILE)
    ws = wb["config"]
    used_names = set()
    for f in ledger_files:
        src_wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        src_ws = src_wb[src_wb.sheetnames[0]]
        header_row = detect_header_row(src_ws)
        file_type, reported_cols, industry_col = classify_file_type(src_ws)
        output_sheet = ensure_unique_sheet_name(src_ws.title, used_names)
        ws.append(
            [
                output_sheet,
                os.path.relpath(f, ROOT),
                TYPE_KEY_TO_LABEL[file_type],
                header_row,
                header_row + 1,
                industry_col,
                ",".join(str(x) for x in reported_cols),
                ",".join(f"{x}:{infer}" for x, infer in _default_col_category_map(file_type, reported_cols)),
                {"digital": 2, "elder": 3, "tech": 4}[file_type],
                2,
                13,
                19,
                21,
            ]
        )
        src_wb.close()
    wb.save(CONFIG_FILE)
    wb.close()


def _ensure_mapping_sheet_layout():
    """若 mapping 表为旧格式（B1 为「国民经济行业分类映射表」），则迁移为 A1/B1、A2/B2 新格式。"""
    if not os.path.exists(CONFIG_FILE):
        return
    wb = openpyxl.load_workbook(CONFIG_FILE)
    if "mapping" not in wb.sheetnames:
        wb.close()
        return
    ws = wb["mapping"]
    if ws.max_row < 2:
        wb.close()
        return
    b1_val = str(ws.cell(1, 2).value or "").strip()
    if b1_val != "国民经济行业分类映射表":
        wb.close()
        return
    # 旧格式：A2=参照表路径, B2=映射表路径 → 新格式：B1=参照表路径, B2=映射表路径，A1/A2 为表头
    ref_path = str(ws.cell(2, 1).value or "").strip()
    map_path = str(ws.cell(2, 2).value or "").strip()
    ws.cell(1, 1).value = "参照表路径"
    ws.cell(1, 2).value = ref_path
    ws.cell(2, 1).value = "映射表路径"
    ws.cell(2, 2).value = map_path
    wb.save(CONFIG_FILE)
    wb.close()


def ensure_config_schema():
    if not os.path.exists(CONFIG_FILE):
        return
    _ensure_mapping_sheet_layout()
    wb = openpyxl.load_workbook(CONFIG_FILE)
    ws = wb[wb.sheetnames[0]]
    current_headers = [ws.cell(1, i).value for i in range(1, ws.max_column + 1)]
    if current_headers == CONFIG_HEADERS and "mapping" in wb.sheetnames and "log" in wb.sheetnames:
        wb.close()
        return

    normalized_rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None or all(v is None or str(v).strip() == "" for v in row):
            continue
        row = list(row)
        out_sheet = row[0] if len(row) > 0 else None
        source_file = row[1] if len(row) > 1 else None
        type_label = row[2] if len(row) > 2 else None
        header_row = row[3] if len(row) > 3 else None
        data_start = row[4] if len(row) > 4 else None
        industry_col = row[5] if len(row) > 5 else None
        reported_cols = row[6] if len(row) > 6 else None

        has_mapping_col = len(row) >= 12 and isinstance(row[7], str) and ":" in row[7]
        if has_mapping_col:
            col_cat_map_text = row[7]
            match_sheet = row[8] if len(row) > 8 else None
            class_col = row[9] if len(row) > 9 else None
            ind_col = row[10] if len(row) > 10 else None
            star_col = row[11] if len(row) > 11 else None
            raw_map_col = row[12] if len(row) > 12 else 21
        else:
            type_key = TYPE_LABEL_TO_KEY.get(str(type_label).strip(), "")
            rep_cols = parse_idx_list(reported_cols)
            col_cat_map_text = ",".join(f"{c}:{t}" for c, t in _default_col_category_map(type_key, rep_cols))
            match_sheet = row[7] if len(row) > 7 else None
            class_col = row[8] if len(row) > 8 else None
            ind_col = row[9] if len(row) > 9 else None
            star_col = row[10] if len(row) > 10 else None
            raw_map_col = row[12] if len(row) > 12 else 21

        normalized_rows.append(
            [
                out_sheet,
                source_file,
                type_label,
                header_row,
                data_start,
                industry_col,
                reported_cols,
                col_cat_map_text,
                match_sheet,
                class_col,
                ind_col,
                star_col,
                raw_map_col,
            ]
        )
    # 重写config sheet，并保留/创建mapping、log sheet（此步需要能写config.xlsx，若config被占用会在上层报错）
    if "mapping" in wb.sheetnames:
        mapping_ws = wb["mapping"]
        mapping_rows = [
            [cell for cell in row]
            for row in mapping_ws.iter_rows(min_row=1, max_row=mapping_ws.max_row, values_only=True)
        ]
        # 迁移旧格式：原 A1/B1 为表头、A2/B2 为路径 → 现 A1/B1 为参照表、A2/B2 为映射表
        if len(mapping_rows) >= 2 and str(mapping_rows[0][1] or "").strip() == "国民经济行业分类映射表":
            ref_path = str(mapping_rows[1][0] or "").strip() if len(mapping_rows[1]) > 0 else ""
            map_path = str(mapping_rows[1][1] or "").strip() if len(mapping_rows[1]) > 1 else ""
            mapping_rows = [["参照表路径", ref_path], ["映射表路径", map_path]] + mapping_rows[2:]
    else:
        mapping_rows = [["参照表路径", ""], ["映射表路径", ""]]

    if "log" in wb.sheetnames:
        log_ws = wb["log"]
        log_rows = [
            [cell for cell in row]
            for row in log_ws.iter_rows(min_row=1, max_row=log_ws.max_row, values_only=True)
        ]
    else:
        log_rows = []

    clue_rows = []
    if "clue" in wb.sheetnames:
        clue_ws_old = wb["clue"]
        for row in clue_ws_old.iter_rows(min_row=1, max_row=clue_ws_old.max_row, values_only=True):
            clue_rows.append([cell for cell in row])

    # 清空原工作簿，重建
    while wb.worksheets:
        wb.remove(wb.worksheets[0])

    new_ws = wb.create_sheet("config")
    new_ws.append(CONFIG_HEADERS)
    for c in range(1, len(CONFIG_HEADERS) + 1):
        new_ws.cell(1, c).font = Font(name="Arial", bold=True)
    for r in normalized_rows:
        new_ws.append(r)

    mapping_ws = wb.create_sheet("mapping")
    for row in mapping_rows:
        mapping_ws.append(row)

    if log_rows:
        log_ws = wb.create_sheet("log")
        for row in log_rows:
            log_ws.append(row)

    clue_ws = wb.create_sheet("clue")
    if clue_rows and len(clue_rows) >= 1:
        for row in clue_rows:
            clue_ws.append(row)
        for c in range(1, len(CLUE_HEADERS) + 1):
            clue_ws.cell(1, c).font = Font(name="Arial", bold=True)
    else:
        clue_ws.append(CLUE_HEADERS)
        for c in range(1, len(CLUE_HEADERS) + 1):
            clue_ws.cell(1, c).font = Font(name="Arial", bold=True)
        for row in CLUE_DEFAULT_ROWS:
            clue_ws.append(row)

    wb.save(CONFIG_FILE)
    wb.close()


def get_mapping_file_from_config():
    """从 config.xlsx 的 mapping 工作表读取 B1，即参照表（五篇大文章与国民经济行业分类对应参照表）的文件路径。"""
    if not os.path.exists(CONFIG_FILE):
        return None
    wb = openpyxl.load_workbook(CONFIG_FILE, read_only=True, data_only=True)
    if "mapping" not in wb.sheetnames:
        wb.close()
        return None
    ws = wb["mapping"]
    path_cell = ws.cell(1, 2).value  # B1 = 参照表路径
    wb.close()
    if not path_cell:
        return None
    raw = str(path_cell).strip()
    if not raw:
        return None
    # 允许绝对或相对路径
    if os.path.isabs(raw):
        p = raw
    else:
        p = os.path.join(ROOT, raw)
    return p if os.path.exists(p) else None


def get_source_file_map_from_mapping():
    """
    从 config.xlsx 的 mapping 工作表中读取源文件 → 源文件路径映射。
    期望结构：
    A1: 参照表路径   B1: 参照表（五篇大文章…）文件路径
    A2: 映射表路径   B2: 映射表（国民经济行业分类映射表）文件路径
    空一行（可选）
    A4: 源文件   B4: 源文件路径
    A5..: 源文件名称（或相对路径） B5..: 源文件完整路径
    """
    if not os.path.exists(CONFIG_FILE):
        return {}
    wb = openpyxl.load_workbook(CONFIG_FILE, read_only=True, data_only=True)
    if "mapping" not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb["mapping"]
    mapping = {}
    # 寻找“源文件”表头所在行
    header_row = None
    for r in range(1, ws.max_row + 1):
        v1 = ws.cell(r, 1).value
        v2 = ws.cell(r, 2).value
        if str(v1).strip() == "源文件" and str(v2).strip() == "源文件路径":
            header_row = r
            break
    if not header_row:
        wb.close()
        return {}
    for r in range(header_row + 1, ws.max_row + 1):
        name = ws.cell(r, 1).value
        path = ws.cell(r, 2).value
        if not name or not path:
            continue
        name_s = str(name).strip()
        path_s = str(path).strip()
        if not name_s or not path_s:
            continue
        mapping[name_s] = path_s
    wb.close()
    return mapping


# openpyxl 仅支持以下格式，.xls 等需用 xlrd 或先转为 xlsx
OPENPYXL_SUPPORTED_SUFFIXES = (".xlsx", ".xlsm", ".xltx", ".xltm")


def build_config_rows_from_files(file_paths):
    """
    根据给定的源文件路径列表，自动推断config行（台账类型、列序号等）。
    仅处理 openpyxl 支持的格式（.xlsx/.xlsm/.xltx/.xltm），跳过 .xls 等并提示。
    """
    rows = []
    for p in sorted(file_paths):
        if not p:
            continue
        path = str(p).strip()
        if not path:
            continue
        if not os.path.isabs(path):
            path_abs = os.path.join(ROOT, path)
        else:
            path_abs = path
        if not os.path.exists(path_abs):
            continue
        if not path_abs.lower().endswith(OPENPYXL_SUPPORTED_SUFFIXES):
            print("跳过（openpyxl 不支持该格式，仅支持 .xlsx/.xlsm/.xltx/.xltm）：{}".format(path_abs))
            continue
        try:
            wb = openpyxl.load_workbook(path_abs, read_only=True, data_only=True)
        except InvalidFileException:
            print("跳过（无法用 openpyxl 打开，请用 Excel 另存为 .xlsx 后重试）：{}".format(path_abs))
            continue
        ws = wb[wb.sheetnames[0]]
        header_row = detect_header_row(ws)
        file_type, reported_cols, industry_col = classify_file_type(ws)
        output_sheet = ws.title
        wb.close()

        rows.append(
            {
                "输出工作表名称": output_sheet,
                "来源台账文件": os.path.relpath(path_abs, ROOT),
                "台账类型": file_type,
                "表头行号": header_row,
                "数据起始行号": header_row + 1,
                "贷款投向行业列序号": industry_col,
                "机构报送产业分类列序号": reported_cols,
                "报送列-类别映射": _default_col_category_map(file_type, reported_cols),
                "参照表工作表序号": {"digital": 2, "elder": 3, "tech": 4}[file_type],
                "参照表产业分类代码列序号": 2,
                "参照表行业4位码列序号": 13,
                "参照表星标列序号": 19,
                "参照表原始映射列序号": 21,
            }
        )
    return rows

def append_run_log(mapping_file, mapping_sheetnames, executed_entries, output_path):
    if not executed_entries or not os.path.exists(CONFIG_FILE):
        return
    wb = openpyxl.load_workbook(CONFIG_FILE)
    if "log" in wb.sheetnames:
        ws = wb["log"]
        start_row = ws.max_row + 1
    else:
        ws = wb.create_sheet("log")
        ws.append(
            [
                "运行时间",
                "参照表路径",
                "参照表sheet数量",
                "参照表sheet列表",
                "结果文件",
                "明细sheet名称",
                "来源台账文件",
                "台账类型",
                "使用参照表sheet序号",
                "使用参照表sheet名称",
                "错报数量",
                "漏报数量",
                "疑似漏报数量",
            ]
        )
        start_row = 2

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sheet_list_str = ",".join(mapping_sheetnames)
    for entry in executed_entries:
        sel = entry["参照表工作表序号"]
        if isinstance(sel, int) and 1 <= sel <= len(mapping_sheetnames):
            mname = mapping_sheetnames[sel - 1]
        elif isinstance(sel, str) and sel.strip():
            mname = sel.strip()
        else:
            mname = ""
        ws.append(
            [
                ts,
                mapping_file,
                len(mapping_sheetnames),
                sheet_list_str,
                output_path,
                entry["输出工作表名称"],
                entry["来源台账文件"],
                TYPE_KEY_TO_LABEL.get(entry["台账类型"], entry["台账类型"]),
                sel,
                mname,
                entry.get("错报数量", ""),
                entry.get("漏报数量", ""),
                entry.get("疑似漏报数量", ""),
            ]
        )

    wb.save(CONFIG_FILE)
    wb.close()


def load_config_rows():
    if not os.path.exists(CONFIG_FILE):
        raise FileNotFoundError("未找到config.xlsx。")
    wb = openpyxl.load_workbook(CONFIG_FILE, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None or all(v is None or str(v).strip() == "" for v in row):
            continue
        row = list(row) + [None] * max(0, 13 - len(row))
        output_sheet = str(row[0]).strip() if len(row) > 0 and row[0] is not None else ""
        source_file = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
        file_type_label = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
        file_type = TYPE_LABEL_TO_KEY.get(file_type_label, "")
        if not output_sheet or not source_file or not file_type:
            continue
        reported_cols = parse_idx_list(row[6])
        col_cat_map = dedup_pairs_keep_order(parse_col_category_map(row[7]))
        if not col_cat_map:
            col_cat_map = _default_col_category_map(file_type, reported_cols)
        # 养老、数字默认只使用最后一列（小类编码）作为报送列
        if file_type in {"digital", "elder"} and len(reported_cols) > 1:
            reported_cols = [reported_cols[-1]]
        rows.append(
            {
                "输出工作表名称": output_sheet,
                "来源台账文件": source_file,
                "台账类型": file_type,
                "表头行号": parse_int(row[3], 1),
                "数据起始行号": parse_int(row[4], 2),
                "贷款投向行业列序号": parse_int(row[5], 15),
                "机构报送产业分类列序号": reported_cols,
                "报送列-类别映射": col_cat_map,
                "参照表工作表序号": parse_sheet_selector(row[8], 4),
                "参照表产业分类代码列序号": parse_int(row[9], 2),
                "参照表行业4位码列序号": parse_int(row[10], 13),
                "参照表星标列序号": parse_int(row[11], 19),
                "参照表原始映射列序号": parse_int(row[12], 21),
            }
        )
    wb.close()
    return rows


def build_mapping_by_config(mapping_file, cfg):
    wb = openpyxl.load_workbook(mapping_file, read_only=True, data_only=True)
    sel = cfg["参照表工作表序号"]
    sheetnames = wb.sheetnames
    # 支持数字序号或工作表名称
    if isinstance(sel, int):
        sheet_idx = sel - 1
        if sheet_idx < 0 or sheet_idx >= len(sheetnames):
            wb.close()
            return defaultdict(dict)
        ws = wb[sheetnames[sheet_idx]]
    else:
        name = str(sel).strip()
        if not name or name not in sheetnames:
            wb.close()
            return defaultdict(dict)
        ws = wb[name]
    code_idx = cfg["参照表产业分类代码列序号"] - 1
    ind_idx = cfg["参照表行业4位码列序号"] - 1
    star_idx = cfg["参照表星标列序号"] - 1
    raw_idx = cfg.get("参照表原始映射列序号", 21) - 1  # 参照表 U 列=21，存小类\中类\大类\门类等原始映射内容（与映射表 B2 无关）
    d = defaultdict(dict)

    for row in ws.iter_rows(min_row=2, values_only=True):
        row = list(row) if row is not None else []
        if max(code_idx, ind_idx, star_idx) >= len(row):
            continue
        class_code = row[code_idx]
        ind4 = row[ind_idx]
        star_raw = row[star_idx]
        raw_mapping = ""
        if raw_idx < len(row) and row[raw_idx] is not None:
            raw_mapping = str(row[raw_idx]).strip()
        class_code = str(class_code).strip().upper() if class_code is not None else ""
        ind4 = str(ind4).strip().upper() if ind4 is not None else ""
        if not class_code or not re.fullmatch(r"[A-Z]{2,4}\d{2,6}", class_code):
            continue
        if not re.fullmatch(r"[A-Z]\d{4}", ind4):
            m = re.search(r"[A-Z]\d{4}", ind4)
            if not m:
                continue
            ind4 = m.group(0)
        star = is_star_value(star_raw)
        old = d[ind4].get(class_code)
        # 存 (是否星标, 原始映射内容列表)，同代码多条时保留全部层级映射（如小类/中类/大类）
        if old is None:
            raws = [raw_mapping] if raw_mapping else []
            d[ind4][class_code] = (star, raws)
        else:
            old_star, old_raws = old
            if isinstance(old_raws, list):
                raws = old_raws[:]
            elif isinstance(old_raws, str) and old_raws.strip():
                raws = [old_raws.strip()]
            else:
                raws = []
            if raw_mapping and raw_mapping not in raws:
                raws.append(raw_mapping)
            d[ind4][class_code] = (old_star or star, raws)

    wb.close()
    return d


def write_results():
    init_config_file_if_missing()
    ensure_config_schema()
    source_file_map = get_source_file_map_from_mapping()
    config_rows = load_config_rows()
    # 若config表中没有有效行，则尝试根据mapping中的源文件路径自动构造
    if not config_rows and source_file_map:
        config_rows = build_config_rows_from_files(source_file_map.values())
    # 仍然没有，则退回扫描目录自动识别
    if not config_rows:
        auto_files = find_ledger_files()
        config_rows = build_config_rows_from_files(auto_files)
    if not config_rows:
        raise ValueError("未能从config.xlsx或目录中推断任何台账配置行。")

    mapping_file = find_mapping_file()
    # 预读取参照表sheet信息，便于日志记录
    map_wb = openpyxl.load_workbook(mapping_file, read_only=True, data_only=True)
    mapping_sheetnames = map_wb.sheetnames
    map_wb.close()

    out_wb = Workbook()
    out_wb.remove(out_wb.active)
    used_names = set()
    mapping_cache = {}
    executed_entries = []
    industry_desc_map = load_industry_desc_map()
    clue_rules = load_clue_rules()

    for cfg in config_rows:
        # 解析源文件路径：优先mapping面板中的“源文件路径”，再退回相对路径
        src_key = cfg.get("来源台账文件")
        if src_key is None or not str(src_key).strip():
            print(f"跳过：来源台账文件为空 -> 输出工作表名称「{cfg.get('输出工作表名称', '')}」")
            continue
        src_key = str(src_key).strip()
        if os.path.isabs(src_key):
            source_abs = src_key
        else:
            mapped = None
            # 先尝试以“源文件”列的值直接匹配
            if src_key in source_file_map:
                mapped = source_file_map[src_key]
            else:
                # 再按文件名匹配
                for k, v in source_file_map.items():
                    if os.path.basename(k) == os.path.basename(src_key):
                        mapped = v
                        break
            if mapped:
                source_abs = mapped
            else:
                source_abs = os.path.join(ROOT, src_key)
        if not os.path.exists(source_abs):
            print(f"跳过：来源文件不存在 -> {cfg['来源台账文件']}")
            continue

        src_wb = openpyxl.load_workbook(source_abs, read_only=True, data_only=True)
        src_ws = src_wb[src_wb.sheetnames[0]]
        dst_name = ensure_unique_sheet_name(cfg["输出工作表名称"], used_names)
        dst_ws = out_wb.create_sheet(dst_name)
        copy_sheet(src_ws, dst_ws)

        header_row = cfg["表头行号"]
        data_start_row = cfg["数据起始行号"] or (header_row + 1)
        industry_col = cfg["贷款投向行业列序号"]
        reported_cols = cfg["机构报送产业分类列序号"]
        col_cat_map = dedup_pairs_keep_order(cfg.get("报送列-类别映射", []))
        if not col_cat_map:
            col_cat_map = _default_col_category_map(cfg["台账类型"], reported_cols)
        category_order = dedup_keep_order([cat for _, cat in col_cat_map])

        start_col = src_ws.max_column + 1
        cache_key = (
            cfg["参照表工作表序号"],
            cfg["参照表产业分类代码列序号"],
            cfg["参照表行业4位码列序号"],
            cfg["参照表星标列序号"],
            cfg.get("参照表原始映射列序号", 21),
        )
        if cache_key not in mapping_cache:
            mapping_cache[cache_key] = build_mapping_by_config(mapping_file, cfg)
        mp = mapping_cache[cache_key]

        headers = []
        for cat in category_order:
            c_name = category_display_name(cat)
            headers.extend(
                [
                    f"报送-{c_name}",
                    f"匹配-{c_name}",
                    f"疑似正确-{c_name}",
                    f"错报-{c_name}",
                    f"漏报-{c_name}",
                    f"疑似漏报-{c_name}",
                    f"是否线索-{c_name}",
                    f"是否疑似线索-{c_name}",
                    f"备注-{c_name}",
                ]
            )
        headers.append("行业小类描述")
        for i, h in enumerate(headers):
            cell = dst_ws.cell(header_row, start_col + i)
            cell.value = h
            cell.font = Font(name="Arial", bold=True)

        count_by_cat = {cat: {"错报": 0, "miss": 0, "sus_miss": 0} for cat in category_order}

        for r in range(data_start_row, src_ws.max_row + 1):
            industry4 = extract_industry4(src_ws.cell(r, industry_col).value)
            match_map = mp.get(industry4, {})
            if (
                not match_map
                and industry4
                and len(industry4) == 5
                and industry4[0].isalpha()
                and industry4[1:].isdigit()
            ):
                _de = industry_desc_map.get(industry4[1:])
                if isinstance(_de, tuple) and _de[0]:
                    _canonical = _de[0].upper().replace("*", "")
                    if _canonical != industry4:
                        match_map = mp.get(_canonical, match_map)
            col_offset = 0
            for cat in category_order:
                reported_cat = []
                for col, mapped_cat in col_cat_map:
                    if mapped_cat != cat:
                        continue
                    if col <= src_ws.max_column:
                        reported_cat.extend(extract_codes(src_ws.cell(r, col).value))
                rpt = dedup_keep_order(reported_cat)

                match_codes = [code for code in match_map.keys() if code.startswith(cat)]
                def _star(mc):
                    v = match_map[mc]
                    return v[0] if isinstance(v, tuple) else bool(v)
                def _raw_list(mc):
                    v = match_map[mc]
                    if not (isinstance(v, tuple) and len(v) > 1):
                        return []
                    raw_v = v[1]
                    if isinstance(raw_v, list):
                        return [x for x in raw_v if str(x).strip()]
                    if isinstance(raw_v, str) and raw_v.strip():
                        return [raw_v.strip()]
                    return []
                nonstar_codes = [c for c in match_codes if not _star(c)]
                star_codes = [c for c in match_codes if _star(c)]

                def _star_display_parts(mc):
                    raws = _raw_list(mc)
                    if raws:
                        return [f"【{mc}*:{raw}】" for raw in raws]
                    return [f"【{mc}*】"]
                def _nonstar_display_parts(mc):
                    raws = _raw_list(mc)
                    if raws:
                        return [f"【{mc}:{raw}】" for raw in raws]
                    return [f"【{mc}】"]
                match_display = []
                for c in match_codes:
                    match_display.extend(_star_display_parts(c) if _star(c) else _nonstar_display_parts(c))
                reported_set = set(rpt)

                unmatched_reported = [c for c in rpt if c not in match_codes]
                matched_star_reported = [c for c in rpt if c in star_codes]
                missing_nonstar = [c for c in nonstar_codes if c not in reported_set]
                missing_star = [c for c in star_codes if c not in reported_set]

                multi = ""
                suspect_multi = ""
                miss = ""
                suspect_miss = ""

                if unmatched_reported:
                    multi = "、".join(f"【{c}】" for c in unmatched_reported)
                if matched_star_reported:
                    parts = []
                    for c in matched_star_reported:
                        parts.extend(_star_display_parts(c))
                    suspect_multi = "、".join(parts)
                if missing_nonstar:
                    parts = []
                    for c in missing_nonstar:
                        parts.extend(_nonstar_display_parts(c))
                    miss = "、".join(parts)
                if missing_star:
                    parts = []
                    for c in missing_star:
                        parts.extend(_star_display_parts(c))
                    suspect_miss = "、".join(parts)

                # 按 clue 表规则得到主标签与副标签（通用六大产业）
                n_cand = len(match_codes)
                # 行业是否含星：与机构报送绑定——机构在本产业下报送的代码中，是否有落在该行业该产业下带*候选里的
                has_star_cat = len(reported_set & set(star_codes)) > 0
                reported_empty_cat = len(rpt) == 0
                hit_non_empty_cat = len(reported_set & set(match_codes)) > 0
                extra_non_empty_cat = len(reported_set - set(match_codes)) > 0
                主标签, 副标签, 规则是否线索, 规则是否疑似线索, 触发列表 = apply_clue_rules(
                    n_cand, has_star_cat, reported_empty_cat, hit_non_empty_cat, extra_non_empty_cat, clue_rules
                )
                is_疑似正确 = 主标签 == "疑似正确" or "疑似正确" in 副标签
                疑似正确内容 = "、".join(match_display) if is_疑似正确 else ""
                # 是否线索/是否疑似线索：优先用 clue 表第7、8列；为空时按主/副标签推导
                if 规则是否线索 in ("是", "否"):
                    is_clue = 规则是否线索 == "是"
                else:
                    is_clue = 主标签 in ("漏报", "多报", "错报") or any(m in ("漏报", "多报", "错报") for m in 副标签)
                if 规则是否疑似线索 in ("是", "否"):
                    is_suspect_clue = 规则是否疑似线索 == "是"
                else:
                    is_suspect_clue = 主标签 in ("疑似漏报", "疑似多报", "疑似错报", "疑似正确") or any(
                        m in ("疑似漏报", "疑似多报", "疑似错报", "疑似正确") for m in 副标签
                    )

                错报内容 = "；".join(s for s in (multi, suspect_multi) if s)
                dst_ws.cell(r, start_col + col_offset).value = "、".join(rpt)
                dst_ws.cell(r, start_col + col_offset + 1).value = "、".join(match_display)
                dst_ws.cell(r, start_col + col_offset + 2).value = 疑似正确内容
                dst_ws.cell(r, start_col + col_offset + 3).value = 错报内容
                dst_ws.cell(r, start_col + col_offset + 4).value = miss
                dst_ws.cell(r, start_col + col_offset + 5).value = suspect_miss
                dst_ws.cell(r, start_col + col_offset + 6).value = "是" if is_clue else ""
                dst_ws.cell(r, start_col + col_offset + 7).value = "是" if is_suspect_clue else ""
                备注显示 = "；".join("触发规则编号 {}：{}".format(rid, note) for rid, note in 触发列表) if 触发列表 else ""
                dst_ws.cell(r, start_col + col_offset + 8).value = 备注显示

                if 错报内容:
                    count_by_cat[cat]["错报"] += 1
                if miss:
                    count_by_cat[cat]["miss"] += 1
                if suspect_miss:
                    count_by_cat[cat]["sus_miss"] += 1

                col_offset += 9

            desc_entry = industry_desc_map.get(industry4) or (
                industry_desc_map.get(industry4[1:])
                if (
                    industry4
                    and len(industry4) == 5
                    and industry4[0].isalpha()
                    and industry4[1:].isdigit()
                )
                else None
            )
            if isinstance(desc_entry, tuple):
                code_disp, desc_text = desc_entry
                if desc_text:
                    industry_desc = f"{code_disp}：{desc_text}"
                else:
                    industry_desc = code_disp
            else:
                industry_desc = desc_entry or ""
            dst_ws.cell(r, start_col + col_offset).value = industry_desc

        src_wb.close()

        # 汇总为 log 用字符串：产业名：数量；产业名：数量；…
        def count_str(key):
            return "；".join(f"{category_display_name(cat)}：{count_by_cat[cat][key]}" for cat in category_order)

        executed_entries.append(
            {
                "输出工作表名称": dst_name,
                "来源台账文件": cfg["来源台账文件"],
                "台账类型": cfg["台账类型"],
                "参照表工作表序号": cfg["参照表工作表序号"],
                "错报数量": count_str("错报"),
                "漏报数量": count_str("miss"),
                "疑似漏报数量": count_str("sus_miss"),
            }
        )

    # 若所有配置行均被跳过（如来源文件不存在），out_wb 无任何 sheet，openpyxl 保存会报 At least one sheet must be visible
    if not out_wb.worksheets:
        ws_info = out_wb.create_sheet("说明", 0)
        ws_info["A1"] = "所有配置的台账来源文件均不存在或已跳过，未生成任何明细表。"
        ws_info["A2"] = "请检查 config 表「来源台账文件」及 mapping 表「源文件路径」，确保文件存在后再执行策略一。"

    for ws in out_wb.worksheets:
        ws.sheet_view.zoomScale = 90

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    actual_output = os.path.join(ROOT, f"策略一核查结果_{ts}.xlsx")
    out_wb.save(actual_output)
    out_wb.close()
    append_run_log(mapping_file, mapping_sheetnames, executed_entries, actual_output)
    return mapping_file, actual_output, CONFIG_FILE
    

def cli_set_mapping_path():
    """配置参照表路径，写入 mapping 表 B1。"""
    path = _ask_open_file("选择参照表文件（五篇大文章与国民经济行业分类对应参照表）")
    if path is None and _tk is None:
        path = input("当前环境无法打开文件选择器，请输入参照表路径（留空取消）：").strip()
    if not path:
        print("已取消设置参照表路径。")
        return
    if not os.path.isabs(path):
        candidate = os.path.join(ROOT, path)
        if os.path.exists(candidate):
            path = candidate
    wb = openpyxl.load_workbook(CONFIG_FILE) if os.path.exists(CONFIG_FILE) else Workbook()
    if not wb.sheetnames:
        wb.create_sheet("config")
    if "mapping" not in wb.sheetnames:
        ws = wb.create_sheet("mapping")
        ws.append(["参照表路径", path])
        ws.append(["映射表路径", ""])
    else:
        ws = wb["mapping"]
        if ws.max_row < 1:
            ws.append(["参照表路径", path])
            ws.append(["映射表路径", ""])
        elif ws.max_row < 2:
            ws.cell(1, 1).value = "参照表路径"
            ws.cell(1, 2).value = path
            ws.append(["映射表路径", ""])
        else:
            ws.cell(1, 1).value = "参照表路径"
            ws.cell(1, 2).value = path
            if str(ws.cell(2, 1).value or "").strip() != "映射表路径":
                ws.cell(2, 1).value = "映射表路径"
    wb.save(CONFIG_FILE)
    wb.close()
    print("参照表路径已更新为：", path)


def cli_set_source_files():
    paths = _ask_open_files("选择源文件（可多选）")
    if not paths and _tk is None:
        print("请输入源文件路径，多个用分号 ; 分隔。")
        raw = input("源文件路径列表：").strip()
        paths = [p.strip() for p in raw.split(";") if p.strip()]
    if not paths:
        print("已取消或未选择任何源文件。")
        return
    wb = openpyxl.load_workbook(CONFIG_FILE) if os.path.exists(CONFIG_FILE) else Workbook()
    if not wb.sheetnames:
        wb.create_sheet("config")
    if "mapping" not in wb.sheetnames:
        ws = wb.create_sheet("mapping")
        ws.append(["参照表路径", ""])
        ws.append(["映射表路径", ""])
        ws.append(["源文件", "源文件路径"])
    else:
        ws = wb["mapping"]
        # 查找或创建 “源文件 / 源文件路径” 表头
        header_row = None
        for r in range(1, ws.max_row + 1):
            v1 = ws.cell(r, 1).value
            v2 = ws.cell(r, 2).value
            if str(v1).strip() == "源文件" and str(v2).strip() == "源文件路径":
                header_row = r
                break
        if header_row is None:
            header_row = ws.max_row + 1
            ws.cell(header_row, 1).value = "源文件"
            ws.cell(header_row, 2).value = "源文件路径"
        # 清理旧的源文件行
        for r in range(header_row + 1, ws.max_row + 1):
            ws.cell(r, 1).value = None
            ws.cell(r, 2).value = None
        ws.delete_rows(header_row + 1, ws.max_row - header_row)
    # 重新填充源文件列表
    ws = wb["mapping"]
    header_row = None
    for r in range(1, ws.max_row + 1):
        v1 = ws.cell(r, 1).value
        v2 = ws.cell(r, 2).value
        if str(v1).strip() == "源文件" and str(v2).strip() == "源文件路径":
            header_row = r
            break
    row = header_row + 1
    for p in paths:
        ws.cell(row, 1).value = os.path.basename(p)
        ws.cell(row, 2).value = p
        row += 1
    wb.save(CONFIG_FILE)
    wb.close()
    print(f"已更新 {len(paths)} 个源文件路径。")


def cli_set_industry_map_path():
    """配置映射表（国民经济行业分类映射表）路径，写入 mapping 表 B2。"""
    path = _ask_open_file("选择国民经济行业分类映射表文件")
    if path is None and _tk is None:
        path = input("当前环境无法打开文件选择器，请输入映射表路径（留空取消）：").strip()
    if not path:
        print("已取消设置国民经济行业分类映射表路径。")
        return
    if not os.path.isabs(path):
        candidate = os.path.join(ROOT, path)
        if os.path.exists(candidate):
            path = candidate
    wb = openpyxl.load_workbook(CONFIG_FILE) if os.path.exists(CONFIG_FILE) else Workbook()
    if not wb.sheetnames:
        wb.create_sheet("config")
    if "mapping" not in wb.sheetnames:
        ws = wb.create_sheet("mapping")
        ws.append(["参照表路径", ""])
        ws.append(["映射表路径", path])
    else:
        ws = wb["mapping"]
        if ws.max_row < 1:
            ws.append(["参照表路径", ""])
            ws.append(["映射表路径", path])
        elif ws.max_row < 2:
            ws.cell(1, 1).value = "参照表路径"
            ws.cell(1, 2).value = ws.cell(1, 2).value or ""
            ws.append(["映射表路径", path])
        else:
            ws.cell(2, 1).value = "映射表路径"
            ws.cell(2, 2).value = path
    wb.save(CONFIG_FILE)
    wb.close()
    print("国民经济行业分类映射表路径已更新为：", path)


def reset_config_sheet():
    cfg_path = CONFIG_FILE
    if not os.path.exists(cfg_path):
        print("未找到config.xlsx，无法重置config页。")
        return
    wb = openpyxl.load_workbook(cfg_path)
    source_map = get_source_file_map_from_mapping()
    paths = list(source_map.values())
    if not paths:
        paths = find_ledger_files()
    rows = build_config_rows_from_files(paths)
    if not rows:
        rows = DEFAULT_CONFIG_ROWS
        print("未从源文件推断到配置行，已使用内置默认配置（科技/数字/养老三表）重置。")
    mapping_rows = []
    if "mapping" in wb.sheetnames:
        ws_m = wb["mapping"]
        for row in ws_m.iter_rows(min_row=1, max_row=ws_m.max_row, values_only=True):
            mapping_rows.append(list(row))
    log_rows = []
    if "log" in wb.sheetnames:
        ws_l = wb["log"]
        for row in ws_l.iter_rows(min_row=1, max_row=ws_l.max_row, values_only=True):
            log_rows.append(list(row))
    clue_rows = []
    if "clue" in wb.sheetnames:
        ws_clue = wb["clue"]
        for row in ws_clue.iter_rows(min_row=1, max_row=ws_clue.max_row, values_only=True):
            clue_rows.append(list(row))
    for s in list(wb.sheetnames):
        wb.remove(wb[s])
    ws_cfg = wb.create_sheet("config")
    ws_cfg.append(CONFIG_HEADERS)
    for c in range(1, len(CONFIG_HEADERS) + 1):
        ws_cfg.cell(1, c).font = Font(name="Arial", bold=True)
    for r in rows:
        type_label = TYPE_KEY_TO_LABEL.get(r["台账类型"], r["台账类型"])
        ws_cfg.append(
            [
                r["输出工作表名称"],
                r["来源台账文件"],
                type_label,
                r["表头行号"],
                r["数据起始行号"],
                r["贷款投向行业列序号"],
                ",".join(str(x) for x in r["机构报送产业分类列序号"]),
                ",".join(f"{c}:{t}" for c, t in r["报送列-类别映射"]),
                r["参照表工作表序号"],
                r["参照表产业分类代码列序号"],
                r["参照表行业4位码列序号"],
                r["参照表星标列序号"],
                r.get("参照表原始映射列序号", 21),
            ]
        )
    # 在第5行写入各表头含义说明（说明行会被加载逻辑自动忽略）
    desc_row = [
        "说明：结果文件中的 sheet 名称，例如“科技产业贷款明细核查”",
        "说明：来源台账文件路径，相对项目根目录或绝对路径",
        "说明：台账类型，只能填“科技/数字/养老”三种之一；脚本会据此推断默认的参照表 sheet 以及“报送列-类别映射”等参数；若填其他值（如 tech/digital/elder），该行配置会被跳过，整张台账不会执行策略一，也不会套用默认规则",
        "说明：表头行号，填数字；若填 3，则第 3 行作为表头",
        "说明：数据起始行号，填数字；若填 4，则从第 4 行开始读取数据",
        "说明：贷款投向行业列序号，对应源数据中“贷款实际投向行业（行业小类）”那一列，单元格内容需包含类似 A1234 的行业码，程序用正则从整格文本中提取；若列号填错或该列没有标准行业码，则无法在参照表中匹配到行业，整行不会产生多报/漏报结果",
        "说明：机构报送产业分类列序号，可填多个列号，用英文逗号分隔，例如 17,20,23；若写少了会漏判该列报送，写错列号会把无关列当作报送代码导致误判",
        "说明：报送列与产业类别代码映射，例如 17:HTP,20:HTS,23:SE,26:PA；若类别映射错位（如 17 映射到 HTS），则该列报送会被归入错误产业，导致该产业多报/漏报统计全部偏移",
        "说明：参照表工作表序号，可填数字序号（从1开始）或参照表 sheet 名称（需与参照表内名称完全一致）",
        "说明：参照表中“产业分类代码”所在列号，例如 HTP01/DE02 等",
        "说明：参照表中“行业4位码”所在列号，例如 C2345，将与贷款投向行业提取出的 4 位码匹配",
        "说明：参照表中“是否标*”所在列号，用于识别星标并判断疑似多报/疑似漏报",
        "说明：参照表中“原始映射内容”所在列号（默认 U 列=21），一般为小类/中类/大类/门类说明，用于结果中【code:原始映射】展示",
    ]
    if ws_cfg.max_row < 4:
        while ws_cfg.max_row < 4:
            ws_cfg.append([])
        ws_cfg.append(desc_row)
    else:
        ws_cfg.insert_rows(5)
        for idx, val in enumerate(desc_row, start=1):
            ws_cfg.cell(5, idx).value = val
    ws_m = wb.create_sheet("mapping")
    if mapping_rows:
        for row in mapping_rows:
            ws_m.append(row)
    else:
        ws_m.append(["参照表路径", ""])
        ws_m.append(["映射表路径", ""])
    # log：始终创建并写入完整表头，有历史则保留数据行
    LOG_HEADERS = [
        "运行时间", "参照表路径", "参照表sheet数量", "参照表sheet列表",
        "结果文件", "明细sheet名称", "来源台账文件", "台账类型",
        "使用参照表sheet序号", "使用参照表sheet名称",
        "错报数量", "漏报数量", "疑似漏报数量",
    ]
    ws_l = wb.create_sheet("log")
    ws_l.append(LOG_HEADERS)
    for c in range(1, len(LOG_HEADERS) + 1):
        ws_l.cell(1, c).font = Font(name="Arial", bold=True)
    if log_rows:
        has_header = log_rows[0] and str(log_rows[0][0]).strip() == "运行时间"
        data_start = 1 if has_header else 0
        for row in log_rows[data_start:]:
            ws_l.append(row)
    ws_clue = wb.create_sheet("clue")
    if clue_rows and len(clue_rows) >= 1:
        for row in clue_rows:
            ws_clue.append(row)
        for c in range(1, len(CLUE_HEADERS) + 1):
            ws_clue.cell(1, c).font = Font(name="Arial", bold=True)
    else:
        ws_clue.append(CLUE_HEADERS)
        for c in range(1, len(CLUE_HEADERS) + 1):
            ws_clue.cell(1, c).font = Font(name="Arial", bold=True)
        for row in CLUE_DEFAULT_ROWS:
            ws_clue.append(row)
    wb.save(cfg_path)
    wb.close()
    print(f"config 工作表已重建，共 {len(rows)} 行配置。")


def main():
    while True:
        print("\n====== 策略一工具面板 ======")
        print("1. 执行策略一（读取config.xlsx和参照表）")
        print("2. 配置参照表路径（mapping面板）")
        print("3. 批量配置源文件路径（mapping面板）")
        print("4. 配置国民经济行业分类映射表路径（mapping面板）")
        print("5. 重置config.xlsx的config工作表")
        print("6. 执行策略一（备用，固定规则不读 clue 表）")
        print("0. 退出")
        choice = input("请输入序号并回车：").strip()
        if choice == "1":
            mapping, out, cfg = write_results()
            print("策略一执行完成。")
            print("参照表：", mapping)
            print("结果文件：", out)
        elif choice == "6":
            mapping, out, cfg = write_results()
            print("策略一（备用）执行完成。")
            print("参照表：", mapping)
            print("结果文件：", out)
        elif choice == "2":
            cli_set_mapping_path()
        elif choice == "3":
            cli_set_source_files()
        elif choice == "4":
            cli_set_industry_map_path()
        elif choice == "5":
            reset_config_sheet()
        elif choice in ("0", ""):
            print("已退出。")
            break
        else:
            print("无效输入，请重新选择。")


if __name__ == "__main__":
    main()
