# -*- coding: utf-8 -*-
"""微信→需求跟踪矩阵 脚本共用模块：配置加载、路径自动识别、日期/列工具。

设计原则：
1. 不写死任何具体项目/人员/客户 —— 一切从 config.json 读取。
2. 本机路径一律自动识别（或由探测结果写入 config），不写死绝对路径。
3. 列位置靠"表头名"识别（probe.py 生成 column_mapping），不写死 A~AA。
"""
import os
import re
import json
import hashlib
from datetime import date

SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# ---------------- 配置 ----------------
def load_config(path=None):
    """加载 config.json。查找顺序：显式路径 → 环境变量 WM_CONFIG → 当前目录 → skill 根目录。"""
    cands = [path,
             os.environ.get("WM_CONFIG"),
             os.path.join(os.getcwd(), "config.json"),
             os.path.join(SKILL_ROOT, "config.json")]
    for p in cands:
        if p and os.path.isfile(p):
            with open(p, encoding="utf-8") as f:
                return json.load(f), p
    return None, None


def col_index(letter):
    """列字母 -> 0-based 索引（A->0, AA->26）"""
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n - 1


def col_letter(idx0):
    """0-based 索引 -> 列字母"""
    s, n = "", idx0 + 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ---------------- 路径自动识别 ----------------
def find_db_storage(preferred=None, hint=None):
    """自动定位微信 db_storage。多账号时可用 hint（账号目录特征串）优选。"""
    base = os.path.join(os.path.expanduser("~"), "Documents", "xwechat_files")
    if preferred:
        if os.path.isdir(preferred):
            return preferred
        cand = os.path.join(base, preferred, "db_storage")
        if os.path.isdir(cand):
            return cand
    if os.path.isdir(base):
        subs = [d for d in os.listdir(base)
                if os.path.isdir(os.path.join(base, d, "db_storage"))]
        if hint:
            for s in subs:
                if hint in s:
                    return os.path.join(base, s, "db_storage")
        if subs:
            return os.path.join(base, subs[0], "db_storage")
    raise FileNotFoundError(
        f"未在 {base} 找到微信 db_storage；请通过 --db-storage 指定或先跑 probe.py accounts。")


def find_wcdb_tool(explicit=None):
    """定位 wcdb-key-tool 的 windows 主脚本（第三方工具，不在 skill 内）。

    顺序：显式路径 → 环境变量 WCDB_KEY_TOOL → skill 内已知位置 → 家目录深度≤3 搜索。
    """
    if explicit and os.path.isfile(explicit):
        return explicit
    env = os.environ.get("WCDB_KEY_TOOL")
    if env and os.path.isfile(env):
        return env
    here = os.path.dirname(os.path.abspath(__file__))
    for c in [os.path.join(here, "tools", "wcdb-key-tool", "wcdb_key_tool_windows.py"),
              os.path.join(here, "wcdb-key-tool-main", "wcdb_key_tool_windows.py")]:
        if os.path.isfile(c):
            return os.path.abspath(c)
    home, target = os.path.expanduser("~"), "wcdb_key_tool_windows.py"
    for root, dirs, files in os.walk(home):
        if root[len(home):].count(os.sep) > 3:
            dirs[:] = []
            continue
        if target in files:
            return os.path.join(root, target)
    return None


# ---------------- 工具 ----------------
def md5hex(s):
    return hashlib.md5(s.encode("utf-8")).hexdigest()


def serial(dstr):
    """日期 -> Excel 序列号（1899-12-30 起）。容错支持 YYYY-MM-DD / YYYY/MM/DD /
    YYYY年M月D日 / 带时间（取前 10 位）等；无法解析返回 ''。"""
    if not dstr:
        return ""
    s = str(dstr).strip()
    m = re.match(r"^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})", s)
    if not m:
        return ""
    try:
        y, mo, dd = map(int, m.groups())
        return (date(y, mo, dd) - date(1899, 12, 30)).days
    except Exception:
        return ""


def pick_sheet(workbook, sheet_hint=None):
    """按名字含 sheet_hint 选子表，找不到退回第一张。workbook 可为 openpyxl 或 GridWorkbook。"""
    names = list(workbook.sheetnames)
    hint = sheet_hint or ""
    cand = [n for n in names if hint and hint in n]
    return workbook[cand[0] if cand else names[0]]


def next_append_row_ws(ws, mapping=None):
    """返回工作表下一个可追加行号（1-based）。

    mapping 为 column_mapping（逻辑列名->列字母）时，以"需求描述/提出时间/提出人"列
    有无值判定数据行；否则扫描**全部已有列**（不再只扫前 27 列，宽表也能判准）。
    """
    probe_cols = []
    if mapping:
        for key in ("需求描述", "提出时间", "提出人"):
            l = mapping.get(key)
            if l:
                probe_cols.append(col_index(l) + 1)
    if not probe_cols:
        probe_cols = list(range(1, max(1, ws.max_column or 0) + 1))
    last = 1
    for r in range(1, ws.max_row + 1):
        if any(ws.cell(row=r, column=c).value not in (None, "") for c in probe_cols):
            last = r
    return last + 1


def next_append_row(workbook_path, sheet_hint="运维", mapping=None):
    """本地 xlsx 版本的追加起始行（依赖 openpyxl）。云文档通道请用 next_append_row_ws + 快照。"""
    import openpyxl
    wb = openpyxl.load_workbook(workbook_path)
    return next_append_row_ws(pick_sheet(wb, sheet_hint), mapping)


# ---------------- 表格快照（云文档 / WPS 通道）----------------
# 脚本无法直连 MCP 连接器，故两端都收敛到 JSON：
#   读：agent 调 kdocs sheet.get_range_data → 原始返回存 raw.json → sparse_to_grid 成快照
#   写：脚本产出 payload → to_kdocs_payload.py 转 rangeData → agent 调 sheet.update_range_data
# 快照结构见 sheet_snapshot.py 文档字符串。

class _Cell:
    """冒充 openpyxl 的 Cell（本 skill 只用到 .value）。"""
    __slots__ = ("value",)

    def __init__(self, value):
        self.value = value


class GridSheet:
    """用二维字符串数组冒充工作表，接口与 openpyxl 用到的部分保持一致（行列 1-based）。"""

    def __init__(self, name, rows, num_formats=None, sheet_id=None):
        self.title = name
        self.sheet_id = sheet_id
        self._rows = [[("" if v is None else str(v)) for v in r] for r in (rows or [])]
        self.num_formats = num_formats or {}

    @property
    def max_row(self):
        return len(self._rows)

    @property
    def max_column(self):
        return max((len(r) for r in self._rows), default=0)

    def cell(self, row=1, column=1):
        r, c = row - 1, column - 1
        v = ""
        if 0 <= r < len(self._rows) and 0 <= c < len(self._rows[r]):
            v = self._rows[r][c]
        return _Cell(v if v != "" else None)

    def row_values(self, row=1):
        r = row - 1
        return list(self._rows[r]) if 0 <= r < len(self._rows) else []


class GridWorkbook:
    """冒充 openpyxl Workbook（只实现 .sheetnames / __getitem__）。"""

    def __init__(self, sheets):
        self._sheets = list(sheets)
        self.sheetnames = [s.title for s in self._sheets]

    def __getitem__(self, name):
        for s in self._sheets:
            if s.title == name:
                return s
        raise KeyError(name)


def _entry_value(entry):
    """从 kdocs 单元格条目取"人看的文本"：cellText 优先（公式则退回原始值）。"""
    txt = entry.get("cellText")
    orig = entry.get("originalCellValue")
    if txt not in (None, "") and not str(txt).startswith("="):
        return str(txt)
    if orig not in (None, ""):
        return str(orig)
    return ""


def merge_range_data(store, range_data):
    """把一个 kdocs rangeData（稀疏）并入累加器 store（原地修改）。

    store = {"cells": {(r,c): str}, "num_formats": {"r,c": fmt}, "max_row": int, "max_col": int}
    """
    store.setdefault("cells", {})
    store.setdefault("num_formats", {})
    for e in range_data or []:
        if not isinstance(e, dict):
            continue
        try:
            rf = int(e.get("rowFrom", 0) or 0)
            cf = int(e.get("colFrom", 0) or 0)
            rt = int(e.get("rowTo", rf) if e.get("rowTo") is not None else rf)
            ct = int(e.get("colTo", cf) if e.get("colTo") is not None else cf)
        except (TypeError, ValueError):
            continue
        v = _entry_value(e)
        nf = e.get("numFormat")
        for r in range(min(rf, rt), max(rf, rt) + 1):
            for c in range(min(cf, ct), max(cf, ct) + 1):
                store["cells"][(r, c)] = v
                if nf:
                    store["num_formats"][f"{r},{c}"] = nf
                if r > store.get("max_row", -1):
                    store["max_row"] = r
                if c > store.get("max_col", -1):
                    store["max_col"] = c
    return store


def store_to_rows(store):
    """累加器 -> 紧凑二维列表（裁掉尾部/右侧全空行列）。"""
    cells = store.get("cells") or {}
    maxr, maxc = store.get("max_row", -1), store.get("max_col", -1)
    rows = [[cells.get((r, c), "") for c in range(maxc + 1)] for r in range(maxr + 1)]
    while rows and not any(str(x).strip() for x in rows[-1]):
        rows.pop()
    width = max((len(r) for r in rows), default=0)
    while width and all(not str(r[width - 1]).strip() for r in rows):
        width -= 1
    return [r[:width] for r in rows]


def sparse_to_grid(range_data):
    """kdocs rangeData（稀疏）-> (rows, num_formats)。"""
    store = merge_range_data({}, range_data)
    return store_to_rows(store), store.get("num_formats") or {}


def load_snapshot(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def workbook_from_snapshot(snap):
    """快照 dict -> GridWorkbook。"""
    sheets = []
    for i, s in enumerate((snap or {}).get("sheets") or [], 1):
        rows = s.get("rows")
        if rows is None:
            rows, fmts = sparse_to_grid(s.get("rangeData"))
            s.setdefault("num_formats", fmts)
        sheets.append(GridSheet(s.get("name") or f"sheet{i}", rows,
                                s.get("num_formats"),
                                s.get("worksheet_id", s.get("sheetId"))))
    if not sheets:
        raise ValueError("快照里没有任何工作表（sheets 为空）：请先按 SKILL.md 的"
                         "「WPS 通道读表」把 kdocs 返回存成 raw.json 再 build。")
    return GridWorkbook(sheets)


def next_append_row_snapshot(snapshot_path, sheet_hint="运维", mapping=None):
    """快照版本的追加起始行（不依赖 openpyxl）。"""
    wb = workbook_from_snapshot(load_snapshot(snapshot_path))
    return next_append_row_ws(pick_sheet(wb, sheet_hint), mapping)


# ---------------- 岗位列复用（final 与 position_reuse 共用这一套）----------------
# 设计要点：读历史、填新行、行序传播三件事各只有一份实现，
# 避免"final 内置"与"position_reuse 独立"两条路各写一套而逐渐走偏。

def position_columns(mapping):
    """从 column_mapping 解析 (提出人列, 岗位列)（1-based）。

    缺映射时退回常见列位（提出人=O=15, 岗位=N=14）。返回 (col_person, col_post)。
    """
    mp = mapping or {}
    col_person = col_index(mp["提出人"]) + 1 if mp.get("提出人") else 15
    col_post = col_index(mp["提出人岗位"]) + 1 if mp.get("提出人岗位") else 14
    return col_person, col_post


def read_history_positions(sheet, header_row, end_row, col_person, col_post):
    """读历史区的「提出人 -> 该人最近的非空岗位」。

    **只读**：历史区永不修改，只作为新行的岗位来源。
    区间为 [header_row+1, end_row]（1-based，含两端）；end_row 一般传「追加起始行-1」。

    返回 (last_post, pairs)；pairs 是读到「人+岗位」成对的行数——为 0 时说明
    表格来源没覆盖到这两列的数据区（快照常见），调用方应据此告警而不是静默补 0 条。
    """
    last_post = {}
    pairs = 0
    for r in range(int(header_row) + 1, int(end_row) + 1):
        person = sheet.cell(row=r, column=col_person).value
        if person is None or not str(person).strip():
            continue
        person = str(person).strip()
        post = sheet.cell(row=r, column=col_post).value
        if post is not None and str(post).strip():
            last_post[person] = str(post).strip()
            pairs += 1
    return last_post, pairs


def reuse_position_fill(new_rows, last_post, person_key="提出人", post_key="提出人岗位",
                        override=False):
    """把 last_post（人 -> 最近岗位）复用到 new_rows 的空岗位上，**原地修改** new_rows。

    返回变更明细 [{"i": 行下标, "person": 人, "old": 原值, "new": 补的值}]。

    语义（与早期 position_reuse 完全一致，只是挪到共用函数里）：
      - 默认**只填空缺**；override=True 才允许把已有值改写成 last_post 里的值。
      - 按行序推进：本批先出现的非空岗位会写回 last_post，供本批后续行使用，
        因此"历史 + 本批"合起来就是整表从上到下的自然语义。
      - last_post 会被原地更新（调用方若还要用需自己拷贝）。
    """
    changes = []
    for i, r in enumerate(new_rows or []):
        person = str(r.get(person_key) or "").strip()
        if not person:
            continue
        cur = str(r.get(post_key) or "").strip()
        hist = str(last_post.get(person) or "").strip()
        if hist and ((not cur) or (override and cur != hist)):
            changes.append({"i": i, "person": person, "old": cur, "new": hist})
            r[post_key] = hist
            cur = hist
        if cur:
            last_post[person] = cur
    return changes


if __name__ == "__main__":
    cfg, path = load_config()
    print("config:", path or "(未找到 config.json，请先从 config.example.json 复制)")
    print("db_storage:", find_db_storage())
    print("wcdb_tool :", find_wcdb_tool() or "(未找到，请 clone 或设 WCDB_KEY_TOOL)")
