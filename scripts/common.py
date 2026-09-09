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


def next_append_row(workbook_path, sheet_hint="运维", mapping=None):
    """返回子表（名含 sheet_hint）下一个可追加行号（1-based）。

    mapping 为 column_mapping（逻辑列名->列字母）时，以"需求描述/提出时间/提出人"列
    有无值判定数据行；否则退化为扫描前 27 列。
    """
    import openpyxl
    wb = openpyxl.load_workbook(workbook_path)
    names = [s for s in wb.sheetnames if sheet_hint in s]
    ws = wb[names[0] if names else wb.sheetnames[0]]
    probe_cols = []
    if mapping:
        for key in ("需求描述", "提出时间", "提出人"):
            l = mapping.get(key)
            if l:
                probe_cols.append(col_index(l) + 1)
    if not probe_cols:
        probe_cols = list(range(1, 28))
    last = 1
    for r in range(1, ws.max_row + 1):
        if any(ws.cell(row=r, column=c).value not in (None, "") for c in probe_cols):
            last = r
    return last + 1


if __name__ == "__main__":
    cfg, path = load_config()
    print("config:", path or "(未找到 config.json，请先从 config.example.json 复制)")
    print("db_storage:", find_db_storage())
    print("wcdb_tool :", find_wcdb_tool() or "(未找到，请 clone 或设 WCDB_KEY_TOOL)")
