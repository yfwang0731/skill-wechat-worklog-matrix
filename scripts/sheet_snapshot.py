# -*- coding: utf-8 -*-
"""云文档（WPS/金山）读表快照：把 kdocs 的稀疏返回重建成脚本认识的网格。

为什么需要这一层：连接器的 MCP 工具只有 agent 能调，Python 脚本调不到。
所以读表拆成「agent 取数 → 落 raw.json → 本脚本重建 → 后续脚本读快照」。

读表建议分两趟，避免把整张表（可能上千行）全量搬进上下文：
  第 1 趟（轻）：只读前 3 行 × 全部列 → raw_hdr.json
                用于识别表头名 → 列映射（跑 probe.py workbook --snapshot）
  第 2 趟（窄）：只读映射到的关键列（需求描述/提出时间/提出人/提出人岗位/解决人…）
                整表行 → raw_cols.json
                用于算「末数据行 / 追加起始行」与岗位复用

用法：
  python sheet_snapshot.py plan --file-id <id> --worksheet-id <n> [--rows N] [--cols N] [--letters C,N,O]
  python sheet_snapshot.py build --raw raw_hdr.json [--raw raw_cols.json ...]
                                --out sheet_snapshot.json [--sheet "运维-2026"]
                                [--worksheet-id N] [--file-id X] [--drive-id Y] [--name 文档名]
  python sheet_snapshot.py inspect --snapshot sheet_snapshot.json

快照结构（后续脚本只认 rows 这个密集二维数组）：
{
  "source": "kdocs",
  "doc": {"file_id": "...", "drive_id": "...", "name": "..."},
  "fetched_at": "2026-09-11T10:00:00+08:00",
  "sheets": [
    {"name": "运维-2026", "worksheet_id": 3,
     "num_formats": {"3,12": "yyyy/m/d"},
     "rows": [["项目编号", "项目名称", ...], ...]}
  ]
}
> 接口要点：kdocs 连接器的工作表参数名是 **worksheet_id**（不是 sheetId），读写都用它。
raw.json 可为以下任一形态（本脚本自动识别）：
  {"result":"ok","detail":{"rangeData":[...]}} / {"rangeData":[...]} / [...] / MCP content 包装
"""
import os
import sys
import json
import argparse
from datetime import datetime, timezone, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from common import merge_range_data, store_to_rows, col_index, workbook_from_snapshot

TZ = timezone(timedelta(hours=8))

# 第 2 趟建议读的关键列（逻辑列名；字母由第 1 趟的列映射给出）
KEY_COLUMNS = ["需求描述", "提出时间", "提出人", "提出人岗位", "解决人", "责任人"]


# ---------------- raw 解析 ----------------
def extract_range_data(obj):
    """从各种可能的 kdocs/MCP 返回形态里取出 rangeData 数组（取不到返回 []）。"""
    if obj is None:
        return []
    if isinstance(obj, list):
        return obj
    if not isinstance(obj, dict):
        return []
    # MCP content 包装：{"content":[{"type":"text","text":"<json>"}]}
    content = obj.get("content")
    if isinstance(content, list):
        for it in content:
            if isinstance(it, dict) and isinstance(it.get("text"), str):
                try:
                    got = extract_range_data(json.loads(it["text"]))
                except Exception:
                    got = []
                if got:
                    return got
    for key in ("rangeData",):
        if isinstance(obj.get(key), list):
            return obj[key]
    for wrap in ("detail", "data", "result"):
        inner = obj.get(wrap)
        if isinstance(inner, dict):
            if isinstance(inner.get("rangeData"), list):
                return inner["rangeData"]
            if isinstance(inner.get("detail"), dict) and \
                    isinstance(inner["detail"].get("rangeData"), list):
                return inner["detail"]["rangeData"]
    return []


def load_raw(path):
    """读一个 raw 文件，返回 (rangeData, 透传的 sheets 或 None)。"""
    with open(path, encoding="utf-8") as f:
        obj = json.load(f)
    sheets = obj.get("sheets") if isinstance(obj, dict) else None
    return extract_range_data(obj), sheets


def pick_meta(obj):
    """从 raw 里顺带捞文档元信息（file_id / drive_id / worksheet_id），没有就返回 {}。"""
    KEYS = ("file_id", "drive_id", "worksheet_id", "sheetId", "sheet_id", "name")
    if not isinstance(obj, dict):
        return {}
    out = {}
    for k in KEYS:
        if obj.get(k) not in (None, ""):
            out[k] = obj[k]
    for wrap in ("detail", "data"):
        inner = obj.get(wrap)
        if isinstance(inner, dict):
            for k in KEYS:
                if inner.get(k) not in (None, "") and k not in out:
                    out[k] = inner[k]
    return out


# ---------------- 子命令 ----------------
def cmd_build(args):
    store = {}
    passthrough = []
    meta = {}
    total_cells = 0
    for p in args.raw:
        if not os.path.isfile(p):
            sys.exit(f"✗ raw 文件不存在：{p}")
        try:
            with open(p, encoding="utf-8") as f:
                obj = json.load(f)
        except json.JSONDecodeError as e:
            sys.exit(f"✗ {p} 不是合法 JSON（{e}）。"
                     "请把 kdocs 读表返回**原样**存盘，不要手工改结构。")
        rd, sheets = load_raw(p)
        n = len(rd or [])
        total_cells += n
        merge_range_data(store, rd)
        if sheets:
            passthrough.extend(sheets)
        for k, v in pick_meta(obj).items():
            meta.setdefault(k, v)
        print(f"  [raw] {os.path.basename(p)}: {n} 个非空单元格")

    sheets_out = list(passthrough)
    if store.get("cells"):
        rows = store_to_rows(store)
        ws_id = args.worksheet_id
        if ws_id is None:
            ws_id = meta.get("worksheet_id", meta.get("sheetId", meta.get("sheet_id")))
        sheets_out.append({
            "name": args.sheet or args.name or "sheet1",
            "worksheet_id": ws_id,
            "num_formats": store.get("num_formats") or {},
            "rows": rows,
        })

    if not sheets_out:
        sys.exit("✗ 所有 raw 文件都没有解析出 rangeData（也没有 sheets）。"
                 "请确认存的是 kdocs sheet.get_range_data 的返回。")

    snap = {
        "source": "kdocs",
        "doc": {
            "file_id": args.file_id or meta.get("file_id"),
            "drive_id": args.drive_id or meta.get("drive_id"),
            "name": args.name or meta.get("name"),
        },
        "fetched_at": datetime.now(TZ).strftime("%Y-%m-%dT%H:%M:%S%z"),
        "sheets": sheets_out,
    }
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(snap, f, ensure_ascii=False, indent=2)

    wb = workbook_from_snapshot(snap)
    print(f"\n[snapshot] {args.out}")
    print(f"  文档：{snap['doc'].get('name') or '(未记录)'}  file_id={snap['doc'].get('file_id')}")
    print(f"  工作表 {len(wb.sheetnames)} 张：" +
          "、".join(f"{s.title}({s.max_row}行×{s.max_column}列)"
                   for s in [wb[n] for n in wb.sheetnames]))
    print(f"  共并入 {total_cells} 个单元格")
    print("  下一步：python scripts/probe.py workbook --snapshot "
          f"{os.path.basename(args.out)} --json")


def cmd_inspect(args):
    snap = json.load(open(args.snapshot, encoding="utf-8"))
    wb = workbook_from_snapshot(snap)
    print(f"[snapshot] {args.snapshot}")
    print(f"  source={snap.get('source')}  拉取时间={snap.get('fetched_at')}")
    doc = snap.get("doc") or {}
    print(f"  文档：{doc.get('name') or '(未记录)'}  file_id={doc.get('file_id')}")
    for name in wb.sheetnames:
        s = wb[name]
        print(f"\n  ── 子表「{name}」 worksheet_id={s.sheet_id}  "
              f"{s.max_row} 行 × {s.max_column} 列")
        for r in range(1, min(4, s.max_row) + 1):
            vals = [v for v in s.row_values(r) if v != ""]
            if vals:
                print(f"     第{r}行（{len(vals)} 个非空）: " + " | ".join(vals[:12]) +
                      (" …" if len(vals) > 12 else ""))
        # 末数据行（以行内任一非空为准）
        last = 0
        for r in range(s.max_row, 0, -1):
            if any(v != "" for v in s.row_values(r)):
                last = r
                break
        print(f"     末非空行 = {last} → 追加起始行 {last + 1}")
        if s.num_formats:
            sample = list(s.num_formats.items())[:3]
            print(f"     已记录 {len(s.num_formats)} 个单元格数字格式，如 "
                  + "、".join(f"({k})={v}" for k, v in sample))


def cmd_plan(args):
    """打印建议的 kdocs 读表调用体（agent 填 file_id / worksheet_id 后直接调用）。"""
    cols = args.cols or 30
    rows = args.rows
    plan = {
        "step1_headers": {
            "why": "只读前 3 行识别表头名 → 列映射；这样第 2 趟才能只读关键列",
            "tool": "sheet.get_range_data",
            "take": "worksheet_id 未定时先调 sheet.get_sheets_info 取 range.rowTo / 列数",
            "body": {
                "file_id": args.file_id or "<file_id>",
                "worksheet_id": args.worksheet_id if args.worksheet_id is not None else "<worksheet_id>",
                "range": {"rowFrom": 0, "rowTo": 2, "colFrom": 0, "colTo": cols - 1},
            },
            "save_raw_as": "raw_hdr.json",
            "then": "python scripts/sheet_snapshot.py build --raw raw_hdr.json "
                    "--out sheet_snapshot.json  →  python scripts/probe.py workbook "
                    "--snapshot sheet_snapshot.json --json",
        },
        "step2_key_columns": {
            "why": "只读映射到的关键列，算末数据行 / 追加起始行 / 岗位复用；避免整表搬进上下文",
            "tool": "sheet.get_range_data（每段列各调一次，或按最小-最大列一次读完）",
            "letters": args.letters or "<第1趟得到的列字母，如 C,N,O>",
            "body": {
                "file_id": args.file_id or "<file_id>",
                "worksheet_id": args.worksheet_id if args.worksheet_id is not None else "<worksheet_id>",
                "range": {"rowFrom": 0,
                          "rowTo": (rows - 1) if rows else "<取 get_sheets_info 的 range.rowTo>",
                          "colFrom": "<最小列 0-based>", "colTo": "<最大列 0-based>"},
            },
            "save_raw_as": "raw_cols.json",
            "then": "python scripts/sheet_snapshot.py build --raw raw_hdr.json --raw raw_cols.json "
                    "--out sheet_snapshot.json  →  再跑 probe.py workbook --snapshot",
        },
        "suggested_key_columns": KEY_COLUMNS,
        "note": "raw 存盘务必原样（含 rangeData 的完整结构），build 会自动识别多种包装格式。"
                "工作表参数名统一用 worksheet_id。",
    }
    if args.letters:
        try:
            idx = sorted(col_index(x.strip()) for x in args.letters.split(",") if x.strip())
            plan["step2_key_columns"]["body"]["range"]["colFrom"] = idx[0]
            plan["step2_key_columns"]["body"]["range"]["colTo"] = idx[-1]
        except Exception:
            pass
    print(json.dumps(plan, ensure_ascii=False, indent=2))


def main():
    ap = argparse.ArgumentParser(description="云文档读表快照构建/检查/读表计划")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p1 = sub.add_parser("build", help="把 kdocs 返回的 raw.json 合并成快照")
    p1.add_argument("--raw", action="append", required=True,
                    help="kdocs sheet.get_range_data 的返回文件（可多次 --raw 合并）")
    p1.add_argument("--out", required=True)
    p1.add_argument("--sheet", default=None, help="快照里子表的名字（也是后续 sheet_match 匹配用）")
    p1.add_argument("--worksheet-id", "--sheet-id", dest="worksheet_id", default=None)
    p1.add_argument("--file-id", default=None)
    p1.add_argument("--drive-id", default=None)
    p1.add_argument("--name", default=None, help="文档名（仅记录，便于人眼核对）")

    p2 = sub.add_parser("inspect", help="查看快照内容概要")
    p2.add_argument("--snapshot", required=True)

    p3 = sub.add_parser("plan", help="打印建议的 kdocs 读表调用体")
    p3.add_argument("--file-id", default=None)
    p3.add_argument("--worksheet-id", "--sheet-id", dest="worksheet_id", default=None)
    p3.add_argument("--rows", type=int, default=None, help="表的大致总行数")
    p3.add_argument("--cols", type=int, default=None, help="表的大致总列数")
    p3.add_argument("--letters", default=None, help="第 2 趟要读的列字母，逗号分隔，如 C,N,O")

    a = ap.parse_args()
    if a.cmd == "build":
        cmd_build(a)
    elif a.cmd == "inspect":
        cmd_inspect(a)
    else:
        cmd_plan(a)


if __name__ == "__main__":
    main()
