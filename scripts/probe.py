# -*- coding: utf-8 -*-
"""前置探测：把"账户 / 会话 / Excel 表头列映射"扫出来，供一次性确认与生成 config。

设计目标：列映射靠**表头名**识别，不再写死 A~AA；账户/会话/时间范围都由用户选择。
三个子命令：
  python probe.py accounts                                   # 列出本机微信账户
  python probe.py sessions --decrypted <dir> [--since] [--keyword]   # 列出可选会话
  python probe.py workbook --workbook <xlsx> [--sheet]       # 解析表头 → 列映射 + 末行 + 处理人候选

输出统一 JSON 到 stdout（便于 agent 读取），同时打印人类可读摘要。
"""
import os
import sys
import json
import sqlite3
import argparse
from datetime import datetime, timezone, timedelta

from common import md5hex, col_index, col_letter

TZ = timezone(timedelta(hours=8))

# 逻辑列 -> 可能的表头名（按优先级）。用于"通过表头读列"，跨模板通用。
COLUMN_ALIASES = {
    "项目编号":   ["项目编号", "项目编码", "项目号"],
    "项目名称":   ["项目名称", "项目"],
    "项目负责人": ["项目负责人", "负责人", "项目经理"],
    "需求描述":   ["用户需求内容描述", "需求描述", "需求内容", "用户需求", "描述", "内容"],
    "提出时间":   ["提出时间", "提出日期", "提出需求的日期", "需求提出时间"],
    "提出人岗位": ["提出需求的岗位", "提出人岗位", "岗位", "提出岗位"],
    "提出人":     ["提出人", "提出人姓名", "需求提出人", "提出者"],
    "产生阶段":   ["产生阶段", "阶段", "需求产生阶段"],
    "需求归类":   ["需求归类", "归类", "分类", "需求分类"],
    "影响级别":   ["影响级别", "影响程度", "级别", "影响"],
    "优先级":     ["优先级", "优先级别", "优先度"],
    "状态":       ["状态", "处理状态", "需求状态"],
    "责任人":     ["责任人", "负责人"],
    "计划时间":   ["计划时间", "计划完成时间", "计划日期"],
    "完成时间":   ["需求完成时间", "完成时间", "实际完成时间", "完成日期"],
    "解决人":     ["解决人", "处理人", "解决人员"],
    "备注":       ["备注", "说明", "补充说明"],
    "是否收费":   ["是否收费", "收费"],
    "需求类型":   ["需求类型", "类型"],
    "子系统名称": ["子系统名称", "子系统", "系统名称"],
}

# 需要按日期格式写入的列（匹配到逻辑列后据此设置 number_format）
DATE_FIELDS = {"提出时间", "计划时间", "完成时间"}


# ---------------- accounts ----------------
def cmd_accounts(args):
    base = os.path.join(os.path.expanduser("~"), "Documents", "xwechat_files")
    out = []
    if os.path.isdir(base):
        for d in sorted(os.listdir(base)):
            ds = os.path.join(base, d, "db_storage")
            if not os.path.isdir(ds):
                continue
            try:
                mtime = datetime.fromtimestamp(os.path.getmtime(ds), TZ).strftime("%Y-%m-%d %H:%M")
            except Exception:
                mtime = ""
            hint_hit = bool(args.hint) and args.hint in d
            out.append({"account_dir": d, "db_storage": ds, "last_modified": mtime,
                        "hint_match": hint_hit})
    print(json.dumps({"accounts": out, "base": base}, ensure_ascii=False, indent=2))
    if getattr(args, "json", False):
        return
    print(f"\n[accounts] 共 {len(out)} 个微信账户目录：")
    for i, a in enumerate(out, 1):
        flag = f"  ← 匹配 --hint「{args.hint}」" if a["hint_match"] else ""
        print(f"  {i}. {a['account_dir']}  (更新于 {a['last_modified']}){flag}")
    print("  请让用户选择使用哪个账户（skill 不内置任何账号特征）。")


# ---------------- sessions ----------------
def list_session_tables(decrypted):
    """返回 [(wxid, 显示名或None)]，会话表 Msg_<md5> 反查需要 contact 表。"""
    contact_db = os.path.join(decrypted, "contact", "contact.db")
    if not os.path.isfile(contact_db):
        raise FileNotFoundError(
            f"未找到解密库 contact 表：{contact_db}\n"
            "请先确认 config.account.decrypted 指向正确的解密输出目录，"
            "再执行 `python pipeline.py run`（或 decrypt.py）完成解密。")
    con = sqlite3.connect(contact_db)
    cur = con.cursor()
    # 显示名：remark 优先，其次 nick_name
    name_by_user = {}
    for username, remark, nick in cur.execute("SELECT username, remark, nick_name FROM contact"):
        if username:
            name_by_user[username] = (remark or "").strip() or (nick or "").strip() or username
    con.close()
    return name_by_user


def cmd_sessions(args):
    decrypted = args.decrypted or os.path.join(os.getcwd(), "wechat_pilot", "output", "decrypted")
    since_epoch = 0
    if args.since:
        since_epoch = int(datetime.strptime(args.since, "%Y-%m-%d")
                          .replace(tzinfo=TZ).timestamp())
    shards = ["message_2.db", "message_1.db", "message_3.db"]
    try:
        name_by_user = list_session_tables(decrypted)
    except FileNotFoundError as e:
        sys.exit(f"✗ 无法列出会话：{e}")

    # 收集所有会话表
    tables = set()
    for sh in shards:
        p = os.path.join(decrypted, "message", sh)
        if not os.path.exists(p):
            continue
        cc = sqlite3.connect(p)
        for (t,) in cc.execute("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'Msg_%'"):
            tables.add(t)
        cc.close()

    # md5 -> wxid（借助 contact 表反查）
    md5_to_user = {}
    for u in name_by_user:
        md5_to_user.setdefault(md5hex(u), u)

    agg = {}
    for sh in shards:
        p = os.path.join(decrypted, "message", sh)
        if not os.path.exists(p):
            continue
        cc = sqlite3.connect(p)
        for t in list(tables):
            try:
                ex = cc.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (t,)).fetchone()
            except Exception:
                ex = None
            if not ex:
                continue
            try:
                row = cc.execute(
                    f'SELECT COUNT(*), MIN(create_time), MAX(create_time) FROM "{t}" '
                    f'WHERE create_time>=?', (since_epoch,)).fetchone()
            except Exception:
                continue
            cnt, mn, mx = row
            if not cnt:
                continue
            wxid = md5_to_user.get(t[4:], "")
            disp = name_by_user.get(wxid, wxid or t)
            a = agg.setdefault(wxid or t, {"wxid": wxid, "name": disp, "count": 0,
                                           "first": None, "last": None})
            a["count"] += cnt
            if mn and (a["first"] is None or mn < a["first"]):
                a["first"] = mn
            if mx and (a["last"] is None or mx > a["last"]):
                a["last"] = mx
        cc.close()

    def fmt(ts):
        return datetime.fromtimestamp(ts, TZ).strftime("%Y-%m-%d %H:%M") if ts else ""

    sessions = []
    for k, v in agg.items():
        kw = args.keyword
        if kw and kw not in (v["name"] or "") and kw not in (v["wxid"] or ""):
            continue
        sessions.append({
            "wxid": v["wxid"], "name": v["name"],
            "is_group": "@chatroom" in (v["wxid"] or ""),
            "msg_count": v["count"],
            "first": fmt(v["first"]), "last": fmt(v["last"]),
        })
    sessions.sort(key=lambda x: x["last"] or "", reverse=True)

    print(json.dumps({"sessions": sessions, "total": len(sessions)}, ensure_ascii=False, indent=2))
    if getattr(args, "json", False):
        return
    print(f"\n[sessions] 共 {len(sessions)} 个会话（since={args.since or '全部'}，keyword={args.keyword or '无'}）：")
    for i, s in enumerate(sessions, 1):
        g = "[群]" if s["is_group"] else "[人]"
        print(f"  {i:>3}. {g} {s['name'][:28]:<30} 消息{s['msg_count']:>5}  {s['first']} ~ {s['last']}")


# ---------------- workbook ----------------
def cmd_workbook(args):
    import openpyxl
    wb = openpyxl.load_workbook(args.workbook)
    names = wb.sheetnames
    if args.sheet:
        ws = wb[args.sheet] if args.sheet in names else wb[[n for n in names if args.sheet in n][0]]
    else:
        # 优先选数据最多的表（排除"首页"/"说明"之类）
        best, best_n = None, -1
        for n in names:
            w = wb[n]
            cnt = sum(1 for r in range(1, w.max_row + 1)
                      if any(w.cell(row=r, column=c).value not in (None, "")
                             for c in range(1, min(w.max_column, 30) + 1)))
            if cnt > best_n:
                best, best_n = n, cnt
        ws = wb[best] if best else wb[names[0]]

    # 找表头行：取前 5 行中非空单元格最多的一行
    header_row, header_cells = 1, 0
    for r in range(1, min(6, ws.max_row + 1)):
        c = sum(1 for i in range(1, ws.max_column + 1)
                if ws.cell(row=r, column=i).value not in (None, ""))
        if c > header_cells:
            header_row, header_cells = r, c

    headers = {}
    for i in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=i).value
        if v not in (None, ""):
            headers[col_letter(i - 1)] = str(v).strip()

    # 表头名 -> 逻辑列。匹配策略：先精确匹配表头名，再做子串匹配（更长别名优先），
    # 避免短别名（如"项目"⊂"项目负责人"）抢占精确列；一个逻辑列只占用一次。
    mapping, used = {}, set()
    unmatched_headers = []
    # pass1: 精确匹配
    for letter, name in headers.items():
        for logic, aliases in COLUMN_ALIASES.items():
            if logic in used:
                continue
            if name in aliases:
                mapping[logic] = letter
                used.add(logic)
                break
    # pass2: 子串匹配（对仍未匹配的表头，取"命中且最长"的别名所对应逻辑列）
    for letter, name in headers.items():
        if letter in mapping.values():
            continue
        hit, best = None, -1
        for logic, aliases in COLUMN_ALIASES.items():
            if logic in used:
                continue
            for al in aliases:
                if len(al) >= 2 and al in name and len(al) > best:
                    hit, best = logic, len(al)
        if hit:
            mapping[hit] = letter
            used.add(hit)
        else:
            unmatched_headers.append({"col": letter, "header": name})

    missing = [k for k in COLUMN_ALIASES if k not in mapping and k in
               ("需求描述", "提出时间", "提出人", "解决人")]

    # 数据末行（以任意已映射列有值为准）
    last_row = header_row
    mapped_cols = [col_index(l) + 1 for l in mapping.values()]
    for r in range(header_row + 1, ws.max_row + 1):
        if any(ws.cell(row=r, column=c).value not in (None, "") for c in mapped_cols):
            last_row = r

    # 处理人候选：解决人/责任人 列中出现最多的值
    handler_candidates = []
    for logic in ("解决人", "责任人", "项目负责人"):
        l = mapping.get(logic)
        if not l:
            continue
        ci = col_index(l) + 1
        freq = {}
        for r in range(header_row + 1, last_row + 1):
            v = ws.cell(row=r, column=ci).value
            if v not in (None, ""):
                freq[str(v).strip()] = freq.get(str(v).strip(), 0) + 1
        for k, v in sorted(freq.items(), key=lambda x: -x[1])[:3]:
            handler_candidates.append({"field": logic, "value": k, "count": v})
        break

    date_cols = [mapping[k] for k in DATE_FIELDS if k in mapping]

    result = {
        "workbook": args.workbook,
        "sheets": names,
        "sheet_selected": ws.title,
        "header_row": header_row,
        "headers": headers,
        "column_mapping": mapping,
        "unmatched_headers": unmatched_headers,
        "missing_required": missing,
        "last_data_row": last_row,
        "next_append_row": last_row + 1,
        "date_columns": date_cols,
        "handler_candidates": handler_candidates,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if getattr(args, "json", False):
        return
    print(f"\n[workbook] {os.path.basename(args.workbook)} / 子表「{ws.title}」 表头第{header_row}行")
    print("  列映射（按表头识别）：")
    for k, v in mapping.items():
        mark = " [日期列]" if k in DATE_FIELDS else ""
        print(f"    {v:<3} = {k}{mark}")
    if unmatched_headers:
        print(f"  未识别表头 {len(unmatched_headers)} 个（不影响写入）："
              + "、".join(f"{u['col']}:{u['header']}" for u in unmatched_headers[:8]))
    print(f"  末数据行 {last_row} → 追加起始行 {last_row + 1}")
    if handler_candidates:
        print("  处理人候选（自动推断）："
              + "、".join(f"{h['value']}({h['count']}次)" for h in handler_candidates))


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    p1 = sub.add_parser("accounts")
    p1.add_argument("--hint", default=None,
                    help="可选的账号目录特征串，仅用于标记匹配项（不内置任何默认值）")
    p1.add_argument("--json", action="store_true",
                    help="只输出 JSON（供 pipeline/脚本调用，不打印人类摘要）")
    p2 = sub.add_parser("sessions")
    p2.add_argument("--decrypted", default=None)
    p2.add_argument("--since", default=None)
    p2.add_argument("--keyword", default=None)
    p2.add_argument("--json", action="store_true",
                    help="只输出 JSON（供 pipeline/脚本调用，不打印人类摘要）")
    p3 = sub.add_parser("workbook")
    p3.add_argument("--workbook", required=True)
    p3.add_argument("--sheet", default=None)
    p3.add_argument("--json", action="store_true",
                    help="只输出 JSON（供 pipeline/脚本调用，不打印人类摘要）")
    a = ap.parse_args()
    if a.cmd == "accounts":
        cmd_accounts(a)
    elif a.cmd == "sessions":
        cmd_sessions(a)
    else:
        cmd_workbook(a)


if __name__ == "__main__":
    main()
