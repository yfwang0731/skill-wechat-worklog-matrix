# -*- coding: utf-8 -*-
"""从解密后的微信库导出指定会话的转录文本（每行：日期 时间 | 发送者 | 内容）。

**导出哪些会话由 config.json 决定**（scope.sessions 显式清单 或 scope.name_filter 关键字筛选），
不再内置任何具体客户/群。所有路径自动识别或由 config 提供。

用法：
  python export_conversations.py                              # 按 config.json 导出
  python export_conversations.py --since 2026-08-01           # 增量：只导该日期之后
  python export_conversations.py --name-filter <关键字>        # 覆盖关键字筛选
  python export_conversations.py --session <wxid> [--session <wxid2>]   # 只导指定会话
  python export_conversations.py --decrypted <dir> --out <dir>
"""
import sqlite3
import os
import re
import sys
import argparse
from datetime import datetime, timezone, timedelta

from common import md5hex, load_config

TZ = timezone(timedelta(hours=8))
SHARDS = ["message_2.db", "message_1.db", "message_3.db"]   # m2 最老 → m3 最新


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default=None)
    ap.add_argument("--decrypted", default=None)
    ap.add_argument("--out", default=None)
    ap.add_argument("--since", default=None)
    ap.add_argument("--until", default=None)
    ap.add_argument("--name-filter", default=None)
    ap.add_argument("--session", action="append", default=None,
                    help="按 wxid/显示名指定会话，可多次；覆盖 config 的 sessions")
    ap.add_argument("--group", action="append", default=None,
                    help="追加要包含的群 wxid（可选）")
    args = ap.parse_args()

    cfg, _ = load_config(args.config)
    cfg = cfg or {}
    scope = cfg.get("scope", {}) or {}
    acct = cfg.get("account", {}) or {}
    out_cfg = cfg.get("output", {}) or {}

    decrypted = (args.decrypted or acct.get("decrypted")
                 or os.path.join(os.getcwd(), "wechat_pilot", "output", "decrypted"))
    out = args.out or os.path.join(out_cfg.get("dir", "./wechat_pilot"), "transcripts")
    since = args.since or scope.get("since") or "1970-01-01"
    until = args.until or scope.get("until")
    name_filter = args.name_filter or scope.get("name_filter")
    sessions = args.session or scope.get("sessions") or []
    groups = list(scope.get("groups") or []) + (args.group or [])

    since_epoch = int(datetime.strptime(since, "%Y-%m-%d").replace(tzinfo=TZ).timestamp())
    until_epoch = int(datetime.strptime(until, "%Y-%m-%d").replace(tzinfo=TZ).timestamp()) if until else None
    os.makedirs(out, exist_ok=True)

    # 解密库完整性预检（友好报错替代裸 traceback）
    if not os.path.isfile(os.path.join(decrypted, "contact", "contact.db")):
        sys.exit(f"✗ 解密库缺少 contact/contact.db（{decrypted}）。请先执行解密（pipeline.py run / decrypt.py）。")
    msg_dir = os.path.join(decrypted, "message")
    if not os.path.isdir(msg_dir):
        sys.exit(f"✗ 解密库缺少 message 目录（{msg_dir}）。请检查解密输出是否完整。")

    # 联系人显示名
    con = sqlite3.connect(os.path.join(decrypted, "contact", "contact.db"))
    cur = con.cursor()
    disp_all = {}
    for username, remark, nick in cur.execute("SELECT username, remark, nick_name FROM contact"):
        if username:
            d = (remark or "").strip() or (nick or "").strip()
            if d:
                disp_all.setdefault(d, []).append(username)
    con.close()
    disp2 = {u: d for d, us in disp_all.items() for u in us}

    def count_msgs(username):
        tb = "Msg_" + md5hex(username)
        total = 0
        for sh in SHARDS:
            p = os.path.join(decrypted, "message", sh)
            if not os.path.exists(p):
                continue
            cc = sqlite3.connect(p)
            c2 = cc.cursor()
            try:
                ex = c2.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (tb,)).fetchone()
            except Exception:
                ex = None
            if ex:
                try:
                    total += c2.execute(
                        f'SELECT COUNT(*) FROM "{tb}" WHERE create_time>=?', (since_epoch,)).fetchone()[0]
                except Exception:
                    pass
            cc.close()
        return total

    # ===== 决定要导出哪些会话 =====
    if sessions:
        pick = []
        for s in sessions:
            if s in disp2 or "@chatroom" in s:
                pick.append((s, disp2.get(s, s)))
            else:
                for d, us in disp_all.items():
                    if d == s:
                        bu, bn = us[0], -1
                        for u in us:
                            n = count_msgs(u)
                            if n > bn:
                                bu, bn = u, n
                        pick.append((bu, d))
        print(f"[export] 按 config 指定导出 {len(pick)} 个会话")
    else:
        pick = []
        if name_filter:
            for d in sorted(disp_all):
                if name_filter not in d:
                    continue
                cand = [u for u in disp_all[d] if "@chatroom" not in u or u in groups]
                if not cand:
                    continue
                bu, bn = cand[0], -1
                for u in cand:
                    n = count_msgs(u)
                    if n > bn:
                        bu, bn = u, n
                if bn > 0:
                    pick.append((bu, d))
            print(f"[export] 按关键字「{name_filter}」筛出 {len(pick)} 个会话")
        for g in groups:
            pick.append((g, disp2.get(g, g)))
    pick = list(dict.fromkeys(pick))
    if not pick:
        print("✗ 没有匹配到任何会话。请检查 config 的 scope.sessions / scope.name_filter，"
              "或先跑 probe.py sessions 看看有哪些会话。")
        return

    for username, dname in pick:
        tb = "Msg_" + md5hex(username)
        out_rows = []
        for sh in SHARDS:
            p = os.path.join(decrypted, "message", sh)
            if not os.path.exists(p):
                continue
            cc = sqlite3.connect(p)
            c2 = cc.cursor()
            try:
                exists = c2.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (tb,)).fetchone()
            except Exception:
                exists = None
            if not exists:
                cc.close()
                continue
            id2name = {rid: un for rid, un in c2.execute("SELECT rowid, user_name FROM Name2Id")}
            sql = (f'SELECT real_sender_id, create_time, local_type, message_content, '
                   f'WCDB_CT_message_content, compress_content FROM "{tb}" '
                   f'WHERE create_time>=? ')
            params = [since_epoch]
            if until_epoch:
                sql += "AND create_time<=? "
                params.append(until_epoch)
            sql += "ORDER BY create_time"
            rows = c2.execute(sql, params).fetchall()
            for sid, ct, lt, content, ct_type, ccomp in rows:
                # ct_type = WCDB_CT_message_content：疑似"内容类型"标志列（数值型），
                # 不作为正文处理；正文兜底仅取 compress_content（空正文时）。
                if isinstance(content, bytes):
                    content = content.decode("utf-8", "replace")
                if not str(content or "").strip() and isinstance(ccomp, bytes):
                    try:
                        ccomp = ccomp.decode("utf-8", "replace")
                        if ccomp.strip():
                            content = ccomp
                    except Exception:
                        pass
                content = str(content or "")
                uname = id2name.get(sid, "")
                sender_disp = disp2.get(uname) or uname or "(未知)"
                t = (lt & 0xFFFFFFFF) if isinstance(lt, int) and lt >= (1 << 31) else lt
                if t == 3:
                    content = "[图片]"
                elif t == 34:
                    content = "[语音]"
                elif t == 43:
                    content = "[视频]"
                elif t == 47:
                    content = "[表情]"
                elif t == 49:
                    content = "[链接/文件]"
                out_rows.append((ct, sender_disp, content))
            cc.close()
        out_rows.sort(key=lambda r: r[0])
        # 文件名 = 显示名前 40 字符 + wxid 短哈希（同名/超长截断后仍唯一，避免互相覆盖）
        fname = re.sub(r'[\\/:*?"<>|]+', "_", dname)[:40] + "_" + md5hex(username)[:6] + ".txt"
        with open(os.path.join(out, fname), "w", encoding="utf-8") as f:
            f.write(f"# 会话: {dname}  ({username})\n")
            for ct, sd, content in out_rows:
                dt = datetime.fromtimestamp(ct, TZ)
                f.write(f"{dt.strftime('%Y-%m-%d %H:%M')} | {sd} | {content.replace(chr(10),' / ').strip()}\n")
        print(f"{len(out_rows):>4} msgs -> {fname}")
    print("done")


if __name__ == "__main__":
    main()
