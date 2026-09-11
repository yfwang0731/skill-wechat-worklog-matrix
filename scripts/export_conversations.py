# -*- coding: utf-8 -*-
"""从解密后的微信库导出指定会话的转录文本（每行：日期 时间 | 发送者 | 内容）。

**导出哪些会话由 config.json 决定**（scope.sessions 显式清单 或 scope.name_filter 关键字筛选），
不再内置任何具体客户/群。所有路径自动识别或由 config 提供。

正文解码（真机踩坑，重要）：
- `WCDB_CT_message_content` 决定 `message_content` 形态：`0`=明文 str，`4`=**ZSTD 压缩 bytes**
  （魔数 28 B5 2F FD）。图片/引用/合并转发/通话等富媒体都是压缩的，长文本也可能压缩。
- 装 `pip install zstandard` 可完整还原；缺库时输出**明确标记**而非乱码（早期版本吐乱码导致整段会话不可读）。
- `lt=49` appmsg 会提取 `<title>` 与 `<refermsg><content>`（对方用「引用回复」提需求时，正文在引用里）。
- 系统消息（`sysmsg`/撤回）归一化为 `[系统消息]`/`[撤回了一条消息]`，不再把 XML 灌进转录。

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

# ---------------- 消息正文解码 ----------------
# 真机实测（微信 4.1 / 2026-09）：同一张 Msg_* 表里，WCDB_CT_message_content 决定
# message_content 的存储形态：
#   0 → str，明文
#   4 → bytes，**ZSTD 压缩帧**（魔数 28 B5 2F FD）。图片/引用/合并转发/通话等富媒体都是这种，
#       长文本消息也可能被压缩。
# 早期版本把 ct=4 的 bytes 直接 utf-8 errors="replace" 解成文本 → 整段乱码，会话不可读、需求漏判。
# 现在：能解压就解压；解压库缺失/失败则给**明确标记**，绝不吐乱码。
ZSTD_MAGIC = b"\x28\xb5\x2f\xfd"
UNDECODED_TAG = "[压缩内容未解码: 需 pip install zstandard]"
BINARY_TAG = "[二进制内容无法解码]"
MAX_TEXT = 800          # 单条正文上限，防止合并转发记录把转录撑爆


def _zstd_decompress(blob):
    """尝试解压 zstd；库缺失或不支持返回 None（不抛异常）。"""
    for mod in ("zstandard", "pyzstd", "compression.zstd"):
        try:
            if mod == "zstandard":
                import zstandard as z
                return z.ZstdDecompressor().decompress(blob, max_output_size=1 << 24)
            if mod == "pyzstd":
                import pyzstd as z
                return z.decompress(blob)
            from compression import zstd as z
            return z.decompress(blob)
        except ImportError:
            continue
        except Exception:
            return None
    return None


def decode_payload(raw, ct_type=None):
    """把 message_content / compress_content 解成可读文本。

    返回 (text, tag)：tag 非空表示"未能解码"，text 为空时用它当占位。
    """
    if raw is None:
        return "", ""
    if isinstance(raw, (bytes, bytearray)):
        blob = bytes(raw)
        if ct_type == 4 or blob[:4] == ZSTD_MAGIC:
            out = _zstd_decompress(blob)
            if out is None:
                return "", UNDECODED_TAG
            try:
                return out.decode("utf-8", "replace"), ""
            except Exception:
                return "", BINARY_TAG
        txt = blob.decode("utf-8", "replace")
        # 乱码占比过高 → 判定为二进制，不往转录里倒垃圾
        if txt and txt.count("\ufffd") / max(1, len(txt)) > 0.3:
            return "", BINARY_TAG
        return txt, ""
    return str(raw), ""


def _xml_field(xml, tag, limit=MAX_TEXT):
    """取 <tag>...</tag> 的内容（第一个匹配），去掉内层标签残留。"""
    m = re.search(rf"<{tag}>(.*?)</{tag}>", xml or "", re.S)
    if not m:
        return ""
    v = re.sub(r"\s+", " ", m.group(1)).strip()
    return v[:limit]


def appmsg_summary(xml):
    """从 appmsg XML 提取可读摘要：标题 + 描述 + **引用消息的原文**。

    真机案例：对方用「引用回复」提需求时，正文在 <refermsg><content>，早期版本只输出
    `[链接/文件]`，导致需求内容整条丢失（本项目就漏过一条）。
    """
    parts = []
    title = _xml_field(xml, "title", 200)
    if title:
        parts.append(title)
    ref = re.search(r"<refermsg>(.*?)</refermsg>", xml or "", re.S)
    if ref:
        c = _xml_field(ref.group(1), "content", 400)
        if c and c not in parts:
            parts.append(f"引用: {c}")
    des = _xml_field(xml, "des", 400)
    if des and all(des not in p for p in parts):
        parts.append(des)
    if not parts:
        # 兜底：既无 title/des、也不是引用回复的 appmsg，正文可能在**外层的** <content>，
        # 或者只有 <url>。以前这些一律只剩 `[链接/文件]`，纯文本型 appmsg 的需求会整条丢掉。
        # 先把 refermsg 摘掉（上面已处理过），避免把引用正文重复算一次；
        # 只取有语义的字段，不做"去标签取全文"——那会把 appid/fromusername 之类 ID 混进来。
        outer = re.sub(r"<refermsg>.*?</refermsg>", "", xml or "", flags=re.S)
        for tag, limit in (("content", 400), ("url", 300)):
            v = _xml_field(outer, tag, limit)
            if v:
                parts.append(v)
                break
    return " / ".join(parts)


def render_content(raw, ct_type, local_type):
    """消息正文 → 转录里的一行文本（含富媒体占位与系统消息归一化）。"""
    text, tag = decode_payload(raw, ct_type)
    t = local_type
    if isinstance(t, int) and t >= (1 << 31):
        t = t & 0xFFFFFFFF
    if t == 3:
        return "[图片]"
    if t == 34:
        return "[语音]"
    if t == 43:
        return "[视频]"
    if t == 47:
        return "[表情]"
    if tag:
        return tag
    s = text.strip()
    if "<sysmsg" in s or "revokemsg" in s:
        return "[撤回了一条消息]" if "revokemsg" in s else "[系统消息]"
    if "<voipmsg" in s:
        return "[通话]"
    if "<appmsg" in s:
        summary = appmsg_summary(s)
        return f"[链接/文件] {summary}" if summary else "[链接/文件]"
    if t == 49 and not s:
        return "[链接/文件]"
    return s[:MAX_TEXT]


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
        pick, unmatched = [], []
        for s in sessions:
            if s in disp2 or "@chatroom" in s:
                pick.append((s, disp2.get(s, s)))
                continue
            hit = False
            for d, us in disp_all.items():
                if d == s:
                    bu, bn = us[0], -1
                    for u in us:
                        n = count_msgs(u)
                        if n > bn:
                            bu, bn = u, n
                    pick.append((bu, d))
                    hit = True
                    break
            if not hit:
                unmatched.append(s)
        print(f"[export] 按 config 指定导出 {len(pick)} 个会话")
        if unmatched:
            # 早期版本会静默丢弃，用户以为导全了其实少了会话
            print(f"  ⚠ config 指定的 {len(unmatched)} 个会话没匹配到联系人（已跳过）："
                  + "、".join(unmatched[:10])
                  + "；请核对拼写，或先跑 probe.py sessions 看真实标识。")
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

    stats = {"undecoded": 0}
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
                text = render_content(content, ct_type, lt)
                if text in (UNDECODED_TAG, BINARY_TAG):
                    stats["undecoded"] += 1
                elif not text.strip() and str(ccomp or "").strip():
                    # 正文为空时才用 compress_content 兜底（同样走解码器）
                    text = render_content(ccomp, ct_type, lt)
                uname = id2name.get(sid, "")
                sender_disp = disp2.get(uname) or uname or "(未知)"
                out_rows.append((ct, sender_disp, text))
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
    if stats["undecoded"]:
        print(f"\n⚠ {stats['undecoded']} 条消息正文未能解码（多为 ZSTD 压缩的富媒体/长文本）。"
              f"安装解码库可恢复：pip install zstandard")
    print("done")


if __name__ == "__main__":
    main()
