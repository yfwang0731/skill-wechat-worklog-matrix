# -*- coding: utf-8 -*-
"""合并子代理 JSON → 预览 CSV（含跨会话重复标记）→ 生成最终行与回填 payload。

**列位置全部按 config.json 的 column_mapping（由 probe.py 依表头生成）决定，
不再写死 A~AA。** 未出现在 mapping 中的逻辑列直接忽略，换模板也能用。

用法：
  1) 预览：python build_matrix_rows.py preview --src <子代理JSON目录> --out <preview.csv>
           （人工裁决：删行序号、合并对）
  2) 终稿：python build_matrix_rows.py final --preview <preview.csv>
           --remove "42,50" --merge "30:52,48:50" [--workbook <xlsx> | --snapshot <json>]

合并对 a:b = 保留 a、删除 b。备注列默认清空（用户要求：单元格只写业务结果，不写判定/分析过程）。
追加起始行来源优先级：--start-row-excel > --snapshot > --workbook > config.excel.start_row。
"""
import json, os, re, csv, argparse, difflib

from common import (load_config, col_index, serial, next_append_row,
                    next_append_row_snapshot)

# 岗位词表：用于从"提出人"文本里剥离岗位（可通过 config.post_words 覆盖/扩展）
DEFAULT_POST_WORDS = ["商务经理", "商务", "客服", "接单客服", "财务", "调度", "前程操作",
                      "后程操作", "后程", "业务员", "保险", "车队", "集港", "散货客服",
                      "操作经理", "操作"]


def get_cfg(args):
    cfg, _ = load_config(args.config)
    return cfg or {}


def post_words(cfg):
    return cfg.get("post_words") or DEFAULT_POST_WORDS


def norm_o(o, cfg):
    """拆分"提出人 + 岗位"，并去掉对方名称前缀关键字（来自 config，不写死）。

    前缀剥离仅在 keyword 后紧跟分隔符（空格/·/、/，等）时执行，
    避免把姓名本身含 keyword 的情况（如 keyword="李四" 而 O="李四财务"）剥空。
    """
    o = re.sub(r"[（(].*?[)）]", "", (o or "").strip()).strip()
    kw = (cfg.get("people", {}) or {}).get("counterparty_keyword") or ""
    if kw:
        # 仅在 kw 作为整词前缀（后随分隔符或行尾）时剥离
        o = re.sub(r"^%s(?:[\s·、,，;；:：]+|$)" % re.escape(kw), "", o).strip()
    post = ""
    for w in sorted(post_words(cfg), key=len, reverse=True):
        if o.endswith(w):
            rest = o[: -len(w)].strip(" .·-–—/\\")
            # 姓名恰好等于岗位词时不要剥空（如 O="客服"、O="财务"），否则提出人会变成"(未知)"
            if rest:
                post, o = w, rest
            break
    return o or "(未知)", post


def load_agents(d, cfg):
    rows = []
    for fn in sorted(os.listdir(d)):
        if not fn.endswith(".json"):
            continue
        try:
            data = json.load(open(os.path.join(d, fn), encoding="utf-8"))
        except Exception as e:
            print(f"skip {fn}: {e}")
            continue
        if not isinstance(data, list):
            print(f"skip {fn}: not a list")
            continue
        rows += data
    for r in rows:
        o, p = norm_o(r.get("O", ""), cfg)
        r["O"] = o
        # 先判空再赋值：键存在但为空串时 setdefault 不会写入，会丢掉从 O 剥离出来的岗位
        if not str(r.get("N") or "").strip():
            r["N"] = p or ""
    # 排序键要先兜 None：子代理 JSON 里出现 "ask_date": null 时，
    # r.get("ask_date", "") 返回 None，会与 str 比较直接 TypeError 崩掉
    rows.sort(key=lambda r: (str(r.get("ask_date") or ""), str(r.get("chat") or "")))
    for i, r in enumerate(rows):
        r["_i"] = i + 1
    return rows


def clean(s):
    return re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]", "", (s or "").lower())


def ids(s):
    """从文本中提取单号/编号特征串。规则：
    - 字母前缀 + ≥6 位数字（如 ANTSQCY250789019）→ 保留；
    - 纯数字 ≥6 位 → 保留，但排除形如 YYYYMMDD 的日期串（如 20260514），
      避免证据里出现的日期把无关需求误判为「同一单号」。
    """
    out = set(re.findall(r"[A-Za-z]{2,}\d{6,}", s or ""))
    for m in re.findall(r"\d{6,}", s or ""):
        if len(m) == 8:
            try:
                y, mo, dd = int(m[:4]), int(m[4:6]), int(m[6:8])
                if 1900 <= y <= 2099 and 1 <= mo <= 12 and 1 <= dd <= 31:
                    continue
            except ValueError:
                pass
        out.add(m)
    return out


def flag_dups(rows, cfg):
    """跨会话疑似重复标记。

    成本：最坏 O(n²)。两道剪枝把它压到实用范围——
      1) 共享单号的候选对直接标记，不进相似度比对；
      2) 相似度通道只在「时间窗 near_days 内」的对上跑，日期在预计算阶段解析一次
         （早期版本每对都 strptime 两次，是主要耗时来源）。
    """
    rules = cfg.get("rules", {}) or {}
    thr = rules.get("similarity_threshold", 0.65)
    near_days = rules.get("near_days", 10)
    if not rules.get("merge_cross_session", True):
        return

    from datetime import datetime as dt

    def _d(s):
        try:
            return dt.strptime(str(s or "")[:10], "%Y-%m-%d")
        except Exception:
            return None

    # 预计算每行的 单号集 / clean 描述 / 时间 / 解析后的日期对象
    prep = []
    for r in rows:
        prep.append({
            "id": ids((r.get("L") or "") + " " + (r.get("evidence") or "")),
            "clean": clean(r.get("L", "")),
            "chat": r.get("chat", ""),
            "date": (r.get("ask_date") or ""),
            "d": _d(r.get("ask_date")),
        })

    # 单号 -> 行下标（仅对"同单号出现 ≥2 次"的行留用）
    id_idx = {}
    for i, p in enumerate(prep):
        for tok in p["id"]:
            id_idx.setdefault(tok, []).append(i)

    def flag(i, j):
        ra, rb = rows[i], rows[j]
        ra["_flag"] = (ra.get("_flag", "") + f" 疑似与#{rb['_i']}重复").strip()
        rb["_flag"] = (rb.get("_flag", "") + f" 疑似与#{ra['_i']}重复").strip()

    done = set()

    # 通道1：共享单号 → 直接标（不要求近时间，跨期重提同单号也算）
    for tok, idxs in id_idx.items():
        if len(idxs) < 2:
            continue
        for a_i in range(len(idxs)):
            for b_i in range(a_i + 1, len(idxs)):
                i, j = idxs[a_i], idxs[b_i]
                if i == j:
                    continue
                key = (min(i, j), max(i, j))
                if key in done:
                    continue
                done.add(key)
                flag(i, j)

    # 通道2：无共享单号但文本高度相似 + 近时间窗（补相似但漏写单号的重复）
    for i in range(len(rows)):
        la, pa = prep[i]["clean"], prep[i]
        if not la:
            continue
        for j in range(i + 1, len(rows)):
            if (i, j) in done:
                continue
            if pa["chat"] == prep[j]["chat"]:
                continue
            lb = prep[j]["clean"]
            if not lb:
                continue
            # 双方都有单号但无交集 → 大概率不同需求，跳过相似比较
            if pa["id"] and prep[j]["id"] and not (pa["id"] & prep[j]["id"]):
                continue
            da, db = pa["d"], prep[j]["d"]
            if da is None or db is None:
                continue            # 日期不可解析 → 不做时间窗内比对（与旧行为一致）
            if abs((da - db).days) > near_days:
                continue
            sim = difflib.SequenceMatcher(None, la, lb).ratio()
            if sim >= thr:
                done.add((i, j))
                flag(i, j)


PREVIEW_COLS = ["序号", "会话", "提出人", "提出人岗位", "提出时间", "需求描述", "需求归类",
                "影响级别", "优先级", "结果", "计划时间", "完成时间", "跨会话标记", "证据节选", "备注"]
QMAP = {"优化或需求": "需求"}
OUTCOME_LABEL = {"done": "答复完成", "default_done": "默认完成(数据修改)",
                 "rejected": "已拒绝(记录)", "no_reply": "无回复(记录)", "pending": "未确认完成"}


def cmd_preview(args):
    cfg = get_cfg(args)
    rows = load_agents(args.src, cfg)
    flag_dups(rows, cfg)
    with open(args.out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(PREVIEW_COLS)
        for r in rows:
            ev = (r.get("evidence") or "").replace("\n", " / ")
            w.writerow([r["_i"], r.get("chat", ""), r["O"], r.get("N", ""), r.get("ask_date", ""),
                        r.get("L", ""), QMAP.get(r.get("Q", ""), r.get("Q", "")),
                        r.get("R", ""), r.get("S", ""),
                        OUTCOME_LABEL.get(r.get("outcome"), r.get("outcome") or ""),
                        r.get("v_date", ""), r.get("w_date", ""), r.get("_flag", "").strip(),
                        ev[:180], r.get("note", "")])
    print(f"preview rows: {len(rows)} -> {args.out}")


def cmd_final(args):
    cfg = get_cfg(args)
    mapping = (cfg.get("excel", {}) or {}).get("column_mapping") or {}
    if not mapping:
        raise SystemExit("✗ 未配置 excel.column_mapping。请先跑 probe.py workbook 生成列映射并写入 config.json。")
    defaults = cfg.get("defaults", {}) or {}
    people = cfg.get("people", {}) or {}
    handler = people.get("handler") or ""
    date_cols = set((cfg.get("excel", {}) or {}).get("date_columns") or [])

    rows = list(csv.DictReader(open(args.preview, encoding="utf-8-sig")))
    seqs = {r["序号"] for r in rows}

    # ---- remove 校验 ----
    remove = {x.strip() for x in (args.remove or "").split(",") if x.strip()}
    bad_rm = remove - seqs
    if bad_rm:
        print(f"⚠ 警告：--remove 中不存在的序号 {sorted(bad_rm)} 已忽略（有效范围 1..{len(rows)}）")

    # ---- merge 校验 + b 字段融合进 a ----
    merge_map = {}
    if args.merge:
        for pair in args.merge.split(","):
            if ":" not in pair:
                print(f"⚠ 警告：跳过格式非法的合并对 {pair!r}（应为 a:b）")
                continue
            a, b = pair.split(":")
            a, b = a.strip(), b.strip()
            if a not in seqs or b not in seqs or a == b:
                print(f"⚠ 警告：跳过无效合并对 {a}:{b}（须存在且不相等）")
                continue
            merge_map[b] = a
    # 保留行（先按序号收集，再处理合并融合）
    keep_map = {}
    for r in rows:
        if r["序号"] not in remove:
            keep_map.setdefault(r["序号"], r)
    if merge_map:
        for b, a in merge_map.items():
            if b in keep_map:          # b 自身不保留
                keep_map.pop(b, None)
            base = keep_map.get(a)
            if base is None:
                continue
            src = next((x for x in rows if x["序号"] == b), None)
            if src is None:
                continue
            # 融合：目标行空字段用 b 的非空值补齐（不覆盖已有值）
            for k, v in src.items():
                if k == "序号":
                    continue
                if not str(base.get(k) or "").strip() and str(v or "").strip():
                    base[k] = v
            base["备注"] = (base.get("备注") or "") + f" [并入原#{b}]"
        print(f"[merge] {len(merge_map)} 对：{', '.join(sorted(f'{b}→{a}' for b, a in merge_map.items()))}（b 补充字段已并入 a）")
    keep = list(keep_map.values())

    # 逻辑列名 -> 值
    final = []
    for r in keep:
        q = QMAP.get(r["需求归类"], r["需求归类"] or "")
        row = dict(defaults)          # 先铺默认值（项目编号/项目名称/固定列）
        row["需求描述"] = (r["需求描述"] or "").strip()
        row["提出时间"] = serial(r["提出时间"])
        row["提出人岗位"] = r["提出人岗位"] or ""
        row["提出人"] = r["提出人"]
        row["需求归类"] = q
        row["影响级别"] = (r["影响级别"] or "").strip() or "中"
        row["优先级"] = (r["优先级"] or "").strip()
        row["计划时间"] = serial(r["计划时间"])
        row["完成时间"] = serial(r["完成时间"])
        if handler:
            for f in ("解决人", "责任人", "项目负责人"):
                if f in mapping:
                    row[f] = handler
        row["备注"] = ""              # 默认留空：不写判定/分析过程
        final.append(row)

    final.sort(key=lambda x: x.get("提出时间") if isinstance(x.get("提出时间"), int) else 10 ** 9)

    outdir = args.out_dir or os.path.dirname(args.preview)
    os.makedirs(outdir, exist_ok=True)
    csv_path = os.path.join(outdir, "final_rows.csv")
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(mapping.keys()))
        w.writeheader()
        for r in final:
            w.writerow({k: r.get(k, "") for k in mapping.keys()})

    # 起始行：显式 > 快照(云文档) > workbook 自动 > config.excel.start_row
    # **不再兜底为 2**：拿不到起始行就报错退出。早期版本 `or 2` 会在忘传
    # --workbook/--snapshot 时静默从第 2 行写，直接覆盖台账开头的数据。
    excel_cfg = cfg.get("excel", {}) or {}
    hint = excel_cfg.get("sheet_match", "运维")
    snap = getattr(args, "snapshot", None)
    wb = args.workbook or excel_cfg.get("workbook")
    start = None
    if args.start_row_excel:
        start = int(args.start_row_excel)
        print(f"[start-row] 显式指定 = {start}")
    elif snap and os.path.isfile(snap):
        start = next_append_row_snapshot(snap, hint, mapping)
        print(f"[start-row] 由云文档快照自动计算 = {start}")
    elif wb and os.path.isfile(wb):
        start = next_append_row(wb, hint, mapping)
        print(f"[start-row] 由工作簿自动计算 = {start}")
    elif excel_cfg.get("start_row"):
        start = int(excel_cfg["start_row"])
        print(f"[start-row] 取自 config.excel.start_row = {start}")
    if not start or start < 2:
        raise SystemExit(
            "✗ 无法确定追加起始行，已中止（避免覆盖台账已有数据）。\n"
            "  请任选其一：--workbook <xlsx> / --snapshot <快照json> / "
            "--start-row-excel <行号> / 在 config.excel.start_row 显式写死。")
    if start == 2:
        print("  ⚠ 起始行为第 2 行（表头下第一行）：确认该行确实是空行，否则会覆盖已有数据。")

    cells = []
    for i, r in enumerate(final):
        for logic, letter in mapping.items():
            v = r.get(logic, "")
            if v in ("", None):
                continue
            row0 = start - 1 + i
            if letter in date_cols:
                cells.append({"row": row0, "col": col_index(letter),
                              "value_type": "NUMBER", "number_value": int(v)})
            else:
                cells.append({"row": row0, "col": col_index(letter),
                              "value_type": "STRING", "string_value": str(v)})
    with open(os.path.join(outdir, "payload.json"), "w", encoding="utf-8") as f:
        json.dump({"values": cells}, f, ensure_ascii=False)
    print(f"final rows: {len(final)} -> {csv_path}, payload -> {os.path.join(outdir,'payload.json')}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    p1 = sub.add_parser("preview")
    p1.add_argument("--src", required=True)
    p1.add_argument("--out", required=True)
    p1.add_argument("--config", default=None)
    p2 = sub.add_parser("final")
    p2.add_argument("--preview", required=True)
    p2.add_argument("--remove", default="")
    p2.add_argument("--merge", default="")
    p2.add_argument("--out-dir")
    p2.add_argument("--workbook", default=None)
    p2.add_argument("--snapshot", default=None,
                    help="云文档快照 json（WPS 通道）；与 --workbook 二选一，用于自动算追加起始行")
    p2.add_argument("--start-row-excel", default=None)
    p2.add_argument("--config", default=None)
    a = ap.parse_args()
    if a.cmd == "preview":
        cmd_preview(a)
    else:
        cmd_final(a)
