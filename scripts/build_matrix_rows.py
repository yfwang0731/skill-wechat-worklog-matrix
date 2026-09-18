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
**拿不到来源就报错退出**，绝不兜底为第 2 行。

**岗位列复用已内置**：算出起始行后，从同一个表格来源读「历史区（表头下 ~ 起始行-1）」的
提出人/岗位，把本批空岗位补齐，再写 final_rows.csv 与 payload.json —— 因此**一次写入即可完成**，
不再需要"先写一遍、再跑 position_reuse 补一遍"。关闭方式：--no-reuse-position 或
config.rules.reuse_position_column=false。

云文档通道同样支持表内复用，**前提是第 2 趟读表把关键列按行读全**（快照的 coverage 会记下实际
覆盖范围，行够不到历史区时直接报错、列没覆盖到这两列时告警）。读不到时才用 `--history <json>` ——
那个文件**要手写**，仓里没有生产者。

**preview 会先校验子代理产出的契约**（形状/必填/值域/日期可解析四类硬拦，契约外字段名只告警），
不合契约即报错退出，完整清单落 `<--out 同目录>/validation_report.txt`；**没有跳过开关**。
"""
import json, os, re, csv, argparse, difflib

from common import (load_config, col_index, serial, pick_sheet, next_append_row_ws,
                    load_snapshot, workbook_from_snapshot, snapshot_sheet_coverage,
                    position_columns, warn_if_inside_git_repo,
                    read_history_positions, reuse_position_fill, open_local_workbook,
                    require_dir, require_file, is_placeholder, load_json_file,
                    ensure_utf8_stdio)

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
# final 会直接下标访问的列：缺一个就是旧格式/错文件，必须早报而不是崩在第 N 行
PREVIEW_REQUIRED = ["序号", "提出人", "提出人岗位", "提出时间", "需求描述",
                    "需求归类", "影响级别", "优先级", "计划时间", "完成时间"]

# 「需求归类」的取值域 —— 2026-09-18 从线上台账「需求归类」列**实测 196 格**得到：
# 数据处理 104 / bug 31 / 答疑 27 / 优化 24 / 需求 9。**优化与需求是并存的两种值**。
# 加新取值前先核对真实模板（这列没有下拉校验，是自由文本，所以更需要白名单兜住）。
Q_VALUES = ("数据处理", "优化", "需求", "bug", "答疑")
# 子代理偶尔把「优化或需求」连写成一个值（提示词规则文本曾写成"需求或优化"所致），
# 而台账里**没有**这个值 → 归一到「需求」。真机实测出现过 3 条，故保留该映射。
QMAP = {"优化或需求": "需求"}
OUTCOME_LABEL = {"done": "答复完成", "default_done": "默认完成(数据修改)",
                 "rejected": "已拒绝(记录)", "no_reply": "无回复(记录)",
                 "vague": "笼统抱怨(记录)", "pending": "未确认完成"}

# ---- 判定子代理的产出契约（与 references/agent-prompt-zh.txt 的【输出】一节一一对应）----
# 为什么要在 preview 里硬拦：子代理的错值不会当场报错，而是**一路顺到 final 写进台账**。
# 旧实现只做宽松兜底（r.get(...)），字段拼错 / 枚举越界 / 日期写成 2026/9/5 都静默通过。
AGENT_FIELDS = ("chat", "O", "N", "ask_date", "L", "Q", "R", "S",
                "outcome", "v_date", "w_date", "evidence", "note")
AGENT_REQUIRED = ("O", "L", "ask_date")      # 缺任一 → 对应单元格为空
AGENT_R_VALUES = ("高", "中", "低")
AGENT_S_VALUES = ("1", "2", "3", "4", "5")
AGENT_DATE_FIELDS = ("ask_date", "v_date", "w_date")
# 冒烟/示例产物里常见这些"不是子代理产出"的 json，不能按契约拦（payload 是 dict 而非数组）
AGENT_NAME_RE = re.compile(r"^agent.*\.json$", re.I)


def _agent_files(src):
    """遍历 --src 下的 *.json。

    返回 [(名字, 是否按契约强校验)]：`agent*.json` **无论如何**都要强校验
    （叫这个名却写成 dict，是"该产出没产出"）；其余 json 只有**本身就是数组**时
    才纳入校验（因为 load_agents 会把数组型 json 一并消费，见其实现）；
    payload.json / payload_n.json 这类 dict 产物保持旧行为——跳过。
    """
    out = []
    for fn in sorted(os.listdir(src)):
        if not fn.lower().endswith(".json"):
            continue
        path = os.path.join(src, fn)
        if AGENT_NAME_RE.match(fn):
            out.append((fn, path, True))
            continue
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            out.append((fn, path, False))    # 解析失败留给下面统一报（可能是坏 agent 张冠李戴）
            continue
        if isinstance(data, list):
            out.append((fn, path, True))
    return out


def validate_agents(src, report_path=None):
    """校验 agent 产出契约，返回 (problems, warnings)。

    problems = 会导致「写进台账的值是错的」或「静默失效」，**硬拦**；
    warnings = 契约外的字段名，通常正是必填项缺失的前兆，**只告警**。

    不提供 --skip-validation：与本项目既有立场一致（position_reuse 缺关键输入、final 拿不到
    起始行都是直接报错退出）。真有个别行不想要，正常流程里本来就有裁决环节（preview 之后删行/合并）。
    """
    problems, warnings = [], []
    for fn, path, strict in _agent_files(src):
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            if strict:
                problems.append({"kind": "JSON 无法解析", "where": fn, "msg": f"{type(e).__name__}: {e}"})
            continue
        if not isinstance(data, list):
            if strict:
                problems.append({"kind": "顶层不是数组", "where": fn,
                                 "msg": f"实际是 {type(data).__name__}；子代理应写数组"})
            continue
        for n, r in enumerate(data, 1):
            where = f"{fn}:{n}"
            if not isinstance(r, dict):
                problems.append({"kind": "元素不是对象", "where": where,
                                 "msg": f"实际是 {type(r).__name__}"})
                continue
            for f in sorted(set(r) - set(AGENT_FIELDS)):
                warnings.append({"kind": "契约外的字段名", "where": f"{where}:{f}",
                                 "msg": "提示词改了实现没跟？通常正是必填项缺失的前兆"})
            for f in AGENT_REQUIRED:
                if not str(r.get(f) or "").strip():
                    problems.append({"kind": f"缺必填 {f}", "where": f"{where}:{f}",
                                     "msg": "空或缺失 → 该列会写成空格"})
            q = str(r.get("Q") or "").strip()
            if q and q not in Q_VALUES and q not in QMAP:
                problems.append({"kind": "Q 取值越界", "where": f"{where}:Q",
                                 "msg": f"{q!r} 不在取值域 {list(Q_VALUES)}（连写值会被归一，其余需核模板）"})
            oc = str(r.get("outcome") or "").strip()
            if oc not in OUTCOME_LABEL:
                problems.append({"kind": "outcome 取值越界", "where": f"{where}:outcome",
                                 "msg": f"{oc!r} 不在 {sorted(OUTCOME_LABEL)}"})
            rv = str(r.get("R") or "").strip()
            if rv and rv not in AGENT_R_VALUES:
                problems.append({"kind": "R 取值越界", "where": f"{where}:R",
                                 "msg": f"{rv!r} 不在 {list(AGENT_R_VALUES) + ['']}"})
            sv = str(r.get("S") or "").strip()
            if sv and sv not in AGENT_S_VALUES:
                problems.append({"kind": "S 取值越界", "where": f"{where}:S",
                                 "msg": f"{sv!r} 不在 {list(AGENT_S_VALUES) + ['']}"})
            for f in AGENT_DATE_FIELDS:
                v = str(r.get(f) or "").strip()
                if v and not serial(v):
                    problems.append({"kind": f"{f} 无法解析", "where": f"{where}:{f}",
                                     "msg": f"{v!r} 解析不出日期 → 该格为空，且 final 会把这行沉到排序末尾"})
    return problems, warnings


def render_agent_report(problems, warnings, report_path=None):
    """按「类别 × 文件」汇总：屏幕每类前 10 个具体位置，完整清单落 report_path。

    截断只影响屏幕，不影响文件——上百行时一个共性字段拼错会刷满屏把关键信息淹掉。
    """
    buf = []
    for title, items in (("✗ 必须修（会写进台账或静默失效）", problems),
                         ("⚠ 建议修（不阻断）", warnings)):
        if not items:
            continue
        buf.append(title)
        groups = {}
        for it in items:
            groups.setdefault(it["kind"], []).append(it)
        for kind in sorted(groups, key=lambda k: (-len(groups[k]), k)):
            rows = groups[kind]
            files = {}
            for it in rows:
                head = it["where"].split(":")[0]
                files[head] = files.get(head, 0) + 1
            detail = "、".join(f"{f}×{n}" for f, n in sorted(files.items()))
            buf.append(f"  {kind}：{len(rows)} 处 —— {detail}")
            for it in rows[:10]:
                buf.append(f"      - {it['where']}  {it['msg']}")
            if len(rows) > 10:
                more = f"（其余 {len(rows) - 10} 处见完整清单）" if report_path else "（其余见上）"
                buf.append(f"      … {more}")
        buf.append("")
    text = "\n".join(buf)
    if text.strip():
        print(text)
    if report_path:
        lines = ["# 判定子代理产出契约校验 —— 完整清单", ""]
        for title, items in (("必须修", problems), ("建议修", warnings)):
            lines.append(f"## {title}（{len(items)} 处）")
            if not items:
                lines.append("（无）")
            for it in items:
                lines.append(f"- [{it['kind']}] {it['where']} —— {it['msg']}")
            lines.append("")
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
    return text


def cmd_preview(args):
    cfg = get_cfg(args)
    require_dir(args.src, "转录目录（--src）")
    # 契约校验必须在 load_agents **之前**：load_agents 遇到坏 json 只 print 后 skip，
    # 错值会一路顺到 final 写进台账。报告落 --out 同目录的 validation_report.txt。
    out_dir = require_dir(os.path.dirname(os.path.abspath(args.out)), "--out 所在目录")
    report = os.path.join(out_dir, "validation_report.txt")
    # 报告里会出现 agent 文件名与违约值（真机文件名含人名拼音）⇒ 落进仓库要提醒
    warn_if_inside_git_repo(out_dir, "产出契约校验报告")
    problems, warnings = validate_agents(args.src, report)
    render_agent_report(problems, warnings, report)
    if problems:
        raise SystemExit(
            f"✗ 判定产出有 {len(problems)} 处不合契约，已拒绝生成 preview。\n"
            f"  完整清单：{report}\n"
            "  请修子代理产出（或提示词）后重跑 —— 本命令**没有**跳过开关；\n"
            "  个别行不想要，请在 preview 之后的裁决环节删行/合并。")
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

    with open(require_file(args.preview, "预览 CSV（--preview）"), encoding="utf-8-sig") as f:
        rdr = csv.DictReader(f)
        # 列名取自 fieldnames 而不是 rows[0]：只剩表头（0 数据行）时前者仍有效，
        # 后者会退化成空集合 → 误报"缺少必需列"，把"空表"说成"列不对"。
        have = set(rdr.fieldnames or [])
        rows = list(rdr)
    lack = [c for c in PREVIEW_REQUIRED if c not in have]
    if lack:
        raise SystemExit(
            f"✗ {args.preview} 缺少必需列 {lack}。\n"
            f"  现有列：{sorted(have)}\n"
            "  请确认传入的是 **probe 之后的 preview 产物**（列名如「提出人 / 提出时间 / 需求描述」）；"
            "旧版产物用的是带列字母的列名（如「提出人O / 岗位N」），不能直接喂给 final。")
    if not rows:
        # 0 数据行时旧实现会安静地写出空 final_rows.csv + 空 payload 并 exit 0 ——
        # 属"看起来跑了其实没做事"，与「绝不静默」冲突：明确说出来，让人去查上游。
        raise SystemExit(f"✗ {args.preview} 只有表头、没有任何数据行 —— 没有可生成的行。\n"
                         "  请确认 preview 的输入（_out 下的 agent_*.json）确实产出了行。")
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
            # 这里**不**往 备注 里写「并入原#b」：备注列一律只放业务内容，
            # 且 final 随后会把备注清空，写了也落不了地（合并信息由下面那行 stdout 给出）。
        print(f"[merge] {len(merge_map)} 对：{', '.join(sorted(f'{b}→{a}' for b, a in merge_map.items()))}（b 补充字段已并入 a）")
    keep = list(keep_map.values())

    # 逻辑列名 -> 值
    final = []
    bad_q = {}
    for r in keep:
        q = QMAP.get(r["需求归类"], r["需求归类"] or "")
        if q and q not in Q_VALUES and q not in bad_q:
            bad_q[q] = r["序号"]
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

    if bad_q:
        print("⚠ 以下「需求归类」不在台账现有取值域 " + str(list(Q_VALUES)) +
              " 内，会**原样写入**，请核对是否该归一：")
        for v, seq in sorted(bad_q.items()):
            print(f"    #{seq}  {v!r}")

    final.sort(key=lambda x: x.get("提出时间") if isinstance(x.get("提出时间"), int) else 10 ** 9)

    outdir = args.out_dir or os.path.dirname(args.preview)
    os.makedirs(outdir, exist_ok=True)
    csv_path = os.path.join(outdir, "final_rows.csv")

    # ---- 表格来源：只加载一次，供「起始行」与「历史岗位」共用 ----
    # 来源优先级：--start-row-excel > --snapshot > --workbook > config.excel.start_row
    # **不再兜底为 2**：拿不到起始行就报错退出。早期版本 `or 2` 会在忘传
    # --workbook/--snapshot 时静默从第 2 行写，直接覆盖台账开头的数据。
    excel_cfg = cfg.get("excel", {}) or {}
    hint = excel_cfg.get("sheet_match", "运维")
    header_row_cfg = int(excel_cfg.get("header_row") or 1)
    snap = getattr(args, "snapshot", None)
    wb_path = args.workbook or excel_cfg.get("workbook")
    # 表格来源校验分两类，判据是**意图强度**不是"文件在不在"：
    #   * 命令行显式给出 → 强意图：不存在就报错，绝不退回 config.excel.start_row
    #     （曾把"路径打错"当成"没给来源"：两个分支都不进 → 起点静默改用 config 值，
    #       并且以 exit 0 产出 payload，与本节"拿不到起始行就报错退出"的原则冲突。）
    #   * 来自 config 且是 `<...>` 占位符 → 是"没配置"（首次使用的正常状态），告警后按未提供处理
    #   * 来自 config 的真实路径 → 已配置：不存在即报错（岗位复用也要读它，静默跳过会改结果）
    if snap:
        require_file(snap, "快照（--snapshot）")
    if args.workbook:
        require_file(args.workbook, "工作簿（--workbook）")
    sheet, src, snap_cov = None, None, None
    if snap:
        snap_obj = load_snapshot(snap)
        sheet = pick_sheet(workbook_from_snapshot(snap_obj), hint)
        snap_cov = snapshot_sheet_coverage(snap_obj, sheet.title)
        src = f"云文档快照 {os.path.basename(snap)}"
    elif wb_path and is_placeholder(wb_path):
        print(f"⚠ config.excel.workbook 还是 config.example.json 里的占位符"
              f"（{wb_path}）→ 视为未提供表格来源。")
    elif wb_path:
        sheet = pick_sheet(open_local_workbook(wb_path), hint)
        src = f"工作簿 {os.path.basename(wb_path)}"

    start = None
    if args.start_row_excel:
        start = int(args.start_row_excel)
        print(f"[start-row] 显式指定 = {start}")
    elif sheet is not None:
        start = next_append_row_ws(sheet, mapping, excel_cfg.get("header_row") or 1)
        print(f"[start-row] 由 {src} 自动计算 = {start}")
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
        if sheet is not None and sheet.max_row <= header_row_cfg:
            print("  ⚠ 表格来源**只有表头行**（未覆盖数据区），无法据此判断末数据行。"
                  "若台账已有数据，请改用 --start-row-excel <行号>，或按 SKILL.md"
                  "「WPS 通道读表」补读数据区后重建快照。")

    # ---- 岗位列复用：在本批落表之前算好，随**同一份** payload 一次写入 ----
    # 历史区 = [header_row+1, start-1]，**只读**；填充只发生在**本批行**上。
    do_reuse = getattr(args, "reuse_position", None)
    if do_reuse is None:
        do_reuse = (cfg.get("rules", {}) or {}).get("reuse_position_column", True) is not False
    if do_reuse:
        col_person, col_post = position_columns(mapping)
        if not (mapping.get("提出人") and mapping.get("提出人岗位")):
            print(f"⚠ column_mapping 缺「提出人/提出人岗位」，复用退回默认列位"
                  f"（提出人=第{col_person}列, 岗位=第{col_post}列）。")
        # ---- 快照覆盖范围的两条检查（只在要复用时才做：不复用时数据区根本不参与计算）----
        # 判据来自 build 写进快照的 coverage（裁边前的原始极值）。旧快照没有该字段 → 跳过，不误报。
        if snap_cov:
            cov_row_to, cov_col_to = snap_cov.get("rowTo"), snap_cov.get("colTo")
            if cov_row_to is not None and (int(cov_row_to) + 1) < start - 1:
                # 确定错误：读回来的行数够不到历史区，绝不可能复用出岗位
                raise SystemExit(
                    f"✗ 快照只覆盖到第 {int(cov_row_to) + 1} 行，而追加起始行是第 {start} 行 ——\n"
                    f"  它**不可能**包含历史区（第 {header_row_cfg + 1}~{start - 1} 行）"
                    "，岗位复用无从下手。\n"
                    "  通常说明「第 2 趟读表」没做，或 rowTo 没读到表末。请按 SKILL.md"
                    "「WPS 通道读表」把关键列**按行读全**后重建快照；\n"
                    "  确实不需要岗位复用时，加 --no-reuse-position 跳过本检查，"
                    "或直接用 --start-row-excel 而不要带快照。")
            if cov_col_to is not None and (int(cov_col_to) + 1) < col_post:
                # 两可：可能没读这两列，也可能读到了但该列在历史区里本来就全空
                print(f"⚠ 快照最右只到第 {int(cov_col_to) + 1} 列，而「提出人岗位」在第 {col_post} 列"
                      f"（「提出人」第 {col_person} 列）→ 本次读表没覆盖这两列，岗位复用无从下手。\n"
                      "  请在第 2 趟读表时把这两列纳入 letters / colFrom-colTo，再重建快照。")
        hist = None
        if getattr(args, "history", None):
            hist = {str(k).strip(): str(v).strip()
                    for k, v in ((load_json_file(args.history, "历史岗位 JSON（--history）")
                                  or {}).get("positions") or {}).items()}
            print(f"[reuse] 历史岗位取自 --history（{len(hist)} 人）")
        elif sheet is not None:
            hist, pairs = read_history_positions(sheet, header_row_cfg, start - 1,
                                                 col_person, col_post)
            print(f"[reuse] 读历史区第 {header_row_cfg + 1}~{start - 1} 行 → "
                  f"{len(hist)} 人 / {pairs} 对")
            if not pairs:
                print("  ⚠ 历史区没有「提出人+岗位」成对数据 —— 三种可能，请对照刚才的 [reuse] 行数判断：\n"
                      "    ① 快照/工作簿只读了表头行（行没读全，见上方 coverage 检查与 SKILL.md「WPS 通道读表」）；\n"
                      "    ② 读到了这两列，但历史区里确实从没填过岗位（这时复用本来就救不了）；\n"
                      "    ③ 只给了 --start-row-excel。本批不补岗位。")
        else:
            print("⚠ 跳过岗位复用：未提供 --workbook/--snapshot，读不到历史岗位。"
                  "需要时请给表格来源，或用 --history <json> 传入。")
        if hist is not None:
            ch = reuse_position_fill(final, hist)
            for c in ch:
                print(f"  [reuse] R{start + c['i']} {c['person']}: "
                      f"'{c['old']}' -> '{c['new']}'")
            print(f"[reuse] 补 {len(ch)} 个岗位（随 payload 一次写入）" if ch else
                  "[reuse] 无需补岗位：本批要么已有岗位，要么历史里没有这些人的岗位")
    else:
        print("[reuse] 已跳过岗位复用（--no-reuse-position 或 "
              "config.rules.reuse_position_column=false）。")

    # final_rows.csv 写在复用之后：**保证 CSV 与真正写入表格的内容一致**。
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(mapping.keys()))
        w.writeheader()
        for r in final:
            w.writerow({k: r.get(k, "") for k in mapping.keys()})

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
    ensure_utf8_stdio()
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
    rs = p2.add_mutually_exclusive_group()
    rs.add_argument("--reuse-position", dest="reuse_position", action="store_true", default=None,
                    help="强制开启岗位列复用（默认取 config.rules.reuse_position_column）")
    rs.add_argument("--no-reuse-position", dest="reuse_position", action="store_false",
                    help="本次跳过岗位列复用")
    p2.add_argument("--history", default=None,
                    help='历史岗位 json（形如 {"positions":{"人名":"岗位"}}）。**要你自己手写** ——'
                         "仓里没有任何脚本会生成这个文件；表内能读到历史时（两条通道都行，"
                         "云文档需第 2 趟把行读全）不要用它")
    p2.add_argument("--config", default=None)
    a = ap.parse_args()
    if a.cmd == "preview":
        cmd_preview(a)
    else:
        cmd_final(a)
