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

from common import (load_config, col_index, serial, parse_date, pick_sheet,
                    counterparty_keyword, next_append_row_ws,
                    load_snapshot, workbook_from_snapshot, snapshot_sheet_coverage,
                    position_columns, warn_if_inside_git_repo,
                    read_history_positions, reuse_position_fill, open_local_workbook,
                    require_dir, require_file, is_placeholder, load_json_file,
                    ensure_utf8_stdio,
                    assert_target_row_empty, assert_no_overlap,
                    read_watermark, write_watermark, WATERMARK_NAME)

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
    o = (o or "").strip()
    # 括号内容**先判是不是岗位词**：`张三（财务）` 的岗位就在括号里，无条件删括号会把它
    # 丢掉（N 列变空，而原文里明明写着）。提示词说的是"去掉前缀 / 括号后缀"，本就隐含
    # "括号里的岗位归到 N"。括号里不是岗位词时照旧删除（那是别名/备注之类）。
    br_post = ""
    _m = re.match(r"^(.*?)[（(]([^（()）]*)[)）]\s*$", o)
    if _m and _m.group(2).strip() in post_words(cfg):
        br_post, o = _m.group(2).strip(), _m.group(1).strip()
    o = re.sub(r"[（(].*?[)）]", "", o).strip()
    kw = counterparty_keyword(cfg)     # people.counterparty_keyword，缺省回落 scope.name_filter
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
    # 尾部岗位优先（更贴近"姓氏+岗位"的书写习惯），括号里的次之
    return o or "(未知)", post or br_post


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
    # 排序键用 parse_date 而不是 ask_date 字符串：`2026/9/5` 与 `2026-09-05` 直接比字符串
    # 会把 10 月排到 9 月前面；而 `final` 用的是 serial() 后的**数值** —— 两处口径必须一致，
    # 否则 preview 的「序号」顺序与最终落表顺序不同，人工按序号做 --remove/--merge 会错位。
    # 解析不出的日期排最后（(1, 0, chat)），与 final 把无日期行沉底的行为对齐。
    def _date_key(r):
        d = parse_date(r.get("ask_date"))
        return (1, 0, str(r.get("chat") or "")) if d is None else (
            0, d.toordinal(), str(r.get("chat") or ""))
    rows.sort(key=_date_key)
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
    # 纯数字要求 **≥8 位**：≥6 位会把电话段、金额、短箱号一网打尽，而通道1 原本
    # "共享单号即标重复、不看时间" —— 两条无关需求只要碰巧含同一串数字就会被并成一条。
    # 带字母前缀的单号（ANTSQCY250789019）不受影响：≥2 字母 + ≥6 位的特征足够强。
    for m in re.findall(r"\d{8,}", s or ""):
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
    # 同会话重复的相似度阈值：**必须高于跨会话的**（见通道2 的注释）。用户把
    # similarity_threshold 调得比 0.9 还高时以用户为准，所以取 max。
    same_chat_thr = max(thr, 0.9)

    # 预计算每行的 单号集 / clean 描述 / 时间 / 解析后的日期对象
    prep = []
    for r in rows:
        prep.append({
            "id": ids((r.get("L") or "") + " " + (r.get("evidence") or "")),
            "clean": clean(r.get("L", "")),
            "chat": r.get("chat", ""),
            "date": (r.get("ask_date") or ""),
            # 走 common.parse_date（全项目唯一入口）：原先这里是严格 strptime，
            # 只认 YYYY-MM-DD，而 serial() 用宽松正则 —— 同一份 2026/9/5 一处认
            # 一处不认，时间窗就此静默失效（失败方式是"少判几条重复"，不报错）。
            "d": parse_date(r.get("ask_date")),
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
                # **同时要求在时间窗内**：单号识别（`ids()`）是启发式的，电话段 / 金额 /
                # 箱号都可能被当成"单号"。不加这道窗，两条毫不相干的需求只要碰巧含同一串
                # 数字就会被并掉 —— 而"并掉"意味着一条真需求消失。跨期重提同单号因此
                # 不再标记（宁可漏标，也不误并）。
                da, db = prep[i]["d"], prep[j]["d"]
                if da is not None and db is not None and abs((da - db).days) > near_days:
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
            # 同一会话里的重复（子代理拆行失误、同一个人重提）**也要看** —— 只是用
            # **更高**的阈值：同一会话的前后文天然相似（同一业务、同一单号），沿用跨会话
            # 阈值会误标一堆本来不相干的行；而拆行失误产生的两行几乎逐字相同，高阈值
            # 照样抓得住。此前这里直接 `continue` 跳过同会话 → 那类重复完全无人管。
            same_chat = pa["chat"] == prep[j]["chat"]
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
            if sim >= (same_chat_thr if same_chat else thr):
                done.add((i, j))
                flag(i, j)


PREVIEW_COLS = ["序号", "会话", "提出人", "提出人岗位", "提出时间", "需求描述", "需求归类",
                "影响级别", "优先级", "结果", "计划时间", "完成时间", "跨会话标记", "证据节选", "备注"]
# final 会直接下标访问的列：缺一个就是旧格式/错文件，必须早报而不是崩在第 N 行
PREVIEW_REQUIRED = ["序号", "提出人", "提出人岗位", "提出时间", "需求描述",
                    "需求归类", "影响级别", "优先级", "计划时间", "完成时间"]

# ---- 判定子代理的产出契约：**唯一真相源是 references/agent-contract.json** ----
# 为什么要有这个文件：这套字段与枚举**提示词里也有一份**（references/agent-prompt-zh.txt
# 的【输出】一节）。两份此前只靠代码注释里一句"一一对应"维系 —— 改一边忘另一边，
# 子代理的**合法**产出就会被判"不合契约"而整批卡住，或者越界值一路顺进台账。
# 现在：代码以 json 为准，自检再比对提示词，单边改动必然红。
AGENT_CONTRACT_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "references", "agent-contract.json")


class ContractError(Exception):
    """产出契约缺失 / 损坏。

    刻意**不抛 SystemExit**：那继承自 BaseException，被 `import` 时会把整个进程杀掉 ——
    自检的第一项（import 冒烟）当场死，后面几十项一条都不跑，人只看到"没输出 + 非零退出"。
    普通异常则能被自检的 `check()` 捕获成 `[FAIL]`，其余检查继续跑；
    而直接当脚本跑时，下面那段会把它转成 SystemExit，仍然只给干净的一句错、不出 traceback。
    """


def load_agent_contract(path=None):
    """读产出契约；缺失 / 损坏时抛 `ContractError`（见上面的取舍说明）。"""
    p = path or AGENT_CONTRACT_PATH
    try:
        with open(p, encoding="utf-8") as f:
            d = json.load(f)
    except FileNotFoundError:
        raise ContractError(f"找不到产出契约 {p} —— 它随 skill 分发，不该缺失。")
    except Exception as e:
        raise ContractError(f"产出契约 {p} 读不出（{type(e).__name__}）：{e}")
    need = ("fields", "required", "q_values", "r_values", "s_values",
            "date_fields", "outcome_label")
    lack = [k for k in need if k not in d]
    if lack:
        raise ContractError(f"产出契约 {p} 缺字段 {lack}")
    return d


try:
    CONTRACT = load_agent_contract()
except ContractError as _e:
    if __name__ == "__main__":
        raise SystemExit(f"✗ {_e}")     # 当脚本跑：干净的一句话 + 非零退出
    raise                                # 被 import（自检）：普通异常 → [FAIL] + 继续跑完

# 「需求归类」的取值域 —— 2026-09-18 从线上台账「需求归类」列**实测 196 格**得到：
# 数据处理 104 / bug 31 / 答疑 27 / 优化 24 / 需求 9。**优化与需求是并存的两种值**。
# 加新取值前先核对真实模板（这列没有下拉校验，是自由文本，所以更需要白名单兜住）。
# 取值域本身登记在 agent-contract.json，这里只是读出来。
Q_VALUES = tuple(CONTRACT["q_values"])
# 子代理偶尔把「优化或需求」连写成一个值（提示词规则文本曾写成"需求或优化"所致），
# 而台账里**没有**这个值 → 归一到「需求」。真机实测出现过 3 条，故保留该映射。
# 2026-09-18 扩容：只收**无歧义**的连写 / 近义写法。有歧义的（"优化需求"同时含两个值）
# 刻意不收 —— 硬归一就是替用户做业务判断，交给下面 normalize_q 的"多义不猜"分支。
QMAP = {"优化或需求": "需求", "功能新增": "需求", "新增功能": "需求",
        "功能改进": "优化", "老功能改进": "优化",
        "故障": "bug", "缺陷": "bug",
        "数据修正": "数据处理", "改数据": "数据处理",
        "咨询": "答疑", "问询": "答疑"}


def normalize_q(v, q_values=Q_VALUES):
    """把「需求归类」归一到取值域；落不进去就**原样返回**，由调用方报错。

    三道，从确定到宽松：
      ① 已在取值域 → 原样；
      ② 命中 QMAP 的连写 / 近义写法 → 取映射值；
      ③ **唯一的**包含匹配（白名单词 ⊂ 值，或值 ⊂ 白名单词）→ 取它。
         0 个候选 = 陌生写法，≥2 个候选 = 有歧义（"优化需求"同时含"优化"与"需求"）
         —— **两种都不猜**。

    为什么宁可不猜：归错了不会报错，只会让台账的取值统计悄悄变形。让人看一眼的成本，
    远低于事后从一堆混值里回捞。
    """
    v = (v or "").strip()
    if not v or v in q_values:
        return v
    if v in QMAP:
        return QMAP[v]
    # `v in q` 这个方向**只在 v 至少两个字时才启用**：否则单字会被包含匹配吞掉 ——
    # "需" 命中「需求」、「化」命中「优化」，而单字根本不是用户想表达的值。
    # （另一个方向 `q in v` 不受此限："系统bug" 本来就该归到 bug。）
    hits = [q for q in q_values if q in v or (len(v) >= 2 and v in q)]
    return hits[0] if len(hits) == 1 else v
OUTCOME_LABEL = dict(CONTRACT["outcome_label"])

# 「确有结论」的 outcome 对应的**中文标签**集合 —— preview CSV 的「结果」列写的是标签，
# 所以任何按「结果」比对的地方都必须用它，不能拿 outcome 英文码去比（那样恒不相等）。
# 单一真相源仍是 agent-contract.json：这里只做一次标签映射，不另抄一份枚举。
_DONE_LABELS = frozenset(OUTCOME_LABEL[k] for k in ("done", "default_done"))

# 契约其余各项同样取自 agent-contract.json —— **别再在别处抄一份**。
# 硬拦的理由：子代理的错值不会当场报错，而是**一路顺到 final 写进台账**；旧实现只做
# 宽松兜底（r.get(...)），字段拼错 / 枚举越界 / 日期写成 2026/9/5 都静默通过。
AGENT_FIELDS = tuple(CONTRACT["fields"])
AGENT_REQUIRED = tuple(CONTRACT["required"])     # 缺任一 → 对应单元格为空
AGENT_R_VALUES = tuple(CONTRACT["r_values"])
AGENT_S_VALUES = tuple(CONTRACT["s_values"])
AGENT_DATE_FIELDS = tuple(CONTRACT["date_fields"])
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
            q_raw = str(r.get("Q") or "").strip()
            q = normalize_q(q_raw)
            if q and q not in Q_VALUES:
                problems.append({"kind": "Q 取值越界", "where": f"{where}:Q",
                                 "msg": f"{q_raw!r} 归不到取值域 {list(Q_VALUES)} 内的任何值"
                                        "（连写 / 近义写法会自动归一；陌生写法与有歧义的写法需人工裁决）"})
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
                        r.get("L", ""), normalize_q(r.get("Q", "")),
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
                         "  请确认 preview 的输入（_out 下 agent 开头的 json，"
                         "真机形态形如 agent<字母>_<账号>.json）确实产出了行。")
    seqs = {r["序号"] for r in rows}

    # ---- remove 校验 ----
    # 序号打错**必须报错**，不能只告警：`--remove` 就是用来删掉"不该进台账"的行的，
    # 静默忽略 = 那些行照旧落表，而调用方以为已经删干净了。
    # 提醒一句：这里的「序号」是 preview CSV 的「序号」列，**不是 Excel 行号** ——
    # 两套编号混用是这套流程里最容易犯的错之一。
    remove = {x.strip() for x in (args.remove or "").split(",") if x.strip()}
    bad_rm = remove - seqs
    if bad_rm:
        raise SystemExit(
            f"✗ --remove 里的序号不存在：{sorted(bad_rm)}\n"
            f"  有效范围 1..{len(rows)}，指的是 preview CSV 的「序号」列（**不是 Excel 行号**）。"
            "请核对后重跑。")

    # ---- merge 校验 + b 字段融合进 a ----
    merge_map = {}
    if args.merge:
        for pair in args.merge.split(","):
            if ":" not in pair:
                raise SystemExit(
                    f"✗ --merge 的合并对格式非法：{pair!r}（应为 a:b —— b 并入 a）")
            a, b = (x.strip() for x in pair.split(":", 1))
            if a not in seqs or b not in seqs:
                raise SystemExit(
                    f"✗ --merge 的合并对 {a}:{b} 里有不存在的序号"
                    f"（有效范围 1..{len(rows)}，preview CSV 的「序号」列）。")
            if a == b:
                raise SystemExit(f"✗ --merge 的合并对 {a}:{b} 两端相同 —— 自己并进自己。")
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
            # **保留较早的那一行**：SKILL.md 承诺"合并一行，提出人取最早提出者"，而此前只做
            # "用 b 补 a 的空字段" —— 若 a 比 b 晚，提出时间与提出人就都不是最早的，台账记下的
            # 不是"谁最早提的"。这里比一次日期，晚的一方退为补充来源。（用 parse_date 比，
            # 不用字符串比：`2026/9/5` 与 `2026-09-05` 直接比字符串会判反。）
            tb, ts = parse_date(base.get("提出时间")), parse_date(src.get("提出时间"))
            if tb and ts and ts < tb:
                base, src = src, base
                keep_map[a] = base
            # 融合：目标行空字段用另一方的非空值补齐（不覆盖已有值）
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
        q = normalize_q(r["需求归类"])
        if q and q not in Q_VALUES and q not in bad_q:
            bad_q[q] = r["序号"]
        row = dict(defaults)          # 先铺默认值（项目编号/项目名称/固定列）
        # 状态列：只给**确有结论**的行保留 defaults 里的"完成"；关态记下的
        # rejected/no_reply/vague/pending 是"待人工裁决"，若也写"完成"，台账会替人得出
        # "已经做完"的结论 —— 而备注列又被强制清空，裁决痕迹就此彻底消失。留空更诚实。
        # ⚠️ 判据必须比对**中文标签**：这里的行来自 preview CSV，那一列写的是
        #    OUTCOME_LABEL[outcome]（"答复完成" / "默认完成(数据修改)"），**不是** outcome 英文码。
        #    本判据曾误写为 `not in ("done", "default_done")` —— 与中文标签恒不相等，
        #    于是**每一行**都被清空、`defaults.状态` 彻底失效（真机 14 行全中，自检当时无断言）。
        if "状态" in mapping and str((r.get("结果") or "")).strip() not in _DONE_LABELS:
            row["状态"] = ""
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
        print("⚠ 以下「需求归类」**归一后仍**不在取值域 " + str(list(Q_VALUES)) +
              " 内，会原样写入。连写 / 近义写法已自动归一，剩下的属于陌生写法或有歧义的写法"
              "（如「优化需求」同时含两个值）—— 请人工裁决后改 preview CSV 再重跑：")
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

    # ---- 落表前的两道复核：目标行必须为空，且不与上一批已生成的区间重叠 ----
    # 上面那些 ⚠ 都只是提示 —— 起始行算错时它们**不会拦住写入**，而回填是覆盖式的：
    # 一旦写下去，历史数据就被盖掉且没有任何检查会发现。所以这里必须硬拦。
    if sheet is None:
        # 没有表格来源（只给了 --start-row-excel 或 config.excel.start_row）就**没得复核**：
        # 一行都读不到，无法证明"第 start 行是空的"。这种情况必须说出来 —— 它恰恰是
        # 手误行号最容易出事的入口。水位那条仍然生效（它不需要读表）。
        print(f"⚠ 未提供表格来源（--workbook / --snapshot），**无法复核第 {start} 行是否为空**：\n"
              "  手误的行号会直接覆盖历史数据。要复核就补上表格来源；"
              "确认行号无误可继续（水位检查不受影响）。")
    assert_target_row_empty(sheet, start, force=args.force)
    assert_no_overlap(start, read_watermark(outdir), force=args.force)

    # ---- 岗位列复用：在本批落表之前算好，随**同一份** payload 一次写入 ----
    # 历史区 = [header_row+1, start-1]，**只读**；填充只发生在**本批行**上。
    do_reuse = getattr(args, "reuse_position", None)
    if do_reuse is None:
        do_reuse = (cfg.get("rules", {}) or {}).get("reuse_position_column", True) is not False
    col_person, col_post = position_columns(mapping)
    if do_reuse and col_person is None:
        # mapping 由 probe 依**表头名**生成：缺这两个键 = 台账本来就没有这两列
        # （「提出人岗位」也不在 probe 的必需列清单里）。没有列可写，跳过是对的，
        # 也没必要打扰用户 —— 旧实现退回写死的 O/N 列位，会把岗位写进完全无关的列。
        print("[reuse] column_mapping 无「提出人/提出人岗位」→ 跳过岗位复用（台账无此列）。")
        do_reuse = False
    if do_reuse:
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

    # 记下本批覆盖的行区间：下次拿同一份快照重跑时，assert_no_overlap 会据此拦住。
    # 语义是"已生成过 payload"而不是"写入成功" —— 脚本管不到写入那一步
    # （kdocs 由 agent 调 update_range_data、local 由 editor_sdk 写），多拦一次远好过漏拦。
    _dates = [r.get("提出时间") for r in final if isinstance(r.get("提出时间"), int)]
    _wm = write_watermark(outdir, start, len(final),
                          max(_dates) if _dates else "",
                          label=os.path.basename(args.preview))
    print(f"[watermark] 本批占第 {_wm['start_row']}~{_wm['last_row']} 行"
          f"（{_wm['generated_at']}）-> {os.path.join(outdir, WATERMARK_NAME)}")


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
    p2.add_argument("--start-row-excel", default=None, type=int,
                    help="显式指定追加起始行（Excel 行号）。不给时由末数据行推断；"
                         "给了也仍会复核该行是否为空 —— 确认要覆盖它才配 --force")
    p2.add_argument("--force", action="store_true",
                    help="跳过落表前的两道复核（① 目标行非空 ② 与上一批生成区间重叠）。"
                         "默认关闭：复核拦下的正是「静默覆盖历史数据」与「重复追加」，别当常规开关用")
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
