# -*- coding: utf-8 -*-
"""判定规则的行为验证：渲染提示词 + 生成盲评 case + 比对判定结果。

**为什么需要它**：本 skill 的判定规则（什么算需求、什么算完成）只能靠"行为验证" ——
静态断言守得住枚举/开关三处一致这类**形状**，守不住规则本身对不对。而这套流程
（造夹具 → 派独立子代理盲评 → 比对）靠人临时做，规则的正确性就
**没有任何可复跑的东西**。夹具与两端固定下来之后，改规则就能对比"排除方向"与"保留方向"。

两个子命令：

  python rules_check.py plan   --config config.json [--out <dir>] [--only id1,id2]
  python rules_check.py verify --config config.json --results <dir> [--fixtures <json>]

- `plan`   —— 渲染提示词（占位符 + 本批开关覆盖表）→ 写 `<out>/prompt_rendered.txt`；
              并为每条夹具生成一个**不含期望答案**的盲评 case 文件，交给并行子代理。
- `verify` —— 读回子代理产出的 `case_<id>.json`，与夹具标注比对，差异分三类并 exit 非 0。

**刻意不做的事**：本脚本**不**做判定（那是模型的事）、**不**进 CI（非确定性）。
CI 只校"夹具本身是否完备"，那部分在 `scripts/smoke_test.py` 的 `smoke_rules_fixtures`。
"""
import os
import re
import sys
import json
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from common import (load_config, is_placeholder, require_dir, require_file,
                    warn_if_inside_git_repo, ensure_utf8_stdio, counterparty_keyword)

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
FIXTURES = os.path.join(ROOT, "references", "rules-fixtures.json")
PROMPT = os.path.join(ROOT, "references", "agent-prompt-zh.txt")

# 提示词占位符 -> config 取值处（与 SKILL.md「派发判定子代理前」那张表一致）
PLACEHOLDERS = {
    "处理人":       ("people", "handler"),
    "我方标识":     ("people", "my_identifiers"),
    "服务对象":     ("people", "service_object"),
    "业务系统描述": ("excel", "business_system_description"),
    "对方关键字":   ("people", "counterparty_keyword"),
}
# 夹具转录里的发送者 token -> 渲染值（counterparty_keyword + 序字，保证群聊里可区分）
SPEAKER_TOKENS = {"[我方]": "我方", "[对方]": "甲", "[对方2]": "乙", "[对方3]": "丙"}


def load_fixtures(path=FIXTURES):
    d = json.load(open(require_file(path, "夹具（--fixtures）"), encoding="utf-8"))
    return d, d.get("cases") or []


def config_switches(cfg):
    """提示词消费的 4 个判定开关（其余开关走代码，见提示词的开关渲染表末尾说明）。"""
    return ("ignore_if_rejected", "ignore_if_no_reply",
            "ignore_vague_complaints", "data_change_default_done")


def parse_switch_table(prompt):
    """从提示词的【开关渲染表】取出 {开关: (开态句, 关态句)}。

    **唯一实现原则**：关态句的权威来源就是提示词里那张表，脚本只读它、不另存一份，
    否则改了提示词而脚本没跟，就会渲染出一份与文档不符的提示词。
    """
    out = {}
    for line in prompt.splitlines():
        s = line.strip()
        if not s.startswith("|") or "rules." not in s:
            continue
        cols = [c.strip() for c in s.strip("|").split("|")]
        if len(cols) != 3:
            continue
        m = re.match(r"rules\.([a-z_]+)", cols[0])
        if m:
            out[m.group(1)] = (cols[1], cols[2])
    return out


def strip_operator_block(text):
    """去掉开头**给执行代理看**的部分（标题行与 `>` 引用块）。

    模板前几行是操作说明（"把 `{{...}}` 替换为…""下发前按 rules.* 把默认句替换为…"）——
    它们是给下发者看的，发给子代理只会变成噪声；更实际的是：留着它，
    "渲染后不得残留 `{{`" 这条校验就永远为真，等于没守。
    """
    lines = text.splitlines()
    i = 0
    while i < len(lines) and (not lines[i].strip()
                              or lines[i].lstrip().startswith(("#", ">"))):
        i += 1
    return "\n".join(lines[i:])


def render(cfg, prompt, allow_unfilled=False):
    """替换占位符。缺值默认**报错**（SKILL.md 明令「不要发未注入的裸模板」）。"""
    people = cfg.get("people", {}) or {}
    excel = cfg.get("excel", {}) or {}
    missing, unfilled = [], []

    def value_for(sec, key):
        v = (people if sec == "people" else excel).get(key)
        if isinstance(v, list):
            v = " / ".join(x for x in v if x)
        v = (v or "").strip() if isinstance(v, str) else (v or "")
        if not v or is_placeholder(v):
            return None
        return str(v)

    for ph, (sec, key) in PLACEHOLDERS.items():
        v = value_for(sec, key)
        if v is None:
            missing.append(f"{ph}（config.{sec}.{key}）")
            v = f"<未填:{ph}>"
            unfilled.append(ph)
        prompt = prompt.replace("{{" + ph + "}}", v)

    if missing and not allow_unfilled:
        raise SystemExit(
            "✗ 提示词占位符没注入全，拒绝生成（SKILL.md：不要发未注入的裸模板）：\n  "
            + "\n  ".join(missing)
            + "\n  补齐 config 后重跑；确要拿半成品做联调时加 --allow-unfilled（会打醒目告警）。")
    # 模板里出现了没登记的占位符 → 会被原样发出去，必须拦下来
    left = sorted(set(re.findall(r"\{\{([^}]+)\}\}", prompt)))
    if left:
        raise SystemExit(
            f"✗ 渲染后仍有未替换的占位符 {left}。\n"
            "  要么补进本脚本的 PLACEHOLDERS（并同步 config / SKILL.md 的占位符来源表），"
            "要么从模板里删掉。")
    return prompt, unfilled


def render_transcript(lines, cfg):
    kw = counterparty_keyword(cfg)   # 与 build/export 共用同一个"对方关键字"解析口径
    if not kw or is_placeholder(kw):
        kw = "对方"
    handler = ((cfg.get("people", {}) or {}).get("handler") or "").strip()
    if not handler or is_placeholder(handler):
        handler = "我方"
    out = []
    for ln in lines:
        for tok, seq in SPEAKER_TOKENS.items():
            who = handler if tok == "[我方]" else f"{kw} {seq}".strip()
            ln = ln.replace(tok, who)
        out.append(ln)
    return out


def cmd_plan(args):
    cfg, cfg_path = load_config(args.config)
    if not cfg:
        raise SystemExit("✗ 未找到 config.json（用 --config 指定，或设环境变量 WM_CONFIG）。")
    fx, cases = load_fixtures(args.fixtures)
    if args.only:
        want = {x.strip() for x in args.only.split(",") if x.strip()}
        cases = [c for c in cases if c["id"] in want]
        lack = want - {c["id"] for c in cases}
        if lack:
            raise SystemExit(f"✗ --only 里有夹具里不存在的 id：{sorted(lack)}")

    prompt_raw = open(PROMPT, encoding="utf-8").read()
    # 先剥掉给执行代理看的操作说明，再注入占位符（否则说明里那句 `{{...}}` 会永远残留）
    prompt, unfilled = render(cfg, strip_operator_block(prompt_raw), args.allow_unfilled)
    if unfilled:
        print(f"⚠ 以下占位符没填，已渲染成 <未填:…>（这份提示词**不能**直接派发）：{unfilled}")

    table = parse_switch_table(prompt_raw)
    rules = cfg.get("rules", {}) or {}
    lines = ["", "=" * 70,
             "## 本批开关覆盖表（由 rules_check.py 依 config.rules 生成）",
             "=" * 70,
             "下列开关的状态**覆盖**正文中的默认句（正文按「开」态书写）："]
    for sw in config_switches(cfg):
        pair = table.get(sw)
        if pair:
            on_txt, off_txt = pair
        else:
            # 提示词表与这里不同步时**明说**，不静默留空
            on_txt = off_txt = "（提示词开关渲染表里没有这一行！）"
        state = "开" if (rules.get(sw, True) is not False) else "关"
        lines.append(f"- rules.{sw} = **{state}**")
        lines.append(f"    - 开：{on_txt}")
        lines.append(f"    - 关：{off_txt}")
    prompt += "\n".join(lines) + "\n"

    out = os.path.abspath(args.out or os.path.join(os.getcwd(), "rules_check_out"))
    cdir = os.path.join(out, "rules_cases")
    os.makedirs(cdir, exist_ok=True)
    # 产物含**渲染后的真实标识**（处理人、我方标识=微信号、对方关键字…）：
    # 默认 --out 就在 cwd 下，若从仓库根目录跑就会落进仓库 ⇒ 必须提醒。
    warn_if_inside_git_repo(out, "判定行为验证产物")
    pfile = os.path.join(out, "prompt_rendered.txt")
    with open(pfile, "w", encoding="utf-8") as f:
        f.write(prompt)

    for c in cases:
        body = [
            f"# 判定用例 {c['id']}",
            "",
            "【规则】见同目录上一级的 prompt_rendered.txt，它就是本次的判定规则。",
            "",
            "【输入】以下是转录片段，每行格式：`日期 时间 | 发送者 | 内容`",
            "",
        ]
        body += ["  " + ln for ln in render_transcript(c["transcript"], cfg)]
        body += [
            "",
            "【输出】按规则判定后，把结果 JSON 数组写入下面这个路径（**只写要计入矩阵的行**；",
            "       被忽略的需求、以及根本不是需求的内容都不输出）：",
            "",
            f"  {os.path.join(cdir, 'case_' + c['id'] + '.json')}",
            "",
            "  数组元素格式：",
            '  {"chat":"' + c["id"] + '","O":"提出人","N":"岗位或空","ask_date":"YYYY-MM-DD",',
            '   "L":"一句话描述","Q":"数据处理|优化|需求|bug|答疑","R":"高|中|低|空","S":"1-5|空",',
            '   "outcome":"done|default_done|rejected|no_reply|vague|pending",',
            '   "v_date":"YYYY-MM-DD","w_date":"YYYY-MM-DD","evidence":"关键原文","note":""}',
            "",
            "完成后回复：输出行数 + 每行摘要（O|L|outcome）。",
        ]
        with open(os.path.join(cdir, f"case_{c['id']}.txt"), "w", encoding="utf-8") as f:
            f.write("\n".join(body) + "\n")

    states = {sw: ("开" if rules.get(sw, True) is not False else "关") for sw in config_switches(cfg)}
    print(f"[plan] config = {cfg_path}")
    print(f"[plan] 夹具 {len(cases)} 条 → 盲评 case 已写入 {cdir}")
    print(f"[plan] 渲染后的提示词 → {pfile}")
    print(f"[plan] 本批开关状态：{states}")
    print("[plan] 下一步：一个 case 派一个**独立**子代理盲评（case 文件里没有期望答案），"
          "回收后用 `rules_check.py verify` 比对。")


def _expected_state(case, rules):
    """这条夹具该按「开」还是「关」的期望比对 —— 由 config.rules 的实际取值决定。"""
    for sw in case.get("rules") or []:
        if rules.get(sw, True) is False:
            return "off"
    return "on"


def cmd_verify(args):
    cfg, _ = load_config(args.config)
    cfg = cfg or {}
    rules = cfg.get("rules", {}) or {}
    fx, cases = load_fixtures(args.fixtures)
    rdir = require_dir(args.results, "判定结果目录（--results）")

    buckets = {"误记": [], "漏记": [], "结果不符": [], "待裁决": [], "未判定": []}
    for c in cases:
        exp = c["expect"][_expected_state(c, rules)]
        rp = os.path.join(rdir, f"case_{c['id']}.json")
        if not os.path.isfile(rp):
            buckets["未判定"].append(f"{c['id']}：没有 {os.path.basename(rp)}")
            continue
        try:
            rows = json.load(open(rp, encoding="utf-8"))
        except Exception as e:
            buckets["未判定"].append(f"{c['id']}：{os.path.basename(rp)} 读不了（{e}）")
            continue
        if not isinstance(rows, list):
            buckets["未判定"].append(f"{c['id']}：结果不是数组")
            continue
        got_n = len(rows)
        want_n = exp.get("rows", 1) if exp["record"] else 0
        if c.get("open"):
            buckets["待裁决"].append(
                f"{c['id']}：实测记录 {got_n} 行 / outcome={sorted({r.get('outcome') for r in rows}) or '—'}"
                f"，标注为 {want_n} 行 / {exp['outcome']}（规则本身未定义，不计失败）")
            continue
        if not exp["record"] and got_n > 0:
            buckets["误记"].append(f"{c['id']}：应忽略却记了 {got_n} 行 → {[r.get('L') for r in rows][:2]}")
        elif exp["record"] and got_n == 0:
            buckets["漏记"].append(f"{c['id']}：应记录却一行都没输出（{c['why'][:40]}…）")
        elif exp["record"] and got_n != want_n:
            buckets["结果不符"].append(f"{c['id']}：行数 {got_n} ≠ 期望 {want_n}")
        elif exp["record"] and exp["outcome"] is not None:
            got_out = {r.get("outcome") for r in rows}
            if got_out != {exp["outcome"]}:
                buckets["结果不符"].append(
                    f"{c['id']}：outcome {sorted(got_out)} ≠ 期望 {exp['outcome']}")

    print("=" * 72)
    print(f"[verify] 夹具 {len(cases)} 条 / 结果目录 {rdir}")
    print("=" * 72)
    for k in ("误记", "漏记", "结果不符", "待裁决", "未判定"):
        items = buckets[k]
        mark = "  " if not items else ("⚠ " if k == "待裁决" else "✗ ")
        print(f"{mark}{k}：{len(items)}")
        for it in items:
            print(f"    - {it}")
    hard = len(buckets["误记"]) + len(buckets["漏记"]) + len(buckets["结果不符"]) + len(buckets["未判定"])
    print()
    if hard:
        print(f"✗ 有 {hard} 条不通过（含未判定）—— 要么规则判歪了，要么夹具的期望要重新裁定。")
        sys.exit(1)
    print("✓ 全部按标注判定（待裁决项不计入）。")


def main():
    ap = argparse.ArgumentParser(description="判定规则的行为验证：渲染提示词 / 生成盲评 case / 比对结果")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p1 = sub.add_parser("plan", help="渲染提示词 + 生成盲评 case 文件（不含期望答案）")
    p1.add_argument("--config", default=None)
    p1.add_argument("--out", default=None)
    p1.add_argument("--only", default=None, help="只生成指定 id（逗号分隔）")
    p1.add_argument("--fixtures", default=FIXTURES)
    p1.add_argument("--allow-unfilled", action="store_true",
                    help="占位符没填也照样渲染（会打告警；仅用于联调）")

    p2 = sub.add_parser("verify", help="比对子代理的判定结果与夹具标注")
    p2.add_argument("--config", default=None)
    p2.add_argument("--results", required=True, help="含 case_<id>.json 的目录")
    p2.add_argument("--fixtures", default=FIXTURES)

    a = ap.parse_args()
    if a.cmd == "plan":
        cmd_plan(a)
    else:
        cmd_verify(a)


if __name__ == "__main__":
    ensure_utf8_stdio()
    main()
