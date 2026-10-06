# -*- coding: utf-8 -*-
"""隐私扫描（**独立工具**，不进本仓自检体系）。

扫描面 = `git ls-files` 列出的**文本**文件（扩展名预筛 ＋ 内容含 NUL 判二进制）。
**只报外置清单里登记的业务标识**（客户名 / 项目编号 / 姓名拼音…），按子串匹配；
形状判据（手机号 / 证件号 / 超长数字串）默认**只提示、不阻断**。

## 覆盖边界（务必知悉）
本工具是**补充，不是保证**：只覆盖清单里**已知**的标识，**不覆盖清单外的新客户 / 新项目名**。
清单里没有的标识它一条也扫不出来 —— 别把它当"无隐私泄漏"的证明。

## 清单放哪
清单**不进仓库**（随包目录假定"内容都会分发"，放进去会当场报红）。建议落在仓外，例如
`~/.workbuddy/wechat-worklog-matrix/privacy_patterns.json`；用 `--patterns` 或环境变量
`PRIVACY_PATTERNS` 指过去。

## 绝不静默
清单**缺失或为空**时报错退出（rc≠0），**不**给出"0 命中"这种假绿。

## 用法
    python scripts/privacy_scan.py --selftest                       # 正反用例 ＋ 注入断言
    python scripts/privacy_scan.py --patterns <仓外清单.json> [--root .] [--json]

## 清单格式
    {"patterns": ["某客户", "某项目编号", ...]}   或直接   ["某客户", ...]
"""
import argparse
import json
import os
import re
import subprocess

from common import ensure_utf8_stdio

ENV_VAR = "PRIVACY_PATTERNS"

# 自检专用的**合成标识**：按片段拼装，使整串在源码里**不连读出现**（见 selftest 末的不变量）。
# 沿用全大写＋下划线风格；这些值不指向任何真实业务实体，故塞进真实清单也扫不中和本仓相撞。
_SYNTH_A = "".join(("__SYNTH_", "ALPHA_", "ZZ__"))
_SYNTH_B = "".join(("__SYNTH_", "BETA_", "ZZ__"))

# 便宜的扩展名预筛：这些不当文本读。真正判二进制还看内容是否含 NUL，见 scan_files。
BIN_EXT = {".png", ".jpg", ".jpeg", ".gif", ".ico", ".pdf", ".xlsx", ".zip",
           ".pyc", ".woff", ".woff2", ".db"}

# 形状判据（辅助，默认不阻断）：命中只在报告里列出，不改退出码。
SHAPE_RULES = (
    ("手机号", re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")),
    ("18 位证件号", re.compile(r"(?<!\d)\d{17}[\dXx](?!\d)")),
    ("超长数字串（≥16 位）", re.compile(r"(?<!\d)\d{16,}(?!\d)")),
)


def load_patterns(path):
    """读外置清单。**缺失 / 为空一律报错退出** —— 静默返回空表等于伪造"0 命中"。"""
    path = path or os.environ.get(ENV_VAR) or ""
    if not path:
        raise SystemExit(f"错误：未给清单 —— 用 --patterns 或环境变量 {ENV_VAR} 指定；"
                         "清单缺失时不产出「0 命中」这种假绿。")
    if not os.path.isfile(path):
        raise SystemExit(f"错误：清单文件不存在：{path}")
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError) as e:
        raise SystemExit(f"错误：清单读不出（{type(e).__name__}: {e}）：{path}")
    if isinstance(data, dict):
        if "patterns" not in data:
            keys = "、".join(str(k) for k in data.keys()) or "（无）"
            raise SystemExit(f"错误：未识别到 `patterns` 键：{path} —— 顶层键为 {keys}。"
                             "格式应为 {\"patterns\": [...]} 或直接 [...]；空清单不产出「0 命中」。")
        data = data["patterns"]
    pats = [str(x) for x in data if str(x).strip()] if isinstance(data, list) else []
    if not pats:
        raise SystemExit(f"错误：清单为空：{path} —— 空清单不产出「0 命中」。")
    return pats


def git_files(root):
    """`git ls-files` 给出的文本文件（相对根）。git 不可用 / 非仓库 **报错退出**，不静默降级。"""
    try:
        pr = subprocess.run(["git", "-C", root, "ls-files", "-z"],
                            capture_output=True, timeout=120, env=dict(os.environ))
    except Exception as e:
        raise SystemExit(f"错误：取不到 `git ls-files`（{type(e).__name__}）—— 扫描面依赖它。")
    if pr.returncode != 0:
        raise SystemExit(f"错误：{root} 不是 git 仓库（`git ls-files` 退出码 {pr.returncode}）。")
    out = []
    for rel in pr.stdout.decode("utf-8", "replace").split("\0"):
        if not rel.strip() or os.path.splitext(rel)[1].lower() in BIN_EXT:
            continue
        out.append(rel)
    return out


def scan_files(root, files, patterns):
    """返回 (命中, 形状提示)。命中 = [(rel, 行号, 标识, 行内容)]；形状 = [(rel, 名称, 片段)]。"""
    hits, shapes = [], []
    for rel in files:
        try:
            raw = open(os.path.join(root, rel), "rb").read()
        except OSError:
            continue
        if b"\x00" in raw[:4096]:          # 内容判二进制（扩展名预筛不完备，这里兜一手）
            continue
        text = raw.decode("utf-8", "replace")
        for i, line in enumerate(text.splitlines(), 1):
            for pat in patterns:
                if pat in line:
                    hits.append((rel, i, pat, line.strip()[:120]))
        for name, rx in SHAPE_RULES:
            for m in rx.finditer(text):
                shapes.append((rel, name, m.group(0)))
    return hits, shapes


def selftest():
    """正反用例 ＋ 注入断言：植入一条**清单内**的假标识必须被报红；清单缺失 / 为空必须报错退出。"""
    import tempfile

    with tempfile.TemporaryDirectory(prefix="ps_") as d:
        pats = [_SYNTH_A, _SYNTH_B]

        def _w(name, body):
            with open(os.path.join(d, name), "w", encoding="utf-8") as f:
                f.write(body)

        _w("clean.md", "普通文档，不含任何登记过的标识。\n")
        _w("dirty.md", f"这里出现了登记过的假标识 {_SYNTH_A}。\n")

        hits, _ = scan_files(d, ["clean.md", "dirty.md"], pats)
        assert any(h[0] == "dirty.md" for h in hits), "漏报：清单内已登记的假标识没被报出"
        assert not any(h[0] == "clean.md" for h in hits), "误报：干净文件被判为命中"

        # 注入断言：把标识植入**另一份**文件，必须被报红（证明扫描确实按清单在工作）
        _w("inject.md", f"绕过尝试：把 {_SYNTH_B} 写进这一行。\n")
        inj, _ = scan_files(d, ["inject.md"], pats)
        assert inj, "注入断言失败：植入的清单内标识没被报红"
        assert not scan_files(d, ["inject.md"], ["一个不存在的标识"])[0], \
            "误报：清单外的标识本不该命中"

        # 清单缺失 / 为空 ⇒ 必须报错退出（不给"0 命中"）
        _w("empty.json", '{"patterns": []}')
        for bad in (os.path.join(d, "nope.json"), os.path.join(d, "empty.json")):
            try:
                load_patterns(bad)
            except SystemExit:
                continue
            raise AssertionError(f"漏报：清单 {os.path.basename(bad)} 没有报错退出")
        os.environ.pop(ENV_VAR, None)
        try:
            load_patterns(None)
        except SystemExit:
            pass
        else:
            raise AssertionError("漏报：未给清单也没报错退出")

    # 不变量：示例清单不得自撞 —— 它不得包含任何会被同一份清单命中的字面串。
    # 等价判据：拿示例清单的 patterns 扫本仓，必须 0 命中。示例文件随包分发、本身就在
    # 扫描面里，一旦有人把真实值填进示例 patterns，这里立刻红（不给"示例文件豁免"）。
    ex_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           "privacy_patterns.example.json")
    with open(ex_path, encoding="utf-8") as f:
        ex_data = json.load(f)
    ex_pats = ex_data.get("patterns", []) if isinstance(ex_data, dict) else ex_data
    ex_pats = [str(x) for x in ex_pats if str(x).strip()] if isinstance(ex_pats, list) else []
    if ex_pats:
        ex_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        try:
            ex_files = git_files(ex_root)
        except SystemExit:
            ex_files = []
        ex_rel = os.path.relpath(ex_path, ex_root).replace(os.sep, "/")
        if ex_rel not in ex_files:
            ex_files = ex_files + [ex_rel]
        ex_hits, _ = scan_files(ex_root, ex_files, ex_pats)
        assert not ex_hits, (
            "示例清单自撞：patterns 里的字面串在仓库内也能命中 —— 示范值必须移出 patterns：\n  "
            + "\n  ".join(f"{r}:{n} [{p}]" for r, n, p, _ in ex_hits[:10]))

    # 不变量：自检用的合成标识不得出现在本脚本源码里。它们就是拿来扫别人的字面串，一旦在
    # 源码里连读出现，拿它扫本仓就会命中自身 —— 与「示例清单自撞」同类。片段拼装保证整串不连读。
    here = os.path.abspath(__file__)
    src_hits, _ = scan_files(os.path.dirname(here), [os.path.basename(here)], pats)
    assert not src_hits, (
        "自检数据自撞：合成标识在本脚本源码里也能命中 —— 合成值必须是源码中**不连读出现**的字面串：\n  "
        + "\n  ".join(f"{r}:{n} [{p}]" for r, n, p, _ in src_hits[:10]))

    print("✓ 隐私扫描自检通过（正例 / 反例 / 注入 / 清单缺失与为空 / 示例清单不自撞 / 合成标识不自撞源码）")
    return 0


def main():
    ap = argparse.ArgumentParser(
        description="隐私扫描：只报**仓外清单**里登记的业务标识。它是补充、不是保证 —— "
                    "不覆盖清单外的新客户 / 新项目名；**先把真实业务标识填进仓外清单**，"
                    "空清单 fail-closed（清单缺失或为空即报错退出，绝不静默给 0 命中）。")
    ap.add_argument("--patterns", help="仓外清单（JSON）；不给则读环境变量 %s" % ENV_VAR)
    ap.add_argument("--root", default=".", help="仓库根（默认当前目录）")
    ap.add_argument("--json", action="store_true", help="以 JSON 输出")
    ap.add_argument("--strict-shapes", action="store_true",
                    help="形状判据也计入失败（默认只提示、不改退出码）")
    ap.add_argument("--selftest", action="store_true", help="跑正反用例与注入断言后退出")
    args = ap.parse_args()

    if args.selftest:
        return selftest()

    patterns = load_patterns(args.patterns)
    root = os.path.abspath(args.root)
    files = git_files(root)
    hits, shapes = scan_files(root, files, patterns)

    if args.json:
        print(json.dumps({"files": len(files), "hits": hits, "shapes": shapes},
                         ensure_ascii=False, indent=2))
    else:
        print(f"清单：{len(patterns)} 个标识 | 扫描：{len(files)} 个跟踪的文本文件")
        if hits:
            print(f"✗ 命中 {len(hits)} 处清单内标识：")
            for rel, n, pat, line in hits:
                print(f"  {rel}:{n} [{pat}] {line}")
        else:
            print("✓ 清单内标识 0 命中（仅覆盖清单里**已知**的标识，清单外的新名字不覆盖）")
        if shapes:
            print(f"（提示：形状判据命中 {len(shapes)} 处，默认不阻断）")
            for rel, name, frag in shapes[:20]:
                print(f"  ~ {rel} [{name}] {frag}")

    if hits or (args.strict_shapes and shapes):
        return 1
    return 0


if __name__ == "__main__":
    ensure_utf8_stdio()
    raise SystemExit(main())
