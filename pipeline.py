# -*- coding: utf-8 -*-
"""微信→需求跟踪矩阵 编排器：探测(供一次性确认) + 执行。

两个子命令：
  probe  一次性探测「本机微信账户 / 可选会话 / Excel 表头列映射」，输出汇总供用户一次确认。
  run    确认后执行：解密 → 导出转录 → 分包给并行子代理，并打印后续 LLM/人工步骤。

所有可变项都从 config.json 读；config 由 probe 的探测结果 + 用户选择生成。
用法：
  python pipeline.py probe [--workbook <xlsx>] [--keyword <关键字>]
  python pipeline.py run [--skip-decrypt] [--since <日期>]
"""
import os
import sys
import json
import subprocess
import argparse

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "scripts"))

from common import load_config, find_wcdb_tool


def run(py, *cli):
    print(f"\n$ python {py} {' '.join(cli)}")
    return subprocess.run([sys.executable, os.path.join(HERE, "scripts", py), *cli],
                          cwd=os.path.join(HERE, "scripts"))


# ---------------- probe ----------------
def run_probe_json(cmd):
    """以 --json 模式调用 probe 子命令，返回解析后的 dict；失败返回 {}。"""
    try:
        r = subprocess.run([sys.executable] + cmd,
                           cwd=os.path.join(HERE, "scripts"),
                           capture_output=True, text=True, encoding="utf-8", errors="replace")
        if r.returncode != 0:
            print(r.stderr or r.stdout or f"(probe {cmd[1]} 退出码 {r.returncode})")
            return {}
        # --json 模式下 stdout 应为纯 JSON
        return json.loads(r.stdout or "{}")
    except Exception as e:
        print(f"(解析 probe 输出失败: {e})")
        return {}


def cmd_probe(args):
    cfg, cfg_path = load_config(args.config)
    cfg = cfg or {}
    out = {"config_path": cfg_path, "accounts": [], "sessions": [], "workbook": None}

    print("=" * 60)
    print("① 本机微信账户")
    print("=" * 60)
    acc = run_probe_json(["probe.py", "accounts", "--json"])
    out["accounts"] = acc.get("accounts", [])
    for i, a in enumerate(out["accounts"], 1):
        print(f"  {i}. {a['account_dir']}  (更新于 {a.get('last_modified','')})")

    wb = args.workbook or (cfg.get("excel", {}) or {}).get("workbook")
    if wb and os.path.isfile(wb):
        print("=" * 60)
        print("② Excel 表头 → 列映射")
        print("=" * 60)
        wbj = run_probe_json(["probe.py", "workbook", "--workbook", wb, "--json"])
        out["workbook"] = wbj
        if wbj:
            print(f"  子表「{wbj.get('sheet_selected')}」表头第 {wbj.get('header_row')} 行；"
                  f"列映射 {len(wbj.get('column_mapping', {}))} 项 → 追加起始行 {wbj.get('next_append_row')}")
    else:
        print("\n② 未提供 Excel（--workbook 或 config.excel.workbook），跳过列映射探测。")

    dec = (cfg.get("account", {}) or {}).get("decrypted") or \
        os.path.join(os.getcwd(), "wechat_pilot", "output", "decrypted")
    if os.path.isdir(dec):
        print("=" * 60)
        print("③ 可选会话（来自已解密库）")
        print("=" * 60)
        cmd = ["probe.py", "sessions", "--decrypted", dec, "--json"]
        if args.since:
            cmd += ["--since", args.since]
        if args.keyword:
            cmd += ["--keyword", args.keyword]
        sj = run_probe_json(cmd)
        out["sessions"] = sj.get("sessions", [])
        print(f"  共 {sj.get('total', len(out['sessions']))} 个会话")
    else:
        print(f"\n③ 尚无解密库（{dec}）。选定账户后先跑 run 解密，再重跑 probe 看会话清单。")

    if args.dump:
        with open(args.dump, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        print(f"\n探测结果已写入 {args.dump}")

    print("\n" + "=" * 60)
    print("请把上面①②③的探测结果一次性确认/修正（账户、会话、时间范围、Excel、处理人），")
    print("然后写入 config.json，再执行：python pipeline.py run")
    print("=" * 60)


# ---------------- run ----------------
def cmd_run(args):
    cfg, cfg_path = load_config(args.config)
    cfg = cfg or {}
    if not cfg_path:
        sys.exit("✗ 未找到 config.json。请先复制 config.example.json 并跑 `pipeline.py probe` 探测后填写。")

    if not args.skip_decrypt and not find_wcdb_tool():
        sys.exit(
            "✗ 未找到 wcdb-key-tool。请先准备：\n"
            "  git clone https://github.com/TANGandXUE/wcdb-key-tool "
            f"{os.path.join(HERE, 'scripts', 'tools', 'wcdb-key-tool')}\n"
            "  或设置环境变量 WCDB_KEY_TOOL 指向 wcdb_key_tool_windows.py\n"
            "（合规提示：第三方灰色工具，只读本机，使用前向用户确认。）"
        )

    # 路径统一：以"运行 pipeline 的目录"为基准解析为绝对路径后再传给子进程。
    # （子进程 cwd 是 scripts/，若不显式传路径，相对目录会被解析到 scripts/ 下，
    #   导致解密产物与导出读取位置不一致、打印给用户的路径也找不到文件）
    acct = cfg.get("account", {}) or {}
    out_cfg = cfg.get("output", {}) or {}
    base_abs = os.path.abspath(out_cfg.get("dir") or "./wechat_pilot")
    tx_dir = os.path.join(base_abs, "transcripts")
    out_sub = os.path.join(tx_dir, "_out")
    decrypted = os.path.abspath(acct.get("decrypted") or os.path.join(base_abs, "output", "decrypted"))
    db_storage = acct.get("db_storage") or None

    if not args.skip_decrypt:
        dcli = ["decrypt.py", "--output", decrypted]
        if db_storage:
            dcli += ["--db-storage", db_storage]
        if args.reextract:
            dcli += ["--reextract"]
        rc = run(*dcli)
        if rc.returncode != 0:
            sys.exit("✗ 解密失败。确认微信前台登录后重试（必要时 --reextract）。")

    cli = ["export_conversations.py", "--decrypted", decrypted, "--out", tx_dir]
    if args.since:
        cli += ["--since", args.since]
    rc = run(*cli)
    if rc.returncode != 0:
        sys.exit("✗ 导出失败，见上。")

    n = out_cfg.get("agents", 6)
    rc = run("split_for_agents.py", "--n", str(n),
             "--transcripts", tx_dir, "--out", out_sub)
    if rc.returncode != 0:
        sys.exit("✗ 分包失败，见上。")

    outdir = out_sub
    print("\n========== 下一步（需 LLM + 人工）==========")
    print(f"① 把 references/agent-prompt-zh.txt 模板 + 各 {outdir}/agent_N.txt 清单，")
    print(f"   交给 {n} 个并行子代理，产出 JSON 写到 {outdir}/。")
    print("② 预览并交用户裁决删行/合并：")
    print(f"   python scripts/build_matrix_rows.py preview --src {outdir} --out {outdir}/merged_preview.csv")
    print("③ 用户确认后生成最终行 + 回填 payload：")
    print(f"   python scripts/build_matrix_rows.py final --preview {outdir}/merged_preview.csv "
          f'--remove "<删行序号>" --merge "<a:b>,..."')
    print("④ 岗位列复用：python scripts/position_reuse.py --workbook <xlsx> --out payload_n.json")
    print("⑤ 用 tencent-local-office-edit 把 payload.json / payload_n.json 回填到目标子表，")
    print("   日期列设 number_format yyyy-mm-dd，保存后确认扩展名=.xlsx。")
    print("=============================================")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default=None)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p1 = sub.add_parser("probe", help="探测账户/会话/表头，一次性确认")
    p1.add_argument("--workbook", default=None)
    p1.add_argument("--since", default=None)
    p1.add_argument("--keyword", default=None)
    p1.add_argument("--dump", default=None, help="把探测结果写为 JSON")

    p2 = sub.add_parser("run", help="解密→导出→分包")
    p2.add_argument("--since", default=None)
    p2.add_argument("--skip-decrypt", action="store_true")
    p2.add_argument("--reextract", action="store_true")

    a = ap.parse_args()
    if a.cmd == "probe":
        cmd_probe(a)
    else:
        cmd_run(a)


if __name__ == "__main__":
    main()
