# -*- coding: utf-8 -*-
"""岗位列向上复用计划：按行序向上查同一"提出人"最近的"岗位"值，用于填补空缺。

**列位置由 config.json 的 column_mapping 决定**（提出人 / 提出人岗位），不再写死 N/O。
默认只填空缺（不覆盖已有值）；加 --override 才允许覆盖为上方最近值。

**作用范围（重要）**：历史行只当"岗位来源"读，**永不修改**；填充只发生在起始行起的行。
起始行缺省自动取「追加起始行」（由 --workbook/--snapshot 推算），推不出来就报错退出 ——
早期版本默认从第 2 行扫整表，会把历史里位于"该人首个岗位值之后"的空缺格一并回填
（已实测复现：会静默改写超出用户意图的历史单元格）。

**与 final 的关系（2026-09 起）**：`build_matrix_rows.py final` 已**内置**同一套复用逻辑，
常规流程**一轮写入**即可完成，本脚本不再是必经步骤。保留它是为两个场景：
  1) 事后补跑：把已落表的行补齐岗位；
  2) 自定义范围：--start-row / --end-row 指定任意区间。
两条路共用 `common.reuse_position_fill`，**不存在两套实现**。

**必须显式说明"哪些行算本批新行"**，否则直接报错退出 —— 默认模式无从判断：
追加起始行 = 末数据行 + 1，从那里向下扫恒为空 → 只会得到 0 条（表格尾部若还有
空的"带格式行"，甚至一条提示都不打，属静默无操作）：
  - `--new-rows <final_rows.csv>`（推荐）：从 final 产物取新行，Excel 行号 = 起始行 + 序号，
    **不要求新行已落表**；行号由 --workbook/--snapshot 推算，或显式 --start-row。
  - `--start-row <本批首行号>`：显式指定扫描起点（会扫该行之后的**全部**行，含历史）。
历史行永远只读。

输出 payload（editor_sdk 格式）供回填；云文档通道再经 to_kdocs_payload.py 转成 kdocs rangeData。

依赖：openpyxl（仅本地 --workbook 通道；--snapshot / --history 不需要）
用法：
  python position_reuse.py --workbook <xlsx> --new-rows final_rows.csv --out payload_n.json
  python position_reuse.py --snapshot <json> --new-rows final_rows.csv --out payload_n.json
  python position_reuse.py --history <json> --start-row 192 --new-rows final_rows.csv --out payload_n.json
  python position_reuse.py --workbook <xlsx> --start-row 192 --out payload_n.json [--override]
"""
import csv
import json
import os
import sys
import argparse

from common import (load_config, load_snapshot, workbook_from_snapshot, pick_sheet,
                    next_append_row_ws, position_columns, read_history_positions,
                    reuse_position_fill)

PERSON_KEY = "提出人"
POST_KEY = "提出人岗位"


def load_history(args, sheet, header_row, start, col_person, col_post):
    """历史区 = [header_row+1, start-1]，**只读**。返回 (last_post, pairs)。"""
    if args.history:
        with open(args.history, encoding="utf-8") as f:
            d = {str(k).strip(): str(v).strip()
                 for k, v in ((json.load(f) or {}).get("positions") or {}).items()}
        print(f"[history] 取自 {args.history}：{len(d)} 人")
        return d, None
    d, pairs = read_history_positions(sheet, header_row, start - 1, col_person, col_post)
    print(f"[history] 读表第 {header_row + 1}~{start - 1} 行：{len(d)} 人 / {pairs} 对")
    if not pairs:
        print("  [warn] 历史区没有「提出人+岗位」成对数据：表格来源可能只含表头行"
              "（云文档快照常见）。若确有历史岗位，请用 --history 传入。")
    return d, pairs


def build_records(args, sheet, start, end, col_person, col_post):
    """产出一批待复用的行记录：[{"row": Excel行号, "提出人": 人, "提出人岗位": 岗位}]。

    两种新行来源（CSV / 当前表格）都收敛到这里，后续只调用一次 reuse_position_fill。
    """
    if args.new_rows:
        with open(args.new_rows, encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
        recs = [{"row": start + i,
                 PERSON_KEY: (r.get(PERSON_KEY) or "").strip(),
                 POST_KEY: (r.get(POST_KEY) or "").strip()}
                for i, r in enumerate(rows)]
        print(f"[new-rows] 从 {os.path.basename(args.new_rows)} 读 {len(recs)} 行"
              f"（对应 Excel 第 {start}~{start + len(recs) - 1} 行）")
        return recs
    recs = []
    for r in range(start, end + 1):
        p = sheet.cell(row=r, column=col_person).value
        q = sheet.cell(row=r, column=col_post).value
        recs.append({"row": r,
                     PERSON_KEY: str(p).strip() if p is not None else "",
                     POST_KEY: str(q).strip() if q is not None else ""})
    print(f"[new-rows] 从表格当前内容取第 {start}~{end} 行，共 {len(recs)} 行")
    return recs


def main():
    ap = argparse.ArgumentParser()
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--workbook", help="本地 .xlsx 路径（本地通道）")
    src.add_argument("--snapshot", help="云文档快照 json（WPS 通道）")
    src.add_argument("--history",
                     help='历史岗位 json（形如 {"positions":{"人名":"岗位"}}）：'
                          "表格来源不含 N/O 列数据区时的通道")
    ap.add_argument("--new-rows", default=None,
                    help="本批新行 CSV（final 产出的 final_rows.csv）。给了就从 CSV 取新行，"
                         "不要求新行已落表；不给则从表格当前内容取。")
    ap.add_argument("--out", default="payload_n.json")
    ap.add_argument("--override", action="store_true")
    ap.add_argument("--start-row", type=int, default=None,
                    help="填充起点（Excel 行号）。缺省=自动取追加起始行，"
                         "**只动新增行、历史行仅作岗位来源（只读）**")
    ap.add_argument("--end-row", type=int, default=None)
    ap.add_argument("--config", default=None)
    args = ap.parse_args()

    cfg, _ = load_config(args.config)
    cfg = cfg or {}
    if (cfg.get("rules", {}) or {}).get("reuse_position_column") is False:
        print("[skip] config.rules.reuse_position_column=false，岗位列复用已跳过。")
        with open(args.out, "w", encoding="utf-8") as f:
            json.dump({"values": []}, f, ensure_ascii=False)
        sys.exit(0)
    excel_cfg = cfg.get("excel", {}) or {}
    mapping = excel_cfg.get("column_mapping") or {}
    sheet_hint = excel_cfg.get("sheet_match") or "运维"
    header_row = int(excel_cfg.get("header_row") or 1)

    if not (mapping.get(PERSON_KEY) and mapping.get(POST_KEY)):
        print("⚠ 警告：config.excel.column_mapping 未包含「提出人/提出人岗位」，"
              "将退回默认列位（提出人=O, 岗位=N）。建议先跑 probe.py workbook 生成映射。")
    col_person, col_post = position_columns(mapping)
    print(f"[columns] 提出人=第{col_person}列  岗位=第{col_post}列（来源："
          f"{'表头映射' if mapping.get(PERSON_KEY) else '默认位置'}）")

    # ---- 表格来源（可选：--history 通道不需要）----
    sheet = None
    if args.snapshot:
        sheet = pick_sheet(workbook_from_snapshot(load_snapshot(args.snapshot)), sheet_hint)
        print(f"[source] WPS 快照 {args.snapshot}"
              f"（worksheet_id={getattr(sheet, 'sheet_id', None)}）")
    elif args.workbook:
        import openpyxl
        from openpyxl.utils.exceptions import InvalidFileException
        # 旧版 .xls/.xlt 二进制格式 openpyxl 不支持（.xlsx/.xlsm/.xltx 可以）
        if args.workbook.lower().endswith((".xls", ".xlt")):
            sys.exit("✗ openpyxl 不支持旧版 .xls/.xlt 格式。请先用 Excel/WPS 把工作簿"
                     "「另存为 .xlsx」再运行。")
        try:
            wb = openpyxl.load_workbook(args.workbook)
        except InvalidFileException:
            sys.exit(f"✗ 无法以 .xlsx 解析 {args.workbook}。若是旧版 .xls，请先另存为 .xlsx。")
        sheet = pick_sheet(wb, sheet_hint)
    end = args.end_row or (sheet.max_row if sheet is not None else 0)

    # ---- 锚点：必须明确"哪些行算本批新行" ----
    # 默认模式（既不给 --new-rows 也不给 --start-row）**无从判断本批新行**：
    # 追加起始行 = 末数据行 + 1，从那里向下扫恒为空 → 只会得到 0 条；
    # 若表格尾部恰好还有"空的带格式行"，甚至一条提示都不会打印（静默无操作）。
    # 这属于"看起来跑了、其实什么也没做"，与「绝不静默」原则冲突 → 改为直接报错退出。
    # （早期版本缺省从第 2 行扫整表，会回填历史空缺格，属更严重的静默越界改写，早已移除。）
    if args.start_row:
        start = int(args.start_row)
        print(f"[start-row] 显式指定 = {start}")
    elif args.new_rows:
        if sheet is None:
            sys.exit("✗ --new-rows 需要起点来推算 Excel 行号：请给 --workbook / --snapshot "
                     "让脚本自动推算，或显式 --start-row <本批首行号>。")
        start = next_append_row_ws(sheet, mapping, header_row)
        if not start:
            sys.exit("✗ 无法由表格推算起始行，请显式给 --start-row <本批首行号>。")
        print(f"[start-row] 自动取追加起始行 = {start}（--new-rows 的 Excel 行号 = 该行 + 序号）")
    else:
        sys.exit(
            "✗ 未指定「本批新行」的范围，已中止（避免静默 0 条 / 误改历史行）。\n"
            "  默认模式无法判断哪些行是本批新行：追加起始行 = 末数据行 + 1，\n"
            "  从该行向下扫恒为空 → 只会得到 0 条。请任选其一：\n"
            "    --new-rows <final_rows.csv>   # 推荐：从 final 产物取新行（Excel 行号 = 起始行 + 序号）\n"
            "    --start-row <本批首行号>       # 显式指定扫描起点（会扫该行之后的**全部**行，含历史）")
    start = max(start, header_row + 1)

    if sheet is not None and start > end:
        print(f"[warn] 扫描起点 {start} 超过末行 {end}：该区间为空 → 0 条。"
              f"请确认 --start-row 是否指向本批新行（Excel 行号），"
              f"或表格来源是否覆盖到了数据区。")

    hist, _ = load_history(args, sheet, header_row, start, col_person, col_post)
    recs = build_records(args, sheet, start, end, col_person, col_post)
    changed = reuse_position_fill(recs, hist, PERSON_KEY, POST_KEY, override=args.override)

    cells = [{"row": recs[c["i"]]["row"] - 1, "col": col_post - 1,
              "value_type": "STRING", "string_value": c["new"]} for c in changed]
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump({"values": cells}, f, ensure_ascii=False)
    print(f"sheet: {sheet.title if sheet is not None else '(无表格来源)'} | changed: {len(changed)}")
    for c in changed:
        print(f"  R{recs[c['i']]['row']} {c['person']}: '{c['old']}' -> '{c['new']}'")
    print("payload ->", args.out)


if __name__ == "__main__":
    main()
