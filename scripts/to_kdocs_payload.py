# -*- coding: utf-8 -*-
"""payload -> kdocs rangeData：把脚本产出的回填 payload 转成云文档可用的写入请求。

为什么需要这一层：脚本无法直连 MCP，只能产出 JSON，由 agent 拿着它去调
`sheet.update_range_data`。**用 update_range_data 按坐标写，不用 sheet.add_row**——
add_row 非幂等（重复调用会插多行），网络重试一次就多一行脏数据。

输入（两者都接受）：
  A. editor_sdk 格式（build_matrix_rows / position_reuse 的默认产物）
     {"values":[{"row":0,"col":11,"value_type":"STRING","string_value":"x"},
                {"row":0,"col":12,"value_type":"NUMBER","number_value":46156}]}
  B. 简化格式
     {"values":[{"row":0,"col":11,"value":"x"}]}

输出（可直接作为 sheet.update_range_data 的入参）：
  {"file_id":"...","sheetId":3,"rangeData":[
     {"opType":"formula","rowFrom":0,"rowTo":0,"colFrom":11,"colTo":11,"formula":"x"},
     {"opType":"format","rowFrom":0,"rowTo":0,"colFrom":12,"colTo":12,"xf":{"numfmt":"yyyy-mm-dd"}}]}

用法：
  python to_kdocs_payload.py --payload payload.json --out kdocs_update.json
        [--file-id X] [--sheet-id N] [--date-cols "M,V,W"] [--date-numfmt yyyy-mm-dd]
        [--no-date-format] [--mode values|rangeData]

file_id / sheetId / date_columns 缺省时从 config.json（excel.kdocs / excel.date_columns）读取。
"""
import os
import sys
import json
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from common import load_config, col_index


def _num_to_str(v):
    """数字转单元格文本：整数不带小数点。"""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else repr(v)
    return str(v)


def normalize_values(payload):
    """把两种输入形态统一成 [(row, col, text)]（行列 0-based，已剔除空值）。"""
    vals = payload.get("values") if isinstance(payload, dict) else payload
    out = []
    for it in vals or []:
        if not isinstance(it, dict):
            continue
        try:
            r = int(it["row"])
            c = int(it["col"])
        except (KeyError, TypeError, ValueError):
            continue
        if "value_type" in it:
            if it.get("value_type") == "NUMBER":
                v = it.get("number_value")
            else:
                v = it.get("string_value")
        else:
            v = it.get("value")
        if v is None or (isinstance(v, str) and v == ""):
            continue
        out.append((r, c, _num_to_str(v)))
    return out


def build_range_data(cells, date_cols0=(), date_numfmt="yyyy-mm-dd", with_date_format=True):
    """cells -> rangeData：值用 formula op，日期列额外补 format op（按列合并连续行）。"""
    rd = [{"opType": "formula", "rowFrom": r, "rowTo": r, "colFrom": c, "colTo": c,
           "formula": v} for r, c, v in cells]
    if with_date_format and date_cols0:
        by_col = {}
        for r, c, _v in cells:
            if c in date_cols0:
                by_col.setdefault(c, []).append(r)
        for c, rows in by_col.items():
            for r0, r1 in _runs(sorted(set(rows))):
                rd.append({"opType": "format", "rowFrom": r0, "rowTo": r1,
                           "colFrom": c, "colTo": c, "xf": {"numfmt": date_numfmt}})
    return rd


def _runs(sorted_rows):
    """[1,2,3,7,8] -> [(1,3),(7,8)]：把连续行压成一段，减少请求条目。"""
    runs, start, prev = [], None, None
    for r in sorted_rows:
        if start is None:
            start = prev = r
            continue
        if r == prev + 1:
            prev = r
            continue
        runs.append((start, prev))
        start = prev = r
    if start is not None:
        runs.append((start, prev))
    return runs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--payload", required=True, help="payload.json / payload_n.json")
    ap.add_argument("--out", default=None, help="输出文件；缺省打印到 stdout")
    ap.add_argument("--file-id", default=None)
    ap.add_argument("--sheet-id", default=None)
    ap.add_argument("--date-cols", default=None,
                    help="日期列字母，逗号分隔（如 M,V,W）；缺省读 config.excel.date_columns")
    ap.add_argument("--date-numfmt", default=None, help="日期数字格式，默认 yyyy-mm-dd")
    ap.add_argument("--no-date-format", action="store_true",
                    help="不补 format op（仅写值）")
    ap.add_argument("--config", default=None)
    args = ap.parse_args()

    cfg, cfg_path = load_config(args.config)
    cfg = cfg or {}
    excel_cfg = cfg.get("excel", {}) or {}
    kdocs = excel_cfg.get("kdocs", {}) or {}

    file_id = args.file_id or kdocs.get("file_id")
    sheet_id = args.sheet_id if args.sheet_id is not None else kdocs.get("sheet_id")
    date_numfmt = args.date_numfmt or kdocs.get("date_numfmt") or "yyyy-mm-dd"

    raw_date_cols = args.date_cols
    if raw_date_cols is None:
        cols = excel_cfg.get("date_columns") or []
        raw_date_cols = ",".join(cols) if isinstance(cols, list) else str(cols)
    date_cols0 = []
    for x in str(raw_date_cols or "").replace("，", ",").split(","):
        x = x.strip()
        if x:
            date_cols0.append(col_index(x))

    try:
        with open(args.payload, encoding="utf-8") as f:
            payload = json.load(f)
    except FileNotFoundError:
        sys.exit(f"✗ 找不到 payload：{args.payload}")
    except json.JSONDecodeError as e:
        sys.exit(f"✗ {args.payload} 不是合法 JSON（{e}）")

    cells = normalize_values(payload)
    if not cells:
        print("[warn] payload 里没有可写入的非空单元格，输出空 rangeData。", file=sys.stderr)
    rd = build_range_data(cells, date_cols0, date_numfmt, not args.no_date_format)

    body = {"file_id": file_id or "<file_id>", "sheetId": sheet_id if sheet_id is not None else "<sheetId>",
            "rangeData": rd}
    text = json.dumps(body, ensure_ascii=False, indent=2)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(text)
        rows = sorted({r for r, _c, _v in cells})
        print(f"[kdocs] {len(cells)} 个单元格 / {len({c for _r, c, _v in cells})} 列"
              f" / 行 {rows[0] + 1}~{rows[-1] + 1}（0-based {rows[0]}~{rows[-1]}）"
              if rows else "[kdocs] 无单元格")
        print(f"[kdocs] rangeData {len(rd)} 条"
              + (f"（含 {len(rd) - len(cells)} 条日期格式）" if len(rd) > len(cells) else ""))
        print(f"[kdocs] 写入请求 → {args.out}")
        if "config" not in (cfg_path or ""):
            print(f"[kdocs] config: {cfg_path or '(未找到，缺省值已生效)'}")
        print("  下一步：agent 用 sheet.get_range_data 回读同一区域核对（不信任 code:0）；"
              "写入用 sheet.update_range_data，**不要用 sheet.add_row**（非幂等）。")
        if not file_id or sheet_id is None:
            print("  ⚠ file_id / sheetId 仍是占位符：请在 config.excel.kdocs 里补全，"
                  "或改用 --file-id / --sheet-id 传入。")
    else:
        print(text)


if __name__ == "__main__":
    main()
