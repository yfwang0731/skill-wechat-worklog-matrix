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

输出（`calls` 里每一项**可直接作为 sheet.update_range_data 的 arguments**）：
  {"file_id":"...","worksheet_id":3,"batch_size":100,"total_ops":110,"call_count":2,
   "calls":[
     {"file_id":"...","worksheet_id":3,"rangeData":[
        {"opType":"formula","rowFrom":191,"rowTo":191,"colFrom":11,"colTo":11,"formula":"x"},
        {"opType":"format","rowFrom":191,"rowTo":196,"colFrom":12,"colTo":12,"xf":{"numfmt":"yyyy-mm-dd"}}]},
     {"file_id":"...","worksheet_id":3,"rangeData":[ ... ]} ]}
  （只有 1 批时，顶层额外给一份 rangeData 方便直接取用。）

> 接口要点（2026-09 真机实测 kdocs 连接器 schema）：
> - 工作表参数名是 **worksheet_id**（不是 sheetId）
> - 选区坐标是 camelCase 的 rowFrom/rowTo/colFrom/colTo + opType；日期格式走 `xf.numfmt`
> - **单次 rangeData 最多 100 条**，超出报 `length N exceeds limit 100` → 必须分批（见 calls）
> - 每条 formula op 只能写「一个值」，N 个不同值就是 N 条 op，**无法靠合并收敛**；
>   能合并的只有 format/merge/picture 这类区域操作
> - 写入后必须 `sheet.get_range_data` 回读核对，不能只信 `code: 0`

用法：
  python to_kdocs_payload.py --payload payload.json --out kdocs_update.json
        [--file-id X] [--worksheet-id N] [--date-cols "M,W"] [--date-numfmt yyyy-mm-dd]
        [--no-date-format] [--batch-size 100]

file_id / worksheet_id / date_columns 缺省时从 config.json（excel.kdocs / excel.date_columns）读取。
"""
import os
import re
import sys
import json
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from common import load_config, col_index

# 以这些字符开头的纯文本，表格引擎可能按「公式」解析而不是当文本存
RISKY_PREFIX = ("=", "+", "-", "@")


def find_risky_text(cells):
    """挑出「以 formula op 写入可能被表格引擎改写」的纯文本值。

    返回 [(row0, col0, 值, 原因)]。只做**检测与告警，不改写值** ——
    转义语义（如前置单引号）依赖具体表格引擎，未在真机确认前不擅自加，
    以免把"可能被改写"变成"一定被改写"。命中的交给调用方与用户确认。
    """
    out = []
    for r, c, v in cells or []:
        s = str(v)
        why = None
        if len(s) > 1 and s[0] in RISKY_PREFIX:
            why = f"以 {s[0]!r} 开头，可能被当作公式/负数解析"
        elif re.fullmatch(r"0\d+", s):
            why = "带前导零的数字串，可能被转成数值而丢零"
        elif re.fullmatch(r"\d{16,}", s):
            why = "16 位以上纯数字，Excel 只有 15 位有效数字，可能丢精度"
        if why:
            out.append((r, c, s, why))
    return out


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


def build_range_data(cells, date_cols0=(), date_numfmt="yyyy/m/d", with_date_format=True):
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


def build_body(file_id, worksheet_id, range_data):
    """组装成 sheet.update_range_data 的 arguments（工作表键名必须是 worksheet_id）。"""
    return {"file_id": file_id if file_id else "<file_id>",
            "worksheet_id": worksheet_id if worksheet_id is not None else "<worksheet_id>",
            "rangeData": range_data}


def chunk_range_data(range_data, size=100):
    """按连接器上限切批：单次 update_range_data 的 rangeData **最多 100 条**。

    实测：超过会报 `rangeData length N exceeds limit 100`。
    注意每条 formula op 只能写「一个值」，所以 N 个不同值就是 N 条 op，无法靠合并收敛；
    故大表回填必须分批（幂等，失败可安全重放本批）。
    """
    size = max(1, int(size))
    return [range_data[i:i + size] for i in range(0, len(range_data), size)] or [[]]


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
    ap.add_argument("--worksheet-id", "--sheet-id", dest="worksheet_id", default=None,
                    help="工作表 id（kdocs 连接器参数名 worksheet_id）")
    ap.add_argument("--date-cols", default=None,
                    help="日期列字母，逗号分隔（如 M,V,W）；缺省读 config.excel.date_columns")
    ap.add_argument("--date-numfmt", default=None, help="日期数字格式，默认 yyyy/m/d")
    ap.add_argument("--no-date-format", action="store_true",
                    help="不补 format op（仅写值）")
    ap.add_argument("--batch-size", type=int, default=100,
                    help="单次请求的 rangeData 上限（连接器实测上限 100，超出会被拒）")
    ap.add_argument("--config", default=None)
    args = ap.parse_args()

    cfg, cfg_path = load_config(args.config)
    cfg = cfg or {}
    excel_cfg = cfg.get("excel", {}) or {}
    kdocs = excel_cfg.get("kdocs", {}) or {}

    file_id = args.file_id or kdocs.get("file_id")
    # 兼容旧键名 sheet_id
    ws_id = args.worksheet_id
    if ws_id is None:
        ws_id = kdocs.get("worksheet_id", kdocs.get("sheet_id"))
    date_numfmt = args.date_numfmt or kdocs.get("date_numfmt") or "yyyy/m/d"

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
    risky = find_risky_text(cells)

    chunks = chunk_range_data(rd, args.batch_size)
    calls = [build_body(file_id, ws_id, c) for c in chunks]
    # 单批时同时在顶层给出 rangeData，便于「一次调用就够」的常见场景直接取用
    out = {"file_id": file_id if file_id else "<file_id>",
           "worksheet_id": ws_id if ws_id is not None else "<worksheet_id>",
           "batch_size": max(1, int(args.batch_size)),
           "total_ops": len(rd),
           "call_count": len(calls),
           "calls": calls}
    if len(calls) == 1:
        out["rangeData"] = calls[0]["rangeData"]
    text = json.dumps(out, ensure_ascii=False, indent=2)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(text)
        rows = sorted({r for r, _c, _v in cells})
        print(f"[kdocs] {len(cells)} 个单元格 / {len({c for _r, c, _v in cells})} 列"
              f" / 行 {rows[0] + 1}~{rows[-1] + 1}（0-based {rows[0]}~{rows[-1]}）"
              if rows else "[kdocs] 无单元格")
        print(f"[kdocs] rangeData 共 {len(rd)} 条"
              + (f"（含 {len(rd) - len(cells)} 条日期格式）" if len(rd) > len(cells) else ""))
        print(f"[kdocs] 需分 **{len(calls)}** 次调用"
              f"（单次上限 {max(1, int(args.batch_size))} 条）："
              + "、".join(f"第{i+1}批 {len(c)} 条" for i, c in enumerate(chunks)))
        print(f"[kdocs] 写入请求 → {args.out}")
        if "config" not in (cfg_path or ""):
            print(f"[kdocs] config: {cfg_path or '(未找到，缺省值已生效)'}")
        print("  下一步：agent 逐批调 sheet.update_range_data（**不要用 sheet.add_row**，非幂等），"
              "全部写完后用 sheet.get_range_data 回读核对（不信任 code:0）。")
        if not file_id or ws_id is None:
            print("  ⚠ file_id / worksheet_id 仍是占位符：请在 config.excel.kdocs 里补全，"
                  "或改用 --file-id / --worksheet-id 传入。")
        if risky:
            print(f"  ⚠ 检测到 {len(risky)} 个「以 formula op 写入可能被引擎改写」的纯文本值：")
            for r, c, v, why in risky[:10]:
                print(f"      R{r + 1}C{c + 1}  {v!r} — {why}")
            if len(risky) > 10:
                print(f"      …另有 {len(risky) - 10} 个（完整清单请自行从 payload 里筛）")
            print("      → **写入前先与用户确认这些值是否允许被改写**；"
                  "必须原样保留时，先在 WPS 里把目标列设为「文本」格式再写。")
            print("      （脚本刻意不做自动转义：`'` 前缀等语义依赖具体表格引擎，"
                  "未在真机确认前擅自加会把「可能被改写」变成「一定被改写」。）")
    else:
        print(text)


if __name__ == "__main__":
    main()
