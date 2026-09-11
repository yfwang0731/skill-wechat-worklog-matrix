# -*- coding: utf-8 -*-
"""wechat-worklog-matrix 冒烟自检：import 全模块 + 关键纯函数断言。

用法：
  python scripts/smoke_test.py            # 无第三方依赖的部分
  python scripts/smoke_test.py --full     # 额外跑依赖 openpyxl 的部分

不触碰真实微信数据/网络，可随时执行。
"""
import sys
import os

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

FAIL = []


def check(name, fn):
    try:
        fn()
        print(f"  [OK] {name}")
    except Exception as e:
        FAIL.append(name)
        print(f"  [FAIL] {name}: {type(e).__name__}: {e}")


def smoke_imports():
    import common            # noqa: F401
    import probe             # noqa: F401
    import decrypt           # noqa: F401
    import export_conversations  # noqa: F401
    import split_for_agents  # noqa: F401
    import build_matrix_rows  # noqa: F401
    import position_reuse    # noqa: F401  (openpyxl 已改为惰性导入，无依赖也能 import)
    import sheet_snapshot    # noqa: F401
    import to_kdocs_payload  # noqa: F401


def smoke_common():
    from common import serial, col_index, col_letter
    assert serial("2026-05-14") == 46156
    assert serial("2026/5/14") == 46156
    assert serial("2026年5月14日") == 46156
    assert serial("2026-05-14 10:00") == 46156
    assert serial("bad") == ""
    assert col_index("A") == 0 and col_index("AA") == 26
    assert col_letter(0) == "A" and col_letter(26) == "AA"


def smoke_snapshot_grid():
    """快照抽象层：稀疏 kdocs 返回 -> 密集网格，且接口与 openpyxl 一致。"""
    from common import (sparse_to_grid, merge_range_data, store_to_rows,
                        GridSheet, GridWorkbook, workbook_from_snapshot,
                        next_append_row_ws, next_append_row_snapshot, pick_sheet)
    rd = [
        {"rowFrom": 0, "rowTo": 0, "colFrom": 0, "colTo": 0, "cellText": "需求描述"},
        {"rowFrom": 0, "rowTo": 0, "colFrom": 1, "colTo": 1, "cellText": "提出时间", "numFormat": "yyyy-mm-dd"},
        {"rowFrom": 0, "rowTo": 0, "colFrom": 2, "colTo": 2, "cellText": "提出人"},
        {"rowFrom": 1, "rowTo": 1, "colFrom": 0, "colTo": 0, "cellText": "改个费用"},
        {"rowFrom": 1, "rowTo": 1, "colFrom": 1, "colTo": 1, "originalCellValue": "46156",
         "cellText": "2026-05-14", "numFormat": "yyyy-mm-dd"},
        {"rowFrom": 1, "rowTo": 1, "colFrom": 2, "colTo": 2, "cellText": "张三"},
    ]
    rows, fmts = sparse_to_grid(rd)
    assert rows[0] == ["需求描述", "提出时间", "提出人"], rows
    assert rows[1][1] == "2026-05-14", rows[1]        # cellText 优先于原始序列号
    assert fmts.get("0,1") == "yyyy-mm-dd"
    # 公式单元格退回原始值
    rows2, _ = sparse_to_grid([{"rowFrom": 0, "rowTo": 0, "colFrom": 0, "colTo": 0,
                                "cellText": "=A1", "originalCellValue": "=A1"}])
    assert rows2[0][0] == "=A1", rows2
    # 稀疏：中间空单元格应被补空
    rows3, _ = sparse_to_grid([{"rowFrom": 0, "rowTo": 0, "colFrom": 0, "colTo": 0, "cellText": "a"},
                               {"rowFrom": 0, "rowTo": 0, "colFrom": 2, "colTo": 2, "cellText": "c"}])
    assert rows3[0] == ["a", "", "c"], rows3
    # range 覆盖（合并单元格语义）：整块填同一个值
    store = merge_range_data({}, [{"rowFrom": 0, "rowTo": 0, "colFrom": 0, "colTo": 3,
                                   "cellText": "x"}])
    assert store_to_rows(store) == [["x", "x", "x", "x"]], store_to_rows(store)
    # 尾部空行应被裁掉（range 覆盖到空单元格时）
    store2 = merge_range_data({}, [{"rowFrom": 0, "rowTo": 0, "colFrom": 0, "colTo": 0,
                                    "cellText": "x"},
                                   {"rowFrom": 1, "rowTo": 5, "colFrom": 0, "colTo": 0,
                                    "cellText": ""}])
    assert store_to_rows(store2) == [["x"]], store_to_rows(store2)
    # GridShell 接口与 openpyxl 同形
    gs = GridSheet("运维-2026", rows)
    assert gs.max_row == 2 and gs.max_column == 3
    assert gs.cell(row=1, column=1).value == "需求描述"
    assert gs.cell(row=9, column=9).value is None      # 越界 -> None（同 openpyxl 空单元格）
    wb = GridWorkbook([gs, GridSheet("说明", [])])
    assert wb.sheetnames == ["运维-2026", "说明"]
    assert pick_sheet(wb, "运维").title == "运维-2026"
    assert pick_sheet(wb, "不存在").title == "运维-2026"
    # 追加起始行：有映射时按"需求描述/提出时间/提出人"列判定
    assert next_append_row_ws(gs, {"需求描述": "A"}) == 3
    assert next_append_row_ws(GridSheet("x", []), {"需求描述": "A"}) == 2
    # 由快照 dict 还原
    wb2 = workbook_from_snapshot({"sheets": [{"name": "运维", "sheetId": 7, "rows": rows}]})
    assert wb2["运维"].sheet_id == 7
    try:
        workbook_from_snapshot({"sheets": []})
    except ValueError:
        pass
    else:
        raise AssertionError("空快照应抛 ValueError")


def smoke_snapshot_raw_parse():
    """raw.json 的多形态识别（kdocs 返回包装不一，必须都吃得下）。"""
    from sheet_snapshot import extract_range_data
    cell = [{"rowFrom": 0, "colFrom": 0, "cellText": "a"}]
    assert extract_range_data({"result": "ok", "detail": {"rangeData": cell}}) == cell
    assert extract_range_data({"rangeData": cell}) == cell
    assert extract_range_data(cell) == cell
    assert extract_range_data({"data": {"rangeData": cell}}) == cell
    assert extract_range_data({"detail": {"detail": {"rangeData": cell}}}) == cell
    import json as _json
    wrapped = {"content": [{"type": "text", "text": _json.dumps({"detail": {"rangeData": cell}})}]}
    assert extract_range_data(wrapped) == cell
    assert extract_range_data({"code": 0}) == []
    assert extract_range_data(None) == []


def smoke_kdocs_payload():
    """payload -> kdocs rangeData 转换（editor_sdk 格式 / 简化格式 / 日期列 numfmt）。"""
    from to_kdocs_payload import (normalize_values, build_range_data, build_body,
                                  _num_to_str, _runs)
    a = {"values": [{"row": 1, "col": 11, "value_type": "STRING", "string_value": "改费"},
                    {"row": 1, "col": 12, "value_type": "NUMBER", "number_value": 46156},
                    {"row": 1, "col": 13, "value_type": "STRING", "string_value": ""}]}
    cells = normalize_values(a)
    assert cells == [(1, 11, "改费"), (1, 12, "46156")], cells   # 空值被剔除
    b = {"values": [{"row": 0, "col": 0, "value": "x"}]}
    assert normalize_values(b) == [(0, 0, "x")]
    assert _num_to_str(46156.0) == "46156" and _num_to_str(1.5) == "1.5"
    assert _runs([1, 2, 3, 7, 8]) == [(1, 3), (7, 8)] and _runs([]) == []
    rd = build_range_data(cells, date_cols0=[12])
    assert rd[0] == {"opType": "formula", "rowFrom": 1, "rowTo": 1, "colFrom": 11,
                     "colTo": 11, "formula": "改费"}, rd[0]
    assert rd[-1]["opType"] == "format" and rd[-1]["xf"]["numfmt"] == "yyyy/m/d", rd
    # 同列连续行压成一段
    nhi = build_range_data([(0, 12, "1"), (1, 12, "2"), (2, 12, "3")], date_cols0=[12])
    fmts = [x for x in nhi if x["opType"] == "format"]
    assert len(fmts) == 1 and fmts[0]["rowFrom"] == 0 and fmts[0]["rowTo"] == 2, nhi
    # 不补格式时只返回值
    assert all(x["opType"] == "formula" for x in build_range_data(cells, [12], with_date_format=False))
    # 关键契约：工作表键名必须是 worksheet_id（连接器参数名），不是 sheetId
    body = build_body("F1", 3, rd)
    assert set(body) == {"file_id", "worksheet_id", "rangeData"}, body.keys()
    assert body["worksheet_id"] == 3 and "sheetId" not in body
    assert build_body(None, None, [])["worksheet_id"] == "<worksheet_id>"


def smoke_probe_analyze():
    """probe.analyze_workbook 应对 GridWorkbook（云文档快照）同样可用。"""
    from common import GridSheet, GridWorkbook
    from probe import analyze_workbook
    rows = [
        ["项目编号", "项目名称", "需求描述", "提出时间", "提出人岗位", "提出人", "解决人"],
        ["P1", "某项目", "改个费用", "2026-05-14", "客服", "张三", "李四"],
        ["", "", "导数据", "2026-05-15", "", "张三", "李四"],
    ]
    res = analyze_workbook(GridWorkbook([GridSheet("运维-2026", rows)]), None, "snap.json")
    assert res["header_row"] == 1
    assert res["column_mapping"]["需求描述"] == "C", res["column_mapping"]
    assert res["column_mapping"]["提出人"] == "F"
    assert res["column_mapping"]["解决人"] == "G"
    assert res["next_append_row"] == 4, res["next_append_row"]
    assert res["date_columns"] == ["D"], res["date_columns"]
    assert res["handler_candidates"][0]["value"] == "李四", res["handler_candidates"]


def smoke_build():
    from build_matrix_rows import norm_o, ids, QMAP
    cfg = {"people": {"counterparty_keyword": "李四"}, "post_words": ["财务", "客服"]}
    assert norm_o("李四财务", cfg) == ("李四", "财务"), norm_o("李四财务", cfg)
    cfg2 = {"people": {"counterparty_keyword": "某公司"}, "post_words": ["客服"]}
    assert norm_o("某公司 王五客服", cfg2) == ("王五", "客服")
    assert "ANTSQCY250789019" in ids("单号 ANTSQCY250789019 改费")
    assert "20260514" not in ids("20260514 提了需求")
    assert QMAP  # 非空


def smoke_position_reuse():
    import position_reuse      # noqa: F401


def smoke_position_reuse_openpyxl():
    """--full：岗位复用需要 openpyxl 的部分（用临时网格走快照通道，不碰真实文件）。"""
    import position_reuse      # noqa: F401
    from common import GridSheet, GridWorkbook
    gs = GridSheet("运维", [
        ["需求描述", "提出人岗位", "提出人"],
        ["a", "客服", "张三"],
        ["b", "", "张三"],
    ])
    assert gs.max_row == 3 and gs.cell(row=2, column=2).value == "客服"


def main():
    full = "--full" in sys.argv
    print("== import 冒烟 ==")
    check("9 个模块 import", smoke_imports)
    print("== common 单测 ==")
    check("serial/col_index/col_letter", smoke_common)
    print("== 云文档快照抽象层 ==")
    check("稀疏->密集网格 / GridSheet 接口 / 追加行", smoke_snapshot_grid)
    check("raw.json 多形态识别", smoke_snapshot_raw_parse)
    check("payload -> kdocs rangeData", smoke_kdocs_payload)
    check("probe.analyze_workbook 吃 GridWorkbook", smoke_probe_analyze)
    print("== build_matrix_rows 单测 ==")
    check("norm_o/ids/QMAP", smoke_build)
    print("== position_reuse ==")
    check("import position_reuse", smoke_position_reuse)
    if full:
        check("positions/快照通道 Cell 读取", smoke_position_reuse_openpyxl)
    else:
        print("  （加 --full 启用附加检查）")

    print()
    if FAIL:
        print(f"✗ 冒烟失败 {len(FAIL)} 项：{FAIL}")
        sys.exit(1)
    print("✓ 全部通过")


if __name__ == "__main__":
    main()
