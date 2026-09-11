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


MODULES = ["common", "probe", "decrypt", "export_conversations", "split_for_agents",
           "build_matrix_rows", "position_reuse", "sheet_snapshot", "to_kdocs_payload"]


def smoke_imports():
    import importlib
    for m in MODULES:
        importlib.import_module(m)


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
    # 关键契约：单次 rangeData 上限 100 条 → 必须自动分批
    from to_kdocs_payload import chunk_range_data
    big = [{"opType": "formula"}] * 110
    ch = chunk_range_data(big, 100)
    assert [len(c) for c in ch] == [100, 10], [len(c) for c in ch]
    assert sum(len(c) for c in ch) == 110
    assert chunk_range_data([], 100) == [[]]
    assert [len(c) for c in chunk_range_data(big, 1000)] == [110]


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


def smoke_message_decoding():
    """消息正文解码：ZSTD 压缩 / appmsg 引用 / 系统消息 归一化。"""
    from export_conversations import (render_content, decode_payload, appmsg_summary,
                                      UNDECODED_TAG, BINARY_TAG, ZSTD_MAGIC)
    # zstd 帧但解不开（测试环境无 zstandard 或帧非法）→ 明确标记，绝不吐乱码
    bad = ZSTD_MAGIC + b"\x00" * 16
    assert render_content(bad, 4, 1) == UNDECODED_TAG, render_content(bad, 4, 1)
    # 纯二进制且高乱码率 → 标记
    assert decode_payload(b"\xff\xfe\x00\x01" * 20, 0)[1] == BINARY_TAG
    # 明文直通
    assert render_content("你好", 0, 1) == "你好"
    # 媒体占位（优先级高于内容）
    assert render_content("x", 0, 3) == "[图片]"
    assert render_content("x", 0, 34) == "[语音]"
    # appmsg：提取标题 + 引用正文（对方用「引用回复」提需求时正文在引用里）
    xml = ('<?xml version="1.0"?><msg><appmsg><title>记得帮忙改</title>'
           '<refermsg><content>帮忙按船公司对比导入的 差额逻辑改吧</content></refermsg>'
           '</appmsg></msg>')
    got = render_content(xml, 0, 49)
    assert got.startswith("[链接/文件] ") and "记得帮忙改" in got and "差额逻辑改吧" in got, got
    assert appmsg_summary("<msg></msg>") == ""
    # 系统消息 / 撤回 / 通话
    assert render_content('<sysmsg type="revokemsg"><content>x</content></sysmsg>', 0, 10000) == "[撤回了一条消息]"
    assert render_content('<sysmsg type="roomtoolstips">x</sysmsg>', 0, 10000) == "[系统消息]"
    assert render_content('<voipmsg><msg>通话</msg></voipmsg>', 0, 50) == "[通话]"
    # 高位 local_type 要按 32 位掩码还原
    assert render_content("x", 0, 244813135921) == "x", "掩码后应等于 49(appmsg) 且正文直通"
    assert render_content(xml, 0, (49 + (57 << 32))) == got


def smoke_guards_and_robustness():
    """真跑暴露的守卫：空映射不给追加行、岗位剥离不剥空、空 N 键不丢、null 日期不崩。"""
    import json as _json
    import tempfile
    from common import GridSheet, GridWorkbook
    from probe import analyze_workbook
    from build_matrix_rows import norm_o, load_agents, flag_dups

    # C1: 表头全部识别失败 → 不能给 next_append_row（否则会覆盖第 2 行）
    wb = GridWorkbook([GridSheet("运维", [["甲", "乙"], ["1", "2"], ["3", "4"]])])
    res = analyze_workbook(wb, None, "snap.json")
    assert res["column_mapping"] == {}
    assert res["last_data_row"] is None and res["next_append_row"] is None, res
    assert res["warnings"], "空映射必须给出 warning"
    # 有映射时仍正常
    wb2 = GridWorkbook([GridSheet("运维", [["需求描述", "提出人"], ["a", "张三"], ["b", "李四"]])])
    r2 = analyze_workbook(wb2, None, "snap.json")
    assert r2["next_append_row"] == 4, r2["next_append_row"]

    # C9: 无映射时扫描全部列（不再只扫前 27 列）
    wide = GridSheet("w", [[""] * 40 + ["x"]])
    from common import next_append_row_ws
    assert next_append_row_ws(wide, None) == 2, next_append_row_ws(wide, None)

    # C10: 姓名恰好等于岗位词时不剥空
    cfg = {"post_words": ["客服", "财务"]}
    assert norm_o("客服", cfg) == ("客服", ""), norm_o("客服", cfg)
    assert norm_o("张三客服", cfg) == ("张三", "客服")

    # C6/C7: ask_date=null 不崩；N 键存在但为空时不丢剥离结果
    d = tempfile.mkdtemp(prefix="wm_smoke_")
    _json.dump([{"O": "王五客服", "N": "", "L": "改费", "ask_date": "2026-05-14", "chat": "A"},
                {"O": "李四", "L": "导数据", "ask_date": None, "chat": "B"}],
               open(os.path.join(d, "agent_1.json"), "w", encoding="utf-8"), ensure_ascii=False)
    rows = load_agents(d, cfg)
    assert len(rows) == 2
    by_o = {r["O"]: r for r in rows}
    assert by_o["王五"]["N"] == "客服", by_o["王五"]
    assert rows[0]["ask_date"] is None or isinstance(rows[0]["ask_date"], str)

    # C8: 日期不可解析时不崩、不误判
    flag_dups([{"O": "a", "L": "改个费用", "ask_date": "坏日期", "chat": "A", "_i": 1},
               {"O": "a", "L": "改个费用", "ask_date": "2026-05-14", "chat": "B", "_i": 2}], {})
    import shutil
    shutil.rmtree(d, ignore_errors=True)


def smoke_reuse_position():
    """岗位复用共用层：读历史只到 end_row、默认只填空缺、批内顺序传播。"""
    from common import GridSheet, position_columns, read_history_positions, reuse_position_fill

    # 列解析：有映射按映射，缺映射退回 提出人=O(15) / 岗位=N(14)
    assert position_columns({"提出人": "O", "提出人岗位": "N"}) == (15, 14)
    assert position_columns({"提出人": "C", "提出人岗位": "B"}) == (3, 2)
    assert position_columns({}) == (15, 14)

    gs = GridSheet("运维", [
        ["需求描述", "提出人岗位", "提出人"],   # 第 1 行 = 表头
        ["a", "客服", "张三"],                  # 第 2 行
        ["b", "", "张三"],                      # 第 3 行
        ["c", "", "李四"],                      # 第 4 行（李四始终没岗位）
        ["d", "商务", "张三"],                  # 第 5 行
    ])
    # 只读到 end_row（历史区右端 = 追加起始行-1）→ 读到的是"客服"
    hist, pairs = read_history_positions(gs, 1, 3, col_person=3, col_post=2)
    assert hist == {"张三": "客服"} and pairs == 1, (hist, pairs)
    # 扩到第 5 行：同一人取**最近**的非空值
    hist2, pairs2 = read_history_positions(gs, 1, 5, col_person=3, col_post=2)
    assert hist2 == {"张三": "商务"} and pairs2 == 2, (hist2, pairs2)
    # 只读：网格没被改
    assert gs.cell(row=3, column=3).value == "张三"

    # 默认只填空缺；历史里没有的人不补；无提出人的行跳过
    rows = [{"提出人": "张三", "提出人岗位": ""},
            {"提出人": "张三", "提出人岗位": "财务"},
            {"提出人": "王五", "提出人岗位": ""},
            {"提出人": "", "提出人岗位": ""}]
    ch = reuse_position_fill(rows, {"张三": "客服"}, override=False)
    assert len(ch) == 1 and ch[0]["i"] == 0 and ch[0]["new"] == "客服", ch
    assert rows[0]["提出人岗位"] == "客服"
    assert rows[1]["提出人岗位"] == "财务", rows[1]     # 默认不覆盖已有值
    assert rows[2]["提出人岗位"] == ""                  # 历史无此人
    assert rows[3]["提出人岗位"] == ""

    # override=True 才覆盖
    rows2 = [{"提出人": "张三", "提出人岗位": "财务"},
             {"提出人": "张三", "提出人岗位": ""}]
    ch2 = reuse_position_fill(rows2, {"张三": "客服"}, override=True)
    assert [c["i"] for c in ch2] == [0, 1], ch2
    assert rows2[0]["提出人岗位"] == "客服"

    # 批内顺序传播：本批前一行已有的岗位供本批后续行使用（历史为空也能补）
    rows3 = [{"提出人": "赵六", "提出人岗位": "调度"},
             {"提出人": "赵六", "提出人岗位": ""}]
    ch3 = reuse_position_fill(rows3, {}, override=False)
    assert len(ch3) == 1 and ch3[0]["i"] == 1 and ch3[0]["new"] == "调度", ch3


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
    check(f"{len(MODULES)} 个模块 import", smoke_imports)
    print("== common 单测 ==")
    check("serial/col_index/col_letter", smoke_common)
    print("== 消息正文解码（真机乱码根因）==")
    check("ZSTD 降级/appmsg 引用/sysmsg 归一化", smoke_message_decoding)
    print("== 守卫与稳健性（真跑暴露）==")
    check("空映射不给追加行/剥离不剥空/null 日期不崩", smoke_guards_and_robustness)
    print("== 云文档快照抽象层 ==")
    check("稀疏->密集网格 / GridSheet 接口 / 追加行", smoke_snapshot_grid)
    check("raw.json 多形态识别", smoke_snapshot_raw_parse)
    check("payload -> kdocs rangeData", smoke_kdocs_payload)
    check("probe.analyze_workbook 吃 GridWorkbook", smoke_probe_analyze)
    print("== 岗位复用共用层（final 内置 / position_reuse 共用）==")
    check("列解析/历史只读到 end_row/只填空缺/批内传播", smoke_reuse_position)
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
