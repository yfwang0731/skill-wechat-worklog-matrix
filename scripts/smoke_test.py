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


def smoke_common():
    from common import serial, col_index, col_letter
    assert serial("2026-05-14") == 46156
    assert serial("2026/5/14") == 46156
    assert serial("2026年5月14日") == 46156
    assert serial("2026-05-14 10:00") == 46156
    assert serial("bad") == ""
    assert col_index("A") == 0 and col_index("AA") == 26
    assert col_letter(0) == "A" and col_letter(26) == "AA"


def smoke_build():
    from build_matrix_rows import norm_o, ids, QMAP
    cfg = {"people": {"counterparty_keyword": "李四"}, "post_words": ["财务", "客服"]}
    assert norm_o("李四财务", cfg) == ("李四", "财务"), norm_o("李四财务", cfg)
    cfg2 = {"people": {"counterparty_keyword": "某公司"}, "post_words": ["客服"]}
    assert norm_o("某公司 王五客服", cfg2) == ("王五", "客服")
    assert "ANTSQCY250789019" in ids("单号 ANTSQCY250789019 改费")
    assert "20260514" not in ids("20260514 提了需求")
    assert QMAP  # 非空


def smoke_n_reuse():
    import n_reuse_plan      # noqa: F401  (需 openpyxl)


def main():
    full = "--full" in sys.argv
    print("== import 冒烟 ==")
    check("8 个模块 import", smoke_imports)
    print("== common 单测 ==")
    check("serial/col_index/col_letter", smoke_common)
    print("== build_matrix_rows 单测 ==")
    check("norm_o/ids/QMAP", smoke_build)
    if full:
        print("== n_reuse_plan（需 openpyxl）==")
        check("import n_reuse_plan", smoke_n_reuse)
    else:
        print("== n_reuse_plan（openpyxl）== 加 --full 启用")

    print()
    if FAIL:
        print(f"✗ 冒烟失败 {len(FAIL)} 项：{FAIL}")
        sys.exit(1)
    print("✓ 全部通过")


if __name__ == "__main__":
    main()
