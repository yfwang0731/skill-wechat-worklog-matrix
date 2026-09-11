# -*- coding: utf-8 -*-
"""微信本地库解密包装器（基于 wcdb-key-tool）。

自动定位工具与 db_storage，先 extract 出密钥（需微信前台），再 decrypt 到指定输出目录。
密钥缓存按账号目录隔离（keys_<账号>.json），切换账号不会复用错密钥。
把"最脆弱、最容易忘参数"的解密步骤固化下来。

用法：
  python decrypt.py                         # 默认解密到 ./wechat_pilot/output/decrypted
  python decrypt.py --reextract             # 强制重新提取密钥（微信需在前台登录）
  python decrypt.py --db-storage <dir> --tool <py> --output <dir>

前置：wcdb-key-tool 不在 skill 内，需自备：
  git clone https://github.com/TANGandXUE/wcdb-key-tool scripts/tools/wcdb-key-tool
  或设置环境变量 WCDB_KEY_TOOL 指向 wcdb_key_tool_windows.py。
合规提示：属灰色工具，使用前向用户确认；只读本机、数据不出电脑。
"""
import os
import re
import sys
import subprocess
import argparse

from common import find_db_storage, find_wcdb_tool


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--db-storage", default=None, help="微信 db_storage 目录（默认自动识别）")
    ap.add_argument("--tool", default=None, help="wcdb_key_tool_windows.py 路径（默认自动/环境变量）")
    ap.add_argument("--output", default=None, help="解密输出目录（默认 ./wechat_pilot/output/decrypted）")
    ap.add_argument("--reextract", action="store_true", help="重新提取密钥（需微信前台登录）")
    a = ap.parse_args()

    tool = find_wcdb_tool(a.tool)
    if not tool:
        sys.exit(
            "✗ 未找到 wcdb-key-tool。\n"
            "  请执行: git clone https://github.com/TANGandXUE/wcdb-key-tool "
            f"{os.path.dirname(os.path.abspath(__file__))}/tools/wcdb-key-tool\n"
            "  或设置环境变量 WCDB_KEY_TOOL 指向 wcdb_key_tool_windows.py"
        )

    out = a.output or os.path.join(os.getcwd(), "wechat_pilot", "output", "decrypted")
    workdir = os.path.abspath(os.path.join(out, os.pardir))   # .../output
    os.makedirs(out, exist_ok=True)

    # 先定位 db_storage，据此把密钥缓存按账号隔离（多账号切换时避免复用错 keys）
    try:
        db = find_db_storage(a.db_storage)
    except FileNotFoundError as e:
        sys.exit(f"✗ {e}")
    # db = .../xwechat_files/<账号目录>/db_storage → 取上一级"账号目录"作为隔离 tag
    acct_dir = os.path.basename(os.path.dirname(os.path.normpath(db)))
    acct_tag = re.sub(r"[^A-Za-z0-9_-]", "_", acct_dir) or "default"
    keys_json = os.path.join(workdir, f"keys_{acct_tag}.json")

    # 1) 提取密钥（仅在缺失或强制时）
    #    必须显式传 --db-dir：一台机器可能有多个微信账户，只靠工具自动探测可能取错账号。
    if a.reextract or not os.path.exists(keys_json):
        print("[decrypt] 提取密钥中 —— 请确保微信已登录且处于前台（只读扫描内存，不读写聊天）…")
        try:
            subprocess.run([sys.executable, tool, "extract",
                            "--db-dir", db,
                            "--output", keys_json], check=True)
        except subprocess.CalledProcessError as e:
            sys.exit(f"✗ extract 失败（退出码 {e.returncode}）。确认微信在前台、未被锁屏，必要时加 --reextract 重试。")
    else:
        print(f"[decrypt] 复用已有密钥: {keys_json}（无需重登微信）")

    # 2) 解密所有分片
    print(f"[decrypt] 解密 {db} -> {out}")
    try:
        subprocess.run([sys.executable, tool, "decrypt",
                        "--db-dir", db, "--keys", keys_json, "--output", out],
                       check=True)
    except subprocess.CalledProcessError as e:
        sys.exit(f"✗ decrypt 失败（退出码 {e.returncode}）。检查 {os.path.basename(keys_json)} "
                 f"与 db_storage（{db}）是否属于同一账号；切换账号后请删除该 keys 文件并重试。")
    print(f"[decrypt] 完成。解密库位于: {out}")


if __name__ == "__main__":
    main()
