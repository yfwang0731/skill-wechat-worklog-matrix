# -*- coding: utf-8 -*-
"""把转录文件清单均分成 N 份，写出 agent_1.txt .. agent_N.txt，供主代理交给并行子代理。

消除"手动把几十个转录文件分成 6 份"的摩擦：主代理读这些清单，把对应文件绝对路径
填入 references/agent-prompt-zh.txt 模板，派发给 N 个并行 agent 即可。

用法：
  python split_for_agents.py --transcripts <dir> --n 6 --out <dir>
默认 transcripts=./wechat_pilot/transcripts，out=./wechat_pilot/transcripts/_out，n=6。
"""
import os
import argparse


def file_weight(path):
    """以字节数近似文件工作量（中文 UTF-8 每字约 3 字节）。"""
    try:
        return os.path.getsize(path)
    except OSError:
        return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--transcripts", default=None)
    ap.add_argument("--out", default=None)
    ap.add_argument("--n", type=int, default=6, help="并行路数")
    a = ap.parse_args()

    tx = a.transcripts or os.path.join(os.getcwd(), "wechat_pilot", "transcripts")
    out = a.out or os.path.join(tx, "_out")
    os.makedirs(out, exist_ok=True)

    files = sorted(
        os.path.join(tx, f) for f in os.listdir(tx)
        if f.endswith(".txt") and not f.startswith("#")
    )
    if not files:
        print(f"✗ {tx} 下没有 .txt 转录文件，请先跑 export_conversations.py")
        return

    n = max(1, a.n)
    # 按文件大小降序贪心分配到当前最轻的桶（比轮询更均衡：避免一个大文件压垮单个 agent）
    buckets = [[] for _ in range(n)]
    loads = [0] * n
    for f in sorted(files, key=file_weight, reverse=True):
        lightest = min(range(n), key=lambda k: loads[k])
        buckets[lightest].append(f)
        loads[lightest] += file_weight(f)

    manifest = []
    for i, b in enumerate(buckets, 1):
        lst = os.path.join(out, f"agent_{i}.txt")
        with open(lst, "w", encoding="utf-8") as fh:
            fh.write("\n".join(os.path.abspath(x) for x in b))
        manifest.append((i, len(b), lst))
        print(f"agent_{i}: {len(b)} 个文件 / 约 {loads[i-1]//1024} KB -> {lst}")

    print(f"\n共 {len(files)} 个转录文件，按大小均衡分成 {n} 路。"
          f"把 references/agent-prompt-zh.txt 模板 + 各 agent_N.txt 清单交给并行子代理。")


if __name__ == "__main__":
    main()
