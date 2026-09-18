# wechat-worklog-matrix

通用「微信 PC 聊天记录 → 需求跟踪矩阵」流水线（Weixin 4.x / WeChat 4.1+）。

把客户/用户在微信里提出的系统需求，自动整理登记成需求跟踪矩阵台账：
**探测 → 一次性确认 → 解密导出 → LLM 识别判定 → 人工裁决 → 回填台账**。

不绑定任何具体客户、项目或人员——服务对象、项目、会话、处理人全部由使用者通过 `config.json` 配置。

当前版本 [`v1.1.0`](https://github.com/yfwang0731/skill-wechat-worklog-matrix/releases/tag/v1.1.0)　·　变更历史见 [`CHANGELOG.md`](CHANGELOG.md)

[![selftest](https://github.com/yfwang0731/skill-wechat-worklog-matrix/actions/workflows/selftest.yml/badge.svg)](https://github.com/yfwang0731/skill-wechat-worklog-matrix/actions/workflows/selftest.yml)

## 特性

- **配置驱动，零内置假设**：账户 / 表格 / 会话 / 时间范围 / 处理人全部由你选择；仓库内不出现任何真实账号、姓名、客户与项目标识。
- **两种表格来源**（`excel.source`）：`local` 本地 `.xlsx`（openpyxl 读、editor_sdk 回填）；
  `kdocs` **WPS 云文档**（只用给文档链接/名字，agent 直接读写，不必把文件下到本地）。
  两条通道共用同一套列映射、判定规则与 payload 结构，只换 I/O 两端。
- **列位置靠表头识别**：`probe` 读目标表头自动生成「表头 → 列」映射，换模板也能用，不写死 A~AA。
- **消息正文解码到位**：微信把长文本/图片/引用/合并转发等以 ZSTD 压缩存储，
  装 `zstandard` 可完整还原（缺库时给出明确标记而非乱码）；`[链接/文件]` 会提取标题与**引用原文**，
  对方用「引用回复」提需求时不再丢内容。
- **绝不静默改写已有数据**：拿不到「追加起始行」就报错退出，不兜底、不猜。
- **人在回路**：先探测一次性确认 → 子代理判定结果必须交人工裁决（删行/合并）后才回填。
- **单元格只写业务结果**：备注等列一律不写判定依据/分析过程。
- **合规边界**：解密用第三方灰色工具 wcdb-key-tool（只读本机、数据不出电脑、使用前向用户确认）；
  聊天明文与探测中间产物全部留在本地并被 `.gitignore` 排除。

## 安装

本技能是 WorkBuddy 的 skill，装到技能目录即可被识别：

```bash
# 方式 A：clone 到用户级技能目录（所有项目都能用）
#   Windows 的技能目录：%USERPROFILE%\.workbuddy\skills\
git clone https://github.com/yfwang0731/skill-wechat-worklog-matrix.git \
  ~/.workbuddy/skills/wechat-worklog-matrix

# 方式 B：网页右上 Code → Download ZIP，解压到同一位置
```

- **只给单个项目用**：放到 `<项目根>/.workbuddy/skills/`。
- **目录名请用 `wechat-worklog-matrix`**（与 `SKILL.md` 里的 `name` 一致），
  不要保留 `skill-` 前缀、也不要多套一层目录。
- 装好后确认路径形如 `…/skills/wechat-worklog-matrix/SKILL.md`，
  然后在 WorkBuddy 里应能看到该技能。

更新已安装的技能：

```bash
git -C ~/.workbuddy/skills/wechat-worklog-matrix pull
```

## 快速开始

```bash
# 0) 环境依赖
pip install zstandard           # 建议：解压 ZSTD 压缩的消息正文（缺库会漏需求）
pip install openpyxl            # 仅本地表格通道需要；云文档通道不需要
#    本地通道：表格须为 .xlsx（.xlsm 也可）；旧版 .xls/.xlt 会被 probe 拒绝，请先另存为 .xlsx
#    云文档通道：先在连接器管理里连接「金山文档」，再按 SKILL.md「WPS 通道读表」取数生成快照

# 1) 探测（账户 / 表头映射，一次跑完供一次性确认）
python pipeline.py probe --workbook <你的需求矩阵.xlsx> --dump probe.json
python pipeline.py probe --snapshot <output.dir>/sheet_snapshot.json --dump probe.json   # 云文档通道
#    会话清单要等解密之后才有 → 先解密，再跑一次 probe（见 SKILL.md 第 1 步）

# 2) 一次性确认（把探测结果一起呈现给用户：账户 / 表格来源 / 会话 / 时间范围 / 处理人）

# 3) 写 config.json（复制 config.example.json 填写 account / excel / people / scope / defaults）
#    ⚠ excel.column_mapping 与 excel.header_row 务必用 probe 的 --json 输出粘进去，不要手填：
#      final 依赖它们拿列位，漏填会直接停在报错上。

# 4) 执行：解密 → 导出转录 → 分包 → 派发子代理识别 → 预览裁决 → 生成 payload → 回填
python pipeline.py run
python scripts/build_matrix_rows.py preview --src <output.dir>/transcripts/_out --out merged_preview.csv
#    预览交用户裁决（删行/合并）后再 final；起始行必须给来源之一
python scripts/build_matrix_rows.py final --preview merged_preview.csv --remove "..." --merge "a:b" \
        --start-row-excel <起始行>
#    final 已内置岗位复用 → 一轮写入；CSV 与真正写进表格的内容永远一致
#    本地通道：用 tencent-local-office-edit 把 payload.json 回填到目标子表
#    云文档通道：final 改用 --snapshot <json>（第 2 趟读表把行读全，岗位复用就自动生效；
#    只有快照确实读不到历史时才用 --history <json> 手写兜底）；再
python scripts/to_kdocs_payload.py --payload payload.json --out kdocs_update.json
#    然后 agent 按输出的 calls 逐批 sheet.update_range_data 写入，并 get_range_data 回读核对

# 可选·事后补跑（必须说明哪些行是本批新行，否则直接报错）
python scripts/position_reuse.py --workbook <xlsx> --new-rows final_rows.csv --out payload_n.json
```

解密工具 wcdb-key-tool（第三方，不在本仓库内）：

```bash
git clone https://github.com/TANGandXUE/wcdb-key-tool scripts/tools/wcdb-key-tool
# 或设置环境变量 WCDB_KEY_TOOL 指向 wcdb_key_tool_windows.py
```

## 实测与可靠性

2026-09 在一台 Windows + 微信 4.1 + WPS 云文档上完整实跑过，修复了一批「**没有报错、却把数据改错了**」
的路径（静默覆盖尾行、云文档静默改写纯文本值等）。这些约束已写进 [`SKILL.md`](SKILL.md) 的对应主题，
机制与逐条实测见 [`references/workflow-notes.md`](references/workflow-notes.md)，
历次独立评测的分数轨迹与**被推翻的结论**见 [`CHANGELOG.md`](CHANGELOG.md)。

## 目录结构

```
wechat-worklog-matrix/
├── SKILL.md                    # 技能主文档：30 秒速览 / 流程 / 表格通道 / 岗位复用 / 判定规则 / 速查表
├── CHANGELOG.md                # 完整开发过程（缘起 / 逐日变更 / 独立评测记录 / 设计不变量）
├── config.example.json         # 配置模板（复制为 config.json 后填写）
├── pipeline.py                 # 编排入口：probe（探测）/ run（解密→导出→分包）
├── test-prompts.json           # 典型使用场景（供技能质量评估用，运行流程时不读它）
├── LICENSE                     # MIT
├── .gitignore                  # 运行产物与本地配置的忽略清单（合规相关的唯一权威）
├── .gitattributes              # 仓库内一律 LF（让各平台检出逐字节一致，消除换行噪声）
├── .github/workflows/
│   └── selftest.yml            # CI：os × python × 依赖 共 8 格，跑 scripts/smoke_test.py
├── scripts/
│   ├── probe.py                #   探测账户/会话/表头映射（本地 --workbook / 云文档 --snapshot）
│   ├── sheet_snapshot.py       #   云文档通道：读表计划 / 稀疏返回→密集快照 / 快照核对
│   ├── to_kdocs_payload.py     #   云文档通道：payload → update_range_data 的 rangeData
│   ├── decrypt.py              #   封装 wcdb-key-tool 解密（密钥按账号隔离）
│   ├── export_conversations.py #   按 config 导出会话转录（含 ZSTD 解压 / 引用正文提取）
│   ├── split_for_agents.py     #   转录按文件大小均衡分 N 份给子代理
│   ├── build_matrix_rows.py    #   preview（契约校验+去重标记）/ final（生成 payload，含岗位复用）
│   ├── position_reuse.py       #   岗位复用独立补跑（与 final 共用 common.reuse_position_fill）
│   ├── rules_check.py          #   判定规则行为验证：plan（渲染提示词 + 盲评 case）/ verify（比对标注）
│   ├── common.py               #   配置加载 / 日期与列工具 / 快照抽象层 / 控制台编码 / 输入路径守卫
│   └── smoke_test.py           #   自检：import / 纯函数 / CLI 契约 / 路径与产出契约 / 快照覆盖 / 文档结构 / 端到端
└── references/
    ├── agent-prompt-zh.txt     # 需求识别子代理提示词模板
    ├── rules-fixtures.json     # 判定规则的行为验证夹具（带标注；改规则后跑 rules_check）
    └── workflow-notes.md       # 机制说明与实测记录（SKILL.md 的详情层）
```

## 自检

```bash
python scripts/smoke_test.py        # 31 项：import / 纯函数 / CLI 契约 / 输入路径守卫 / 子代理产出契约 / 快照覆盖与表内复用 / frontmatter 额度 / 文档分层与引用结构 / 离线端到端 / 仓库卫生 / 编码
python scripts/smoke_test.py --full # 再加 3 项依赖 openpyxl、zstandard 的检查
```

全程**不碰真实微信数据与网络**，可随时执行；失败以退出码 1 结束，适合直接挂 CI。
覆盖的重点是那些"**不会报错、却会改错数据**"的路径：缺关键输入时必须报错退出、列映射靠表头识别、
跨会话去重、岗位列只填空缺、云文档请求体分批与键名、以及 Windows 非 UTF-8 控制台下不能崩。

另外三条**新加的面**值得单说：
- **子代理产出契约**（`smoke_agent_contract`）：错值不会当场报错，而是顺着 `preview → final` 写进台账，
  所以形状/必填/值域/日期可解析四类**硬拦**，契约外字段名只告警，且**没有跳过开关**。
- **快照覆盖范围与表内复用**（`smoke_snapshot_coverage`）：`build` 会给每张子表记 `coverage`，
  `final` 用它判"行够不到历史区 → 报错 / 列没覆盖到岗位列 → 告警"；并有一条**反向断言**守住
  已作废的结论（"云文档不做表内复用"）不许复活。
- **文档本身也是被测对象**（`smoke_doc_refs` / `smoke_doc_structure` / `smoke_doc_claims` 三条）：
  引用与参数速查表的**指向**必须解析得到（章节名 / 文件 / markdown 链接 / 参数归属），
  结构必须自洽（标题层级不跳级、同一父节不重名、有序列表编号连续、README 目录树与脚本清单
  与磁盘逐项对齐、新加的 `.md` 必须纳入守卫），文档里**声称的事实**必须与仓库现状一致
  （术语声明在位、已作废的说法不许当现行、夹具条数与 CI 格数等关键计数与实物一致）；
  同一条链路还守住"`SKILL.md` 写规则 / `workflow-notes` 写依据"的分层。
  > 为什么值得：这类缺陷**不会报任何错** —— 改章节名忘了同步引用、加了脚本忘了同步清单、
  > 改了 CI 维度忘了改文档里的格数，全都会静默留下一个"读者点不到"或"照做跑不通"的坑。

CI（[`.github/workflows/selftest.yml`](.github/workflows/selftest.yml)）跑的就是它 —— 本地与 CI
**同一个入口**，避免"CI 绿、本地跑不到"或反之：

| 轴 | 取值 | 为什么 |
|---|---|---|
| OS | ubuntu / **windows** | Windows 的标准流编码是 cp1252，输出中文会崩，**只在 Windows 上复现** |
| Python | 3.9 / 3.13 | 3.9 是声明的最低支持版本，跨代测避免"你本地能跑、用户装不上库" |
| 依赖 | minimal / **full** | minimal 守"零依赖也能跑 + 缺库给可执行提示"；full 才跑得到本地通道读表与 ZSTD 解压 |

> CI 里**故意不设** `PYTHONUTF8` / `PYTHONIOENCODING` —— 设了就把 Windows 那个编码缺陷盖住，
> windows 格子等于白跑。脚本靠 `common.ensure_utf8_stdio()` 自己把标准流切到 UTF-8。

## 许可

本项目采用 MIT 许可，详见 [LICENSE](LICENSE)。

> 本 skill 依赖的解密工具 wcdb-key-tool 为第三方项目，不在本仓库内，其授权与合规性请自行评估。

## 合规与隐私说明

- 本 skill 不采集、不上传任何聊天数据；探测 / 导出 / 解密 / 回填的产物**全部留在本机**，
  且已被 `.gitignore` 排除。**完整的忽略清单以 [`.gitignore`](.gitignore) 为准**（那里也会随新功能更新），
  其中含聊天明文与个人标识的几类是：`wechat_pilot/`、`*.db`、`keys_*.json` / `all_keys*.json`、
  `config.json`、云文档中间物（`sheet_snapshot.json` / `raw_*.json` / `kdocs_update*.json` /
  `payload*.json`）、以及**会渲染出真实标识**的 `prompt_rendered.txt` / `rules_cases/` /
  `rules_check_out/` / `validation_report.txt`。
- 仓库内不包含任何真实账号目录、姓名、微信号、客户名或项目号。
- wcdb-key-tool 属第三方灰色工具：使用前需向用户确认，只读本机、数据不出电脑。
