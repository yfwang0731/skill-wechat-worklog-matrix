# wechat-worklog-matrix

通用「微信 PC 聊天记录 → 需求跟踪矩阵」流水线（Weixin 4.x / WeChat 4.1+）。

把客户/用户在微信里提出的系统需求，自动整理登记成需求跟踪矩阵台账：
**探测 → 一次性确认 → 解密导出 → LLM 识别判定 → 人工裁决 → 回填台账**。

不绑定任何具体客户、项目或人员——服务对象、项目、会话、处理人全部由使用者通过 `config.json` 配置。

## 特性

- **配置驱动，零内置假设**：账户 / 表格 / 会话 / 时间范围 / 处理人全部由你选择；skill 内不出现任何真实账号、姓名、客户与项目标识。
- **两种表格来源**（`excel.source`）：
  - `local` —— 本地 `.xlsx`，探测用 openpyxl、回填用 editor_sdk；
  - `kdocs` —— **WPS 云文档**，只需给文档链接/名字，agent 直接读写云文档（金山文档连接器），不用把文件下载到本地。
  两条通道共用同一套列映射、判定规则与 payload 结构，只换 I/O。
- **列位置靠表头识别**：`probe` 读取目标表格表头自动生成「表头 → 列」映射，换模板也能用，不写死 A~AA。
- **消息正文解码到位**：微信把长文本/图片/引用/合并转发等以 ZSTD 压缩存在 `message_content`，
  装 `zstandard` 可完整还原（缺库时给出明确标记而非乱码）；`[链接/文件]` 会提取 `<title>` 与
  **引用原文**，对方用「引用回复」提需求时不再丢内容。
- **绝不静默改写已有数据**：拿不到「追加起始行」就报错退出，不兜底、不猜（真机踩出过三条这类路径）。
- **人在回路**：先探测一次性确认 → 子代理判定结果必须交人工裁决（删行/合并）后才回填。
- **单元格只写业务结果**：备注等列一律不写判定依据/分析过程。
- **合规边界**：解密使用第三方灰色工具 wcdb-key-tool（只读本机、数据不出电脑、使用前向用户确认）；聊天明文与探测中间产物全部留在本地并被 `.gitignore` 排除。

## 快速开始

```bash
# 0) 环境依赖
pip install zstandard           # 建议：解压 ZSTD 压缩的消息正文（缺库会漏需求）
pip install openpyxl            # 仅本地表格通道需要；云文档通道不需要
#    本地通道：表格须为 .xlsx（.xlsm 也支持）；旧版 .xls/.xlt 会被 probe 拒绝，请先另存为 .xlsx
#    云文档通道：在 WorkBuddy 连接器管理里连接「金山文档」

# 1) 探测（账户 / 会话 / 表头映射，一次跑完供一次性确认）
python pipeline.py probe --workbook <你的需求矩阵.xlsx> --dump probe.json
#    云文档通道：先按 SKILL.md「WPS 通道读表」取数生成快照，然后
python pipeline.py probe --snapshot <output.dir>/sheet_snapshot.json --dump probe.json

# 2) 一次性确认（把探测结果一起呈现给用户确认：账户 / 表格来源 / 会话 / 时间范围 / 处理人）

# 3) 写 config.json（复制 config.example.json 填写 account / excel / people / scope / defaults）

# 4) 执行：解密 → 导出会话转录 → 分包 → 派发子代理 LLM 识别 → 预览裁决 → 生成 payload → 回填
python pipeline.py run
python scripts/build_matrix_rows.py preview --src <output.dir>/transcripts/_out --out merged_preview.csv
#    预览交用户裁决（删行/合并）后再 final；起始行必须给来源之一（workbook/snapshot/start-row-excel）
python scripts/build_matrix_rows.py final --preview merged_preview.csv --remove "..." --merge "a:b" \
        --start-row-excel <起始行>
#    final 已**内置岗位复用**：从同一个表格来源读历史区（只读）补齐本批空岗位 → **一轮写入**
#    （<output.dir>/transcripts/_out 即 run 产物目录，output.dir 默认 ./wechat_pilot）
#    （派发判定子代理前，按 config 的 rules.*/people.* 渲染 references/agent-prompt-zh.txt，勿发裸模板）
#    本地通道：用 tencent-local-office-edit 把 payload.json 回填到目标子表
#    云文档通道：final 改用 --snapshot <json>；需要补历史岗位时再加 --history <json>（见 SKILL.md）；再
python scripts/to_kdocs_payload.py --payload payload.json --out kdocs_update.json
#    然后 agent 逐批调 sheet.update_range_data（calls 数组），并 sheet.get_range_data 回读核对
# 可选·事后补跑（**必须**说明哪些行是本批新行，否则直接报错）：
python scripts/position_reuse.py --workbook <xlsx> --new-rows final_rows.csv --out payload_n.json
#    或 --start-row <本批首行号>（会扫该行之后的全部行，含历史）
```

解密工具 wcdb-key-tool（第三方，不在本仓库内）：

```bash
git clone https://github.com/TANGandXUE/wcdb-key-tool scripts/tools/wcdb-key-tool
# 或设置环境变量 WCDB_KEY_TOOL 指向 wcdb_key_tool_windows.py
```

## 两条表格通道对照

| 环节 | `local`（本地 xlsx） | `kdocs`（WPS 云文档） |
|---|---|---|
| 读表头/末行 | `probe.py workbook --workbook <xlsx>` | agent 调连接器取数 → `sheet_snapshot.py build` → `probe.py workbook --snapshot <json>` |
| 岗位复用 | **`final` 内置**（读 workbook 历史，零成本） | `final --history <json>`（云文档读历史太贵，见坑 10） |
| 算追加起始行 | `build_matrix_rows.py final --workbook <xlsx>` | `build_matrix_rows.py final --snapshot <json>` |
| 回填 | payload → `tencent-local-office-edit` | payload → `to_kdocs_payload.py` → `sheet.update_range_data` |
| 额外依赖 | openpyxl | 金山文档连接器 |

> 云文档写入**用 `sheet.update_range_data`（幂等），不要用 `sheet.add_row`**（后者非幂等，重试会插多行脏数据）；写后必须回读核对。

## 真机验证过的坑（摘要）

2026-09 在一台 Windows + 微信 4.1 + WPS 云文档上实跑，暴露并已修的问题（详见 `SKILL.md` 与
`references/workflow-notes.md`）：

| # | 坑 | 处理 |
|---|---|---|
| 1 | 工作表参数名是 `worksheet_id` 而非 `sheetId` | 全量统一，并写入 smoke 断言防回归 |
| 2 | `update_range_data` 单次 `rangeData` **上限 100 条** | `to_kdocs_payload.py` 自动分批，输出 `calls` |
| 3 | 每条 `formula` op 只能写一个值 → op 数≈值单元格数 | 属接口设计，文档说明是正常量级，别想着合并收敛 |
| 4 | 日期格式同列老行 `yyyy/m/d`、新行 `yyyy-mm-dd` 并存 | **以最新行为准**，写 `excel.kdocs.date_numfmt` |
| 5 | 消息正文是 ZSTD 压缩，早期当文本解码 → 整段乱码 | 解压 + 缺库给明确标记；不再吐乱码 |
| 6 | `lt=49` 的引用回复，正文在 `<refermsg>` 里 | 提取 title + 引用正文，不再只写 `[链接/文件]` |
| 7 | `decrypt.py` 未传 `--db-dir`，多账户可能取错账号 | 显式传 `--db-dir` |
| 8 | 用户点名的群 9/8 后 0 条消息（两周前才发言） | 先查活跃度，如实报告并让用户决定是否扩窗 |
| 9 | 岗位复用默认扫整表 → 静默回填历史空缺格 | 默认只动新增行，历史只读；显式 `--start-row` 才越界 |
| 10 | 云文档**没有便宜的历史读法**：`download_file` 需登录态（403）、`get_typed_value` 跳空单元格、`get_range_data`/`read_file` 每格带样式（≈450 B）、`find_range_data` 的 `filter` 未公开 | 云文档岗位复用改走 `--history`；拿不到就留空，别硬读整列 |
| 11 | 岗位复用原需"先写一遍再补一遍"，两次写入之间行号可能错位 | **并进 `final`，一轮写入**；CSV 与写入内容永远一致 |
| 12 | `position_reuse` 改成"自动取追加起始行"后**永远 0 条**（表格尾部还有空的带格式行时连提示都不打）= 静默无操作 | 默认模式**报错退出**，必须给 `--new-rows <final_rows.csv>` 或 `--start-row` |
| 13 | 值一律以 `opType=formula` 写入 → `0012` 前导零 / `=A1` 等可能被表格引擎改写 | `to_kdocs_payload` **显式告警**并列出命中值（不改写值）；需原样保留时先把目标列设为「文本」格式 |

**共性教训**：其中三条（空表头映射、`final` 起始行兜底为 2、`position_reuse` 默认扫整表）都会在
**没有告警**的情况下改动「不该动的单元格」。现在的原则是 **拿不到追加起始行就报错退出**。

## 目录结构

```
wechat-worklog-matrix/
├── SKILL.md                    # 技能主文档：流程/判定规则/关键事实/易错点
├── config.example.json         # 配置模板（复制为 config.json 后填写）
├── pipeline.py                 # 编排入口：probe（探测）/ run（解密→导出→分包）
├── scripts/
│   ├── probe.py                #   探测账户/会话/表头映射（本地 --workbook / 云文档 --snapshot，支持 --json）
│   ├── sheet_snapshot.py       #   云文档通道：读表计划 / 稀疏返回→密集网格快照 / 快照核对
│   ├── to_kdocs_payload.py     #   云文档通道：payload → sheet.update_range_data 的 rangeData
│   ├── decrypt.py              #   封装 wcdb-key-tool 解密（密钥按账号隔离）
│   ├── export_conversations.py #   按 config 导出指定会话转录
│   ├── split_for_agents.py     #   转录按文件大小均衡分 N 份给子代理
│   ├── build_matrix_rows.py    #   preview（去重标记+裁决）/ final（生成 payload，**含岗位复用**）
│   ├── position_reuse.py       #   岗位复用**独立补跑**（与 final 共用 common.reuse_position_fill）
│   ├── common.py               #   配置加载 / 路径识别 / 日期与列工具 / 云文档快照抽象层
│   └── smoke_test.py           #   自检：import + 纯函数断言
└── references/
    ├── agent-prompt-zh.txt     # 需求识别子代理提示词模板
    └── workflow-notes.md       # 工作流要点与坑位备忘
```

## 自检

修改代码后运行：

```bash
python scripts/smoke_test.py        # import 全模块 + 纯函数断言（含快照抽象层 / kdocs payload 转换）
python scripts/smoke_test.py --full # 额外跑附加检查
```

## 许可

本项目采用 MIT 许可，详见 [LICENSE](LICENSE)。

> 注意：本 skill 依赖的解密工具 wcdb-key-tool 为第三方项目，不在本仓库内，其授权与合规性请自行评估。

## 合规与隐私说明

- 本 skill 不采集、不上传任何聊天数据；探测/导出/解密产物均在本机，且已被 `.gitignore` 排除（`wechat_pilot/`、`*.db`、`keys_*.json`、`config.json`、`sheet_snapshot.json`、`raw_*.json`、`kdocs_update*.json`、`payload*.json`）。
- 仓库内不包含任何真实账号目录、姓名、微信号、客户名或项目号。
- wcdb-key-tool 属第三方灰色工具：使用前需向用户确认，只读本机、数据不出电脑。
