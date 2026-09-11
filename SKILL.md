---
name: wechat-worklog-matrix
description: >-
  通用「微信聊天记录 → 需求跟踪矩阵」流水线。扫描本机微信账户、由用户选择账户/会话/时间范围/处理人，
  解密导出聊天，LLM 识别对方提出的系统需求并判定答复结果，合并跨会话重复，按**表头**映射追加行。
  台账可落在本地 xlsx，也可落在 WPS 云文档（金山文档连接器：只给文档链接/名字即可读写，无需本地文件）。
  适用于任何"客户通过微信提需求、运维方需登记成台账/矩阵"的场景，不绑定具体客户或项目。
  触发场景：更新需求跟踪矩阵 / 从微信聊天记录整理需求台账 / 把微信里的需求登记进台账 /
  直接改 WPS·金山文档上的需求跟踪表（不用给本地文件）。
agent_created: true
---

# 微信记录 → 需求跟踪矩阵（通用流水线）

把微信 PC 聊天记录整理成需求跟踪矩阵台账。服务对象、项目、人员**全部由用户配置**，skill 本身不内置任何具体客户/项目/姓名。

## 何时使用
用户要"根据微信聊天记录更新/补充需求跟踪矩阵"、要读微信 4.x 本地记录、或要把微信里的需求登记进台账。
表格来源两种都给：本地 `.xlsx`，或**只报文档链接/名字、由 agent 直接去 WPS 云文档读写**（见「表格来源：两条通道」）。

## 环境依赖（首次使用前）
- Python 3.9+。
- `pip install zstandard` —— **强烈建议**。微信把长文本 / 图片 / 引用 / 合并转发 / 通话等消息以
  **ZSTD 压缩**存在 `message_content` 里（见「真机验证过的坑」第 5 条）。缺库不报错，但那些消息只能输出占位标记，
  **会漏需求**（本项目真机上 8 条压缩消息里有 4 条含需求正文）。
- 本地表格通道（`excel.source=local`）需 `pip install openpyxl`；回填走 `tencent-local-office-edit`。
  只支持 `.xlsx`/`.xlsm`，旧版 `.xls`/`.xlt` 会被 `probe` 拒绝（先「另存为 .xlsx」）。
- 云文档通道（`excel.source=kdocs`）需已连接**金山文档连接器**（`sheet.*` 工具），不需要 openpyxl。
- **全流程尽量用同一个解释器**，且确认它装了上面两个库 —— 换了没装 `zstandard` 的解释器不会报错，
  但会静默漏掉压缩消息里的需求（已踩）。本项目实测可用的隔离环境：`~/.workbuddy/binaries/python/envs/default`。
- 解密用第三方 wcdb-key-tool（见「关键事实」第 2 条），skill 内不含。
- 改完代码先自检：`python scripts/smoke_test.py --full`（import + 纯函数断言，不碰真实数据）。

## 核心原则（勿重蹈覆辙）
1. **不写死任何具体项目/人员/客户** —— 一切从 `config.json` 读。
2. **列位置靠"表头名"识别**，不写死 A~AA（换模板也能用）。
3. **账户/会话/时间范围由用户选**，不内置筛选条件。
4. **先探测 → 一次性确认 → 再执行**，不要逐个追问、不要替用户假设。
5. **单元格只写业务结果**，禁止把判定依据/分析过程写进任何单元格（尤其备注列）。
6. **表格 I/O 只走两条通道**（local / kdocs），算法不因通道而变；**绝不写死通道**——由 `config.excel.source` 决定。
7. **绝不静默改写已有数据**：拿不到"追加起始行"就报错退出，不兜底、不猜。

## 真机验证过的坑（必读，2026-09 实测）

> 下面每一条都是真跑一台 Windows + 微信 4.1 + WPS 云文档时踩出来的，代码已修，但了解原因能少走弯路。

1. **工作表参数名是 `worksheet_id`，不是 `sheetId`**。`sheet.get_sheets_info` / `get_range_data` /
   `update_range_data` 全用它。写错直接失败。
2. **`sheet.update_range_data` 单次 `rangeData` 最多 100 条**，超出报
   `rangeData length 110 exceeds limit 100`。→ `to_kdocs_payload.py` 已**自动分批**，用输出的 `calls` 逐批调用。
3. **op 数 ≈ 值单元格数，压不下去**：每条 `formula` op 只能写「一个值」，N 个不同值就是 N 条 op；
   能合并的只有 `format`/`merge`/`picture` 这类**区域**操作。所以 6 行 × 18 列 ≈ 107 条 + 日期格式几条、
   分 2 批调用是**正常量级**，不是脚本啰嗦。
4. **日期格式以"最新行"为准**：同一列老行可能是 `yyyy/m/d`、新行是 `yyyy-mm-dd`（真机实测同列两种并存）。
   取最新行的写法填 `excel.kdocs.date_numfmt`。日期底层存 **Excel 序列号**。
5. **消息正文可能是 ZSTD 压缩的**：`WCDB_CT_message_content` 决定 `message_content` 形态——
   `0`=明文 str，`4`=**ZSTD 压缩 bytes**（魔数 `28 B5 2F FD`）。早期版本把它当文本
   `errors="replace"` 解码 → 整段会话乱码、需求漏判。现已：能解压就解压，缺库/失败输出**明确标记**（绝不吐乱码），
   并在导出末尾统计未解码条数。
6. **`lt=49` 的消息，正文常藏在 `<refermsg>` 里**：对方用「引用回复」提需求时，
   `<title>` 只是"这个记得帮忙改"，真正的需求在引用原文 `<refermsg><content>`。
   只输出 `[链接/文件]` 会**整条丢掉需求**（真机上就漏过一条，最后靠问用户才补上）。
   现已提取 title + 引用正文 + des。
7. **`decrypt.py` 必须显式传 `--db-dir`**：一台机器可能有多个微信账户，工具自动探测可能取到停用的那个。
8. **群可能长期没消息**：真机上一个用户点名要看的群，最后一次发言在两周前。**先查活跃度再决定时间窗**，
   别默认"用户点名=近期有聊"；发现 0 条要如实报告并让用户决定是否扩窗。
9. **岗位复用默认只动新增行**：历史行只作岗位来源。早期版本默认从第 2 行扫整表，
   会回填历史里的空缺格（静默越界改写）。现在起始行缺省自动取追加起始行，推不出来就报错。
10. **云文档通道拿不到"便宜的历史岗位"**：实测四条路都不行 —— `download_file` 给的 URL 需登录态
    （直接 curl 得 `403 userNotLogin`，所以云文档**退化不成**"下载到本地走本地通道"）；
    `get_typed_value` **会跳过空单元格**（8 格只回 7 值），行列对不上；
    `get_range_data` 与 `read_file` 是同一套带字体/边框的格式（≈450 B/格，读 N+O 全列 ≈200 KB 上下文，
    且都无关掉样式的参数）；`find_range_data` 的 `filter` 结构未公开，传错会被**静默忽略并返回整段**（反而更贵）。
    → 云文档的岗位复用一律走 `--history` 手动喂入；能留空就留空（见「岗位复用」）。

## 表格来源：两条通道（`config.excel.source`）

用户可能给**本地文件**，也可能只说**"改 WPS 上那个文档"**。两条通道共用同一套列映射、判定规则与 payload 格式，
只有「读表 / 回填」两端的 I/O 不同。

| 环节 | `local`（本地 xlsx） | `kdocs`（WPS 云文档） |
|---|---|---|
| 读表头/末行 | `probe.py workbook --workbook <xlsx>`（openpyxl） | agent 取数 → `sheet_snapshot.py build` → `probe.py workbook --snapshot <json>` |
| 岗位复用 | **`final` 内置**（读 workbook 历史，零成本） | `final --history <json>`（快照不含数据区，见「岗位复用」） |
| 算追加起始行 | `build_matrix_rows.py final --workbook <xlsx>` | `build_matrix_rows.py final --snapshot <json>` |
| 回填 | payload → `tencent-local-office-edit` | payload → `to_kdocs_payload.py` → agent 调 `sheet.update_range_data` |
| 额外依赖 | openpyxl | 金山文档连接器 |

**为什么云文档要经过"快照"**：连接器的 MCP 工具**只有 agent 能调，Python 脚本调不到**。
所以读表必须由 agent 取数落盘、脚本再读快照；写表由脚本产出请求体、agent 去调工具。
**不要在脚本里试图直连连接器。**

### WPS 通道读表（分两趟，避免整表进上下文）
```bash
# 0) 定位文档：用户给链接最稳（从 URL 取 link_id → get_share_info → file_id/drive_id）；
#    只给文件名 → search_files 搜索（可能同名，务必让用户确认）；
#    再 sheet.get_sheets_info 取工作表清单、worksheet_id、已用区域 range.rowTo。
#    建议把 file_id / drive_id / worksheet_id 写进 config.excel.kdocs，后续不用重复搜。

# 1) 打印建议的读表调用体（含 file_id/worksheet_id/range 占位）
python scripts/sheet_snapshot.py plan --file-id <id> --worksheet-id <n> [--rows N] [--cols N]
#    第 1 趟：只读前 3 行 × 全部列（认表头）→ agent 把 sheet.get_range_data 的返回**原样**存 raw_hdr.json

# 2) 重建快照 → 解析表头列映射
python scripts/sheet_snapshot.py build --raw raw_hdr.json --out <output.dir>/sheet_snapshot.json
python scripts/probe.py workbook --snapshot <output.dir>/sheet_snapshot.json --json

# 3) 第 2 趟：按列映射只读关键列（需求描述/提出时间/提出人/提出人岗位/解决人…）→ raw_cols.json → 再 build
python scripts/sheet_snapshot.py build --raw raw_hdr.json --raw raw_cols.json --out <output.dir>/sheet_snapshot.json
python scripts/sheet_snapshot.py inspect --snapshot <output.dir>/sheet_snapshot.json   # 核对末行/追加行
```
> `get_range_data` 返回**稀疏**数组（只含有值的单元格，行列 0-based），`build` 会重建成密集网格并裁掉尾部/右侧空行列；
> 快照同时记 `num_formats`。表很长时用 `get_typed_value`（A1 记法）窄读更快更省。

### WPS 通道回填
```bash
python scripts/to_kdocs_payload.py --payload payload.json   --out kdocs_update.json
python scripts/to_kdocs_payload.py --payload payload_n.json --out kdocs_update_n.json
```
产出 `calls` 数组（每项即一次 `sheet.update_range_data` 的 arguments，已按 100 条上限分批），
并按 `excel.date_columns` 自动补日期 `format` op（`numfmt` 取 `excel.kdocs.date_numfmt`）。

**必须用 `sheet.update_range_data`（幂等、按坐标），不要用 `sheet.add_row`**（非幂等，重试/重复调用会插多行脏数据）。
写完后**必须用 `sheet.get_range_data` 回读同一区域核对**，不能只信 `code: 0`。

### 云文档通道专属坑
- **限频** `429001`/`429002`（熔断）：不要逐格写、不要密集重试；命中限频要等响应里给的恢复时间。
- **`.ksheet` 智能表格**：字段有类型（日期/单选等），写入形态与 `.xlsx` 不同，先确认文档类型。
- **区域保护**：文档若设了区域权限（`sheet.list_protection_ranges`），写入会失败，先让用户解除。
- **临时文件**：`raw_*.json` / `kdocs_update*.json` 流程结束后清理，别留在用户工作目录。

## 交互式流程

### 第 1 步：探测（agent 自动执行，一次跑完）
```bash
python pipeline.py probe [--workbook <xlsx> | --snapshot <快照json>] \
                         [--since <日期>] [--keyword <串>] [--dump <探测结果json>]
```
（`source=kdocs` 且未给 `--snapshot` 时自动找 `<output.dir>/sheet_snapshot.json`。）
一次探测出三件事并汇总打印：
- **① 本机微信账户列表**（可能有多个账号，让用户选）
- **② 表格表头 → 列映射**（未识别列、末数据行、追加起始行、处理人候选、warnings）
- **③ 可选会话清单**（名称/是否群/消息数/首末时间；需先有解密库，没有则先解密再重跑）

> 若 ② 报"未识别到任何列"，`next_append_row` 会是 `null` —— 这是**有意的**，禁止据此写入。

### 第 2 步：一次性确认（把探测结果一起呈现给用户）
用 AskUserQuestion（一次最多 4 问）或直接列出，一次确认/修正：
1. **用哪个微信账户**（探测①）
2. **表格来源**：本地 xlsx 路径 **或** WPS 云文档（链接/文件名）→ 决定 `config.excel.source`
3. **查哪些会话**（全部 / 按关键字 / 指定标识，来自探测③）
4. **时间范围**（起止日期）

另外需用户提供：**处理人姓名**（可由探测②的"处理人候选人"推断后请其确认）。
会话很多时把清单写成文件让用户勾选，不要逐条追问。
**若点名的群/会话 9/8 后 0 条消息，要如实报告并让用户决定是否扩窗**（见坑 8）。

### 第 3 步：写 config.json
从 `config.example.json` 复制，填入确认结果（account / excel / people / scope / defaults / rules）。

### 第 4 步：执行
```bash
python pipeline.py run                     # 解密 → 导出 → 分包（产物在 <output.dir>/transcripts/_out）

python scripts/build_matrix_rows.py preview --src <output.dir>/transcripts/_out --out <_out>/merged_preview.csv
#   ↑ 交用户裁决（删行/合并）后再进行下一步
python scripts/build_matrix_rows.py final --preview <_out>/merged_preview.csv \
       --remove "42,50" --merge "30:52" --start-row-excel <起始行> --out-dir <_out>
```
回填按 `config.excel.source` 分两路（见「表格来源：两条通道」）：
- `local` → 用 `tencent-local-office-edit` 把 payload 写回本地 xlsx。
- `kdocs` → `to_kdocs_payload.py` 生成 `calls` → agent 逐批 `sheet.update_range_data` → `get_range_data` 回读核对。

**派发判定子代理前**：按 config.json 渲染 `references/agent-prompt-zh.txt` 的**全部占位符**后使用——
`{{处理人}}`←people.handler、`{{我方标识}}`←people.my_identifiers、`{{服务对象}}`←people.service_object、
`{{业务系统描述}}`←excel.business_system_description、`{{对方关键字}}`←people.counterparty_keyword；
判定开关 `rules.*`（ignore_if_rejected / ignore_if_no_reply / data_change_default_done）按模板
【开关渲染表】注入（开=现文，关=替换为表中替代句）。**不要发未注入的裸模板。**

### 岗位复用（`rules.reuse_position_column`，默认开）
按行序向上找同一提出人**最近的岗位值**填补空缺，**只填空缺、不覆盖更具体值**（`--override` 才覆盖）。

**常规流程已内置到 `final`，一轮写入完成**：`build_matrix_rows.py final` 算出追加起始行后，
从**同一个表格来源**读历史区 `[表头行+1, 起始行-1]`（**只读**），把本批空岗位补齐，再写
`final_rows.csv` 与 `payload.json` —— 因此**不再需要**"先写一遍、再跑 `position_reuse` 补一遍"，
`final` 的 CSV 与你真正写进表格的内容**永远一致**（复用发生在写 CSV 之前）。

- 关闭：`--no-reuse-position`，或 `rules.reuse_position_column=false`（脚本自检后整步跳过）。
- **历史区读到 0 个「人+岗位」对 → 显式告警并跳过**（云文档快照只含表头行时的典型情形）。
- 只给 `--start-row-excel`、没给 `--workbook/--snapshot` → 同样**告警跳过**，不静默。
- 本地通道自动读历史（openpyxl，零额外成本）；**云文档通道默认不读**（代价见「真机验证过的坑」第 10 条）。

云文档需要补时，由 agent 取到历史后写成 json，用 `--history` 传入：
```bash
python scripts/build_matrix_rows.py final --preview merged_preview.csv --start-row-excel 192 \
       --history history_positions.json     # {"positions": {"人名": "岗位"}}
```

> **更省的做法**：像"备注无岗位、历史也没岗位"的首次提法人，表内复用本来救不了
> （岗位信息根本不在数据里）—— 直接在 `merged_preview.csv` 的「提出人岗位」列**裁决时补一个值**即可，
> 成本为零且不依赖历史。不要把"补满岗位"当成必须达成的目标。

**独立补跑**：`position_reuse.py` 保留用于事后补漏或自定义区间，与 `final` **共用同一实现**（`common.reuse_position_fill`）：
- 默认从**表格当前内容**取新行（此时新行须已落表，否则 0 条 + `[warn]`）。
- 加 `--new-rows <final_rows.csv>` 从 CSV 取新行（Excel 行号 = 起始行 + 序号），**不要求新行已落表**。
- `--start-row` 缺省自动取追加起始行；历史行永远只读。

## 判定规则（对方提出 → 我方答复）
> 生效方式：merge_cross_session / similarity_threshold / near_days 由 build_matrix_rows 代码读取；
> ignore_if_rejected / ignore_if_no_reply / data_change_default_done 是"判定提示词开关"，由执行代理按
> references/agent-prompt-zh.txt 的【开关渲染表】注入（不入 pipeline 脚本）；reuse_position_column 控制岗位复用。
> 把"忽略"类开关设为 false，会把被拒/无回复的需求也登记为待人工裁决行。
- 只记**对方提出的、针对系统的需求**；寒暄/通知/闲聊不算。
- 我方明确拒绝（不能做/做不了/不支持/你们自己改…）→ 忽略（`rules.ignore_if_rejected`）。
- 提出后无实质回复 → 忽略（`rules.ignore_if_no_reply`）。
- 明确答复完成（好了/已完成/改好了/搞定/你试试…）→ 记行；计划时间=完成时间=答复日期。
- **数据修改类特例**（改费用/改单号/录补调数据/导数据…）：没答完成也没拒绝 → 默认已完成，
  时间取提出日期（`rules.data_change_default_done`）。
- 同一需求跨会话出现 → 合并一行，提出人取最早提出者（`rules.merge_cross_session`）。
- 解决人/责任人 = `people.handler`。
- 提出时间取**首次提出日**（不是我方答复日）；一条消息含多个需求 → 拆多行。

## 脚本清单（`scripts/`）
| 脚本 | 作用 |
|---|---|
| `smoke_test.py` | 自检：import 全模块 + 纯函数断言（改完代码先跑它） |
| `probe.py` | 探测 `accounts` / `sessions` / `workbook`（表头→列映射、末行、处理人候选） |
| `sheet_snapshot.py` | 云文档通道：`plan`（读表调用体）/ `build`（稀疏→密集快照）/ `inspect` |
| `to_kdocs_payload.py` | 云文档通道：payload → `sheet.update_range_data` 的 `rangeData`（自动分批 + 日期格式） |
| `decrypt.py` | 封装 wcdb-key-tool 的 extract/decrypt（密钥按账号隔离，显式 `--db-dir`） |
| `export_conversations.py` | 按 config 导出转录；**含消息正文解码**（ZSTD 解压 / appmsg 引用提取 / 系统消息归一化） |
| `split_for_agents.py` | 转录按**文件大小均衡**分成 N 份，输出 `agent_N.txt` 清单 |
| `build_matrix_rows.py` | `preview`（去重标记+裁决）/ `final`（按列映射生成 payload） |
| `position_reuse.py` | 岗位列向上复用**独立补跑**（与 `final` 共用 `common.reuse_position_fill`；历史只读、只动新增行） |
| `common.py` | 配置加载、路径识别、日期/列工具、云文档快照抽象层（`GridWorkbook`/`GridSheet`/`sparse_to_grid`） |
| `pipeline.py` | `probe`（探测确认）/ `run`（执行） |

## 参数速查（与 argparse 一致）
| 命令 | 参数 |
|---|---|
| `pipeline.py probe` | `--workbook <xlsx>` \| `--snapshot <json>` `--since` `--keyword` `--dump` |
| `pipeline.py run` | `[--since] [--skip-decrypt] [--reextract]` |
| `probe.py` | `accounts --hint [--json]`；`sessions --decrypted <dir> [--since] [--keyword] [--json]`；`workbook (--workbook <xlsx> \| --snapshot <json>) [--sheet <关键字>] [--json]` |
| `sheet_snapshot.py` | `plan --file-id <id> --worksheet-id <n> [--rows] [--cols] [--letters "L,M,O"]`；`build --raw <f>… --out <json> [--sheet <名>] [--worksheet-id] [--file-id] [--drive-id] [--name]`；`inspect --snapshot <json>` |
| `build_matrix_rows.py` | `preview --src <_out> --out <csv> [--config]`；`final --preview <csv> [--remove "1,3"] [--merge "a:b"] [--out-dir] (--workbook <xlsx> \| --snapshot <json> \| --start-row-excel N) [--reuse-position \| --no-reuse-position] [--history <json>] [--config]` |
| `position_reuse.py` | `(--workbook <xlsx> \| --snapshot <json> \| --history <json>) --out <json> [--new-rows <final_rows.csv>] [--override] [--start-row N] [--end-row N] [--config]` |
| `to_kdocs_payload.py` | `--payload <json> [--out <json>] [--file-id] [--worksheet-id] [--date-cols "M,V,W"] [--date-numfmt yyyy-mm-dd] [--no-date-format] [--batch-size 100] [--config]` |

## 可配置项（见 `config.example.json`）
| 类别 | 变量 | 说明 |
|---|---|---|
| 账户 | `account.db_storage` / `decrypted`；`dir`（可选，仅人眼识别） | 多账号由用户选 |
| 表格来源 | `excel.source`（`local` \| `kdocs`） | 决定读表/回填走哪条通道 |
| 本地表格 | `excel.workbook` / `sheet_match` / `header_row` / `column_mapping` / `date_columns` / `start_row` | 列映射由 probe 依表头生成 |
| 云文档 | `excel.kdocs.link` / `file_id` / `drive_id` / `name` / `worksheet_id` / `date_numfmt` | 定位三选一：link（推荐）> file_id > name（搜索，可能同名）；`worksheet_id` 写入必填 |
| 人员 | `people.handler` / `my_identifiers` / `counterparty_keyword` / `service_object` | 处理人、我方标识、对方筛选关键字、服务对象 |
| 范围 | `scope.since` / `until` / `sessions` / `name_filter` / `groups` | 时间窗与会话清单 |
| 规则 | `rules.*` | 见「判定规则」；`reuse_position_column` 控制岗位复用 |
| 默认值 | `defaults.*` | 项目编号/名称/是否收费/需求类型/子系统/产生阶段/状态 等固定列填充 |
| 输出 | `output.dir` / `agents` | 产物目录、并行子代理路数 |
| 词表 | `post_words` | 岗位词表（从"提出人"文本剥离岗位）。**按客户补全**，漏词会剥离不出岗位 |
| 提示词上下文 | `excel.business_system_description` | 渲染进 `{{业务系统描述}}` |

## 关键事实（踩坑记录）
1. 微信 4.x 本地库 `~/Documents/xwechat_files/<账户目录>/db_storage/`，SQLCipher 加密；
   **一台机器可能有多个账户目录**，必须让用户选（活跃的那个可从 `db_storage` 下文件的 mtime 判断）。
2. 解密用第三方 **wcdb-key-tool**（不在 skill 内，灰色工具，只读本机，**使用前须经用户确认**）：
   `git clone https://github.com/TANGandXUE/wcdb-key-tool scripts/tools/wcdb-key-tool` 或设环境变量 `WCDB_KEY_TOOL`。
   口令缓存 `~/.wcdb-key-tool/wechat-passphrase.json` 有效时无需重登；缓存失效需 `extract`（**微信须前台**）。
3. 消息表 `Msg_<md5(wxid)>` 分布在 `message_1/2/3.db`（`_2` 最老 → `_3` 最新）；发送者 id 是**分片内**
   `Name2Id.rowid`（跨分片会变）；`create_time` 秒级；`local_type` 取低 32 位为类型
   （1=文字，3=图片，34=语音，43=视频，47=表情，49=链接/文件，50=通话）。
4. `WCDB_CT_message_content`：`0`=明文，`4`=ZSTD 压缩（见「真机验证过的坑」第 5 条）。
5. 日期写入用 **Excel 序列号**（1899-12-30 起），并对 `date_columns` 设日期格式。
6. 子表名**按关键字动态匹配**（如"运维"），不要写死年份。

## 易错点
- 会话很多时别用 AskUserQuestion 逐条问；探测后写清单交用户勾选，或按关键字过滤。
- 子代理判定出入较大：`merged_preview.csv` 必须交用户裁决后再回填。
- 子代理写出的 JSON 可能带 `null` / 空串字段；脚本已做兜底，但**裁决时也要扫一眼**。
- 本地通道："Export file is occupied" → save_file 到临时路径 → close_file(force) → 二进制覆写原文件；
  引擎保存会把 `.xls` 写成 xlsx，**保存后确认/改为 `.xlsx`**。
- **不要把连接器当脚本可用的库**：`sheet.*` 都是 agent 侧 MCP 工具，脚本只能读写 JSON 快照/请求体。
- 云文档同名文件多：用户只给文件名时**先搜索再让用户确认**，别猜。
- 详细坑位见 `references/workflow-notes.md`。
