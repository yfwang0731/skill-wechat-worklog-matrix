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

把微信 PC 聊天记录整理成需求跟踪矩阵台账。服务对象、项目、人员全部由用户配置，skill 内不含任何具体客户/项目/姓名。

行文约定：`##` 是主题，`###` 是子主题；列表项一律以 `**标签**：` 开头自解释；`>` 只用于警告与提醒，不放规则正文。
代码块里沿用 `**加粗**` 写法表示强调 —— **代码块不渲染 markdown，星号会原样显示，这是有意的，别删**。

## 30 秒速览

**做什么**：把微信里**对方（客户）**提的系统需求，整理成台账的一批**新行**，追加到需求跟踪矩阵。历史行永远只读。

**先认通道**（`config.excel.source`，其余环节两条通道**算法完全一致**）：

| 用户给的 | `source` | 台账在哪 | 读 / 写 |
|---|---|---|---|
| 本地 `.xlsx` 路径 | `local` | 本地文件 | openpyxl 读 / `tencent-local-office-edit` 写 |
| WPS 链接或文件名（或只说"改 WPS 上那个表"） | `kdocs` | 云文档 | agent 取数→快照 / `to_kdocs_payload`→`update_range_data` |

> **术语等价**：`kdocs` = WPS 云文档 = 金山文档 = 那个连接器，**同一个东西的四种叫法**；
> `local` 就是本地 `.xlsx`。"通道"与"来源"也是同一个意思（都指 `config.excel.source`）。
> 下文的「读表」「回填」指两条通道各自那一端的 I/O，与本地/云文档无关。

**主链五步**（每步的输入输出与命令见「交互式流程」对应小节）：

```
① probe        两趟：先探「账户 + 表头映射」→ **解密**（会话清单只存在于解密库里）→ 再探一次拿会话清单
② 一次性确认    账户 / 台账来源 / 会话 / 时间范围 / 处理人   ← **只在这里问用户**，不逐条追问
③ 写 config    从 config.example.json 复制；column_mapping 直接粘 ① 的输出，不手填
④ run → preview → final   导出 → 分包派子代理判定 → **交用户裁决删行/合并** → 生成 payload
                           （① 已解密过就加 `--skip-decrypt` 复用，别重复解密）
⑤ 回填         local：editor_sdk；kdocs：to_kdocs_payload → update_range_data → **回读核对**
```

> **① 为什么要两趟**：②要拿会话清单给用户选，而会话清单**只存在于解密库里** ——
> 所以解密必须发生在 ① 和 ② **之间**，而不是等到 ④。顺序写错会卡在"没有会话可确认"。

**三条不许破的规矩**（历史上都真栽过；第 1、3 条见「核心原则」#7 / #5，第 2 条见「交互式流程」第 4 步）：

1. **绝不静默**：拿不到"追加起始行"、子代理产出不合契约、快照读不到数据区 —— 一律**报错退出**，不猜、不兜底、不降级。
2. **人工裁决必须在回填之前**：`preview` 必须给用户删行/合并，不许跳过直接写台账。
3. **单元格只放业务结果**：备注等列不写判定依据/分析过程。

**卡住了先看哪儿**：

| 症状 | 去哪节 |
|---|---|
| 不知道要问用户什么 | 第 2 步「一次性确认」 |
| 表格读不到 / 末行算错 / 岗位列空着 | 「WPS 通道读表」→「岗位复用」 |
| 该不该记这一条、记成什么 | 「判定规则」（含 Q 列取值域） |
| 写回失败 / 值被引擎改了 | 「云文档接口约束」；纯文本风险见 `references/workflow-notes.md`「云文档纯文本风险」 |
| 命令和参数记不住 | 末尾「速查」 |

> **首次使用前先看「环境依赖」**：解密要第三方 `wcdb-key-tool`（灰色工具，只读本机、数据不出机器），
> **须用户明确同意**后才用。

## 何时使用

- 要根据微信聊天记录**更新/补充需求跟踪矩阵**。
- 要读微信 4.x 的本地聊天记录。
- 要把微信里的需求**登记进台账**。
- 用户只说「改 WPS 上那个文档」——不用给本地文件也能干（见「表格来源：两条通道」）。

## 环境依赖（首次使用前）

- **Python 3.9+**，且**全流程用同一个解释器**。换了没装库的解释器不会报错，但会静默漏内容。
- **`pip install zstandard` —— 强烈建议**。微信把长文本/图片/引用/合并转发/通话等消息以 ZSTD 压缩存储，
  缺库只会输出占位标记，**会漏需求**（真机 8 条压缩消息里 4 条含需求正文）。
  本项目实测可用的隔离环境：`~/.workbuddy/binaries/python/envs/default`。
- **`pip install openpyxl`** —— 仅**本地表格通道**需要；回填另需 `tencent-local-office-edit`。
  只支持 `.xlsx`/`.xlsm`，旧版 `.xls`/`.xlt` 会被 `probe` 拒绝（先「另存为 .xlsx」）。
  缺库时会给出 `pip install openpyxl` 的明确提示，不会抛裸 `ImportError`。
- **金山文档连接器** —— 仅**云文档通道**需要，不需要 openpyxl。
- **wcdb-key-tool** —— 解密用，**第三方且不在 skill 内**。它是灰色工具（只读本机、数据不出电脑），
  **使用前须经用户确认**：

  ```bash
  git clone https://github.com/TANGandXUE/wcdb-key-tool scripts/tools/wcdb-key-tool
  # 或设环境变量 WCDB_KEY_TOOL 指向 wcdb_key_tool_windows.py
  ```

  口令缓存 `~/.wcdb-key-tool/wechat-passphrase.json` 有效时无需重登；失效需重跑 `extract`（**微信须前台**）。
  **一台机器可能有多个微信账户**，所以调 `decrypt.py` 时必须显式给 `--db-storage <目录>` ——
  脚本内部会给 wcdb-key-tool 传 `--db-dir`，但**那是第三方工具的参数，不是 `decrypt.py` 的**，
  直接写给 `decrypt.py` 会被 argparse 拒绝。
- **`python scripts/smoke_test.py --full`** —— 改完代码先跑它。不碰真实数据与网络，
  覆盖纯函数、CLI 守卫契约、离线端到端链路、仓库卫生与编码，失败以退出码 1 结束；
  CI（`.github/workflows/selftest.yml`，os × python × 依赖 共 8 格）跑的就是它。
- **改代码的两条硬约定**（CI 会拦）：
  ① 新增 / 修改 `__main__` 入口时，必须在**任何输出之前**调用 `common.ensure_utf8_stdio()` ——
  Windows 的标准流编码是 cp1252，中文输出会崩，而**本地是 UTF-8 测不出来**；
  ② 自检只许把临时产物写进系统临时目录 —— CI 有一格专门校验跑完 `git status` 仍然干净。

## 核心原则（勿重蹈覆辙）

1. **不写死任何具体项目/人员/客户** —— 一切从 `config.json` 读。
2. **列位置靠表头名识别**，不写死 A~AA（换模板也能用）。
3. **账户/会话/时间范围由用户选**，不内置筛选条件。
4. **先探测 → 一次性确认 → 再执行**，不要逐个追问、不要替用户假设。
5. **单元格只写业务结果**，禁止把判定依据/分析过程写进任何单元格（尤其备注列）。
6. **表格 I/O 只走两条通道**（local / kdocs），算法不因通道而变；**绝不写死通道**——由 `config.excel.source` 决定。
7. **绝不静默改写已有数据**：拿不到「追加起始行」就报错退出，不兜底、不猜。
8. **只记可交付的需求**：能落成"改什么 / 修什么 / 加什么"的才叫需求。笼统抱怨与体感反馈
   （"系统特别卡"）且我方未明确答复已解决的，不记 —— 台账是交付物，不是意见箱。

## 交互式流程

### 第 1 步：探测（agent 自动执行；**分两趟**——会话清单要等解密之后）

```bash
# 1a) 账户 + 表头映射（不需要解密库）
python pipeline.py probe [--workbook <xlsx> | --snapshot <快照json>] [--dump <探测结果json>]

# 1b) 先解密（见第 4 步的 run，或单独跑 decrypt.py），**再**探一次拿会话清单
python pipeline.py probe [--workbook <xlsx> | --snapshot <快照json>] --dump <探测结果json>
```

一次探测出三件事并汇总打印：

- **① 本机微信账户列表**：可能有多个账号，让用户选。
- **② 表格表头 → 列映射**：未识别列、末数据行、追加起始行、处理人候选、warnings。
- **③ 可选会话清单**：名称/是否群/消息数/首末时间。**需要先有解密库**，没有则 ③ 为空 →
  先解密、再重跑一次 probe（这就是上面 1a/1b 两趟的原因；`source=kdocs` 时未给 `--snapshot`
  会自动找 `<output.dir>/sheet_snapshot.json`）。
  - **解密库在哪儿**由 `account.decrypted` 决定；**没给 `--config` 时找默认位置
    `./wechat_pilot/output/decrypted`**（脚本会把实际用的路径打在 ③ 那行里）。所以 1b 要么
    给 `--config`（**写在子命令之前**），要么让解密输出落在默认位置 —— 否则 ③ 会一直空着，
    而你会误以为"这个账户没有会话"。
  - **`decrypt.py --output` 的默认值就是那个默认位置**（`./wechat_pilot/output/decrypted`），
    所以不指定 `--output` 时两边天然对齐；**改过输出路径就要同步 `account.decrypted`**，
    否则 probe 找不到解密库、③ 依旧是空的。
  - 只想拿会话清单、不想跑整条导出链时，**单独跑 `decrypt.py` 即可**（`pipeline.py run`
    会顺带解密，但那会连导出与分包一起做掉）。

两条必须知道的探测行为：

- **表头行**取「前 5 行里**命中内置表头别名数最多**」的一行，非空单元格数只作次级判据。
  只按"非空最多"在很宽的表上会被数据行骗过（数据行往往更长更满），现在不会了。
- **若②报"未识别到任何列"**，`next_append_row` 会是 `null`——这是**有意的**，禁止据此写入。
  先查三件事（按可能性排序）：① **子表选错了** → 用 `--sheet <关键字>` 点名；② **表头行不在前 5 行**
  （探测只在前 5 行里找命中别名最多的一行，换过模板/上面加了标题行就会跑到第 6 行开外）；
  ③ 表头名与内置别名 `COLUMN_ALIASES` 对不上 → 看 `probe` 打出的**未识别清单**，先统一表头。
  这一条与「列映射」里那句"空映射必须拒绝给追加行号"是同一个坑的两端：**判据空 → 起点算错 → 覆盖已有数据**。

### 第 2 步：一次性确认（把探测结果一起呈现给用户）

用 AskUserQuestion（一次最多 4 问）或直接列出，一次确认/修正：

1. **用哪个微信账户**（探测①）
2. **表格来源**：本地 xlsx 路径 **或** WPS 云文档（链接/文件名）→ 决定 `config.excel.source`
3. **查哪些会话**（全部 / 按关键字 / 指定标识，来自探测③）
4. **时间范围**（起止日期）

另需用户提供**处理人姓名**（可由探测②的"处理人候选人"推断后请其确认）。

- 会话很多时把清单写成文件让用户勾选，**不要用 `AskUserQuestion` 逐条追问**（一次最多 4 问，
  别把一次确认问成十几轮）。
- **被点名的群/会话在时间窗内 0 条消息时，先查它的活跃度再定时间窗** —— 真机上一个点名的群
  最后一次发言在两周前。别默认"用户点名=近期有聊"；确认 0 条后**如实报告**，让用户决定是否扩窗。

### 第 3 步：写 config.json

从 `config.example.json` 复制，填入确认结果（account / excel / people / scope / defaults / rules）。

**`excel.column_mapping` 与 `excel.header_row` 不能手填** —— 把第 1 步探测②的输出直接粘过去：

```bash
python probe.py workbook --workbook <xlsx> --json        # 或 --snapshot <快照json> --json
```

这两个键是 `build_matrix_rows final` 的**必需输入**（缺 `column_mapping` 会拿不到列位、无法生成 payload），
漏填是最常见的卡点。

### 第 4 步：执行

```bash
python pipeline.py run                     # 解密 → 导出 → 分包（产物在 <output.dir>/transcripts/_out）

python scripts/build_matrix_rows.py preview --src <output.dir>/transcripts/_out --out <_out>/merged_preview.csv
#   ↑ 这一步会**先校验子代理产出契约**（形状/必填/值域/日期可解析），不合契约直接报错退出、
#     完整清单落 <_out>/validation_report.txt；**没有跳过开关**。通过后交用户裁决（删行/合并）。
python scripts/build_matrix_rows.py final --preview <_out>/merged_preview.csv \
       --remove "42,50" --merge "30:52" --start-row-excel <起始行> --out-dir <_out>
```

**回填按 `config.excel.source` 分两路**：

- `local` → 用 `tencent-local-office-edit` 把 payload 写回本地 xlsx（含"文件被占用"的处理与
  `.xls` 被引擎写出成 xlsx 的坑，见 `references/workflow-notes.md`「回填（通道 A）」）。
- `kdocs` → `to_kdocs_payload.py` 生成 `calls` → agent 逐批 `sheet.update_range_data` → `get_range_data` 回读核对。

裁决 `merged_preview.csv` 时**顺手扫一眼**：子代理产出的 JSON 可能带 `null` / 空串字段
（脚本已兜底，但人工过一遍更稳）。

**派发判定子代理前**，按 `config.json` 渲染 `references/agent-prompt-zh.txt` 的**全部占位符**后使用
（**不要发未注入的裸模板**）：

| 占位符 | 取值来源 |
|---|---|
| `{{处理人}}` | `people.handler` |
| `{{我方标识}}` | `people.my_identifiers` |
| `{{服务对象}}` | `people.service_object` |
| `{{业务系统描述}}` | `excel.business_system_description` |
| `{{对方关键字}}` | `people.counterparty_keyword` |

判定开关 `rules.*`（`ignore_if_rejected` / `ignore_if_no_reply` / `ignore_vague_complaints` /
`data_change_default_done`）按模板的【开关渲染表】注入：开 = 用现文，关 = 替换为表中替代句。

## 表格来源：两条通道（`config.excel.source`）

用户可能给**本地文件**，也可能只说**"改 WPS 上那个文档"**。两条通道共用同一套列映射、判定规则与
payload 格式，只有「读表 / 回填」两端的 I/O 不同。

| 环节 | `local`（本地 xlsx） | `kdocs`（WPS 云文档） |
|---|---|---|
| 读表头/末行 | `probe.py workbook --workbook <xlsx>`（openpyxl） | agent 取数 → `sheet_snapshot.py build` → `probe.py workbook --snapshot <json>` |
| 岗位复用 | `final` 内置（读 workbook 历史，零成本） | 同样 `final` 内置，**前提是「第 2 趟」把行读全**（见下）；读不到时才用 `--history <json>` 手写兜底 |
| 算追加起始行 | `build_matrix_rows.py final --workbook <xlsx>` | `build_matrix_rows.py final --snapshot <json>` |
| 回填 | payload → `tencent-local-office-edit` | payload → `to_kdocs_payload.py` → agent 调 `sheet.update_range_data` |
| 额外依赖 | openpyxl | 金山文档连接器 |

云文档之所以要**经过「快照」**：连接器的 `sheet.*` 都是**只有 agent 能调的 MCP 工具，Python 脚本调不到**。
所以读表由 agent 取数落盘、脚本再读快照；写表由脚本产出请求体、agent 去调工具。
**不要把连接器当成脚本可 import 的库**，也不要在脚本里直连它。

- **中间产物流程结束要清理**：`raw_*.json` / `sheet_snapshot.json` / `kdocs_update*.json` /
  `validation_report.txt`（云文档那几条含真实表格内容与文档 id，报告里还有 agent 文件名）。
  别把它们留在用户工作目录，更**别留在仓库里**（`.gitignore` 已覆盖这些名字）。

- **三个"起始行"参数别混**（含义相同、写法不同，抄错会写错位置）：
  `final --start-row-excel <Excel 行号>` 由你显式指定、跳过自动推算；
  `position_reuse.py --start-row <Excel 行号>` 是它的扫描起点（只动这一行之后的新行）；
  `config.excel.start_row` 是**兜底配置项**，只在没给表格来源时才用得上。

### WPS 通道读表（分两趟，避免整表进上下文）

```bash
# 0) 定位文档（三选一，推荐给链接）
#    链接：从 URL 取 link_id → get_share_info → file_id / drive_id
#    文件名：search_files 搜索（可能同名，务必让用户确认）
#    再 sheet.get_sheets_info 取工作表清单、worksheet_id、已用区域 range.rowTo
#    建议把 file_id / drive_id / worksheet_id 写进 config.excel.kdocs，后续不用重复搜

# 1) 打印建议的读表调用体（含 file_id/worksheet_id/range 占位）
python scripts/sheet_snapshot.py plan --file-id <id> --worksheet-id <n> [--rows N] [--cols N]
#    第 1 趟：只读前 3 行 × 全部列（认表头）→ agent 把 sheet.get_range_data 的返回**原样**存 raw_hdr.json

# 2) 重建快照 → 解析表头列映射
python scripts/sheet_snapshot.py build --raw raw_hdr.json --out <output.dir>/sheet_snapshot.json
python scripts/probe.py workbook --snapshot <output.dir>/sheet_snapshot.json --json

# 3) 第 2 趟：按列映射只读关键列（需求描述/提出时间/提出人/提出人岗位/解决人…）→ raw_cols.json → 再 build
#    ⚠ rowTo 必须**读到表末**（不要只读前几行）：岗位复用要靠这一趟把 提出人+提出人岗位 读进来。
#      读全后回包很大，但**不必担心撑爆上下文** —— 宿主会把它自动**落盘**成文件，
#      直接拿那个落盘文件当 --raw 即可（实测依据见 references/workflow-notes.md「岗位复用」）。
python scripts/sheet_snapshot.py build --raw raw_hdr.json --raw raw_cols.json --out <output.dir>/sheet_snapshot.json
python scripts/sheet_snapshot.py inspect --snapshot <output.dir>/sheet_snapshot.json   # 核对末行/追加行
```

- `build` 会给每张子表写一条 **`coverage`**（本次读表**实际覆盖**到的行列范围，裁边之前的原始极值）
  和 `trimmed`（裁边后的实际长宽）。`final` 用它做两条前置检查：**行够不到历史区 → 直接报错**；
  列没覆盖到 提出人/提出人岗位 → 告警（可能没读，也可能该列历史区本来就全空）。
  ⚠ 回包**不带请求范围**（只有稀疏的 `rangeData`），所以 `coverage` 只能说明"看到了哪里"、
  不能自证"请求了多大范围"。

- **`get_range_data` 返回稀疏数组**（只含有值的单元格，行列 0-based）；`build` 会重建成密集网格，
  裁掉尾部/右侧空行列，并记录 `num_formats`。
- **只想定位末数据行**时，用 `get_typed_value`（A1 记法）更省。但它**会跳过空单元格、无法行对齐**，
  所以**不要**用它读关键列去建「人→岗位」映射——要行级定位必须用 `get_range_data`。

### WPS 通道回填

```bash
python scripts/to_kdocs_payload.py --payload payload.json   --out kdocs_update.json
python scripts/to_kdocs_payload.py --payload payload_n.json --out kdocs_update_n.json
```

产出 `calls` 数组（每项即一次 `sheet.update_range_data` 的 arguments，已按 100 条上限分批），
并按 `excel.date_columns` 自动补日期 `format` op（`numfmt` 取 `excel.kdocs.date_numfmt`）。

- **写入必须用 `sheet.update_range_data`**（幂等、按坐标）。不要用 `sheet.add_row`——非幂等，
  重试或重复调用会插多行脏数据。
- **写完必须用 `sheet.get_range_data` 回读同一区域核对**，不能只信返回的 `code: 0`。
- **日期格式以「最新行」为准**：同一列老行可能是 `yyyy/m/d`、新行是 `yyyy-mm-dd`（真机同列两种并存）。
  把最新行的写法填进 `excel.kdocs.date_numfmt`。
- **纯文本值会被表格引擎静默改写，脚本已自动转义**：因为值一律以 `opType:"formula"` 写入，
  `= + - @ '` 开头、带前导零、或 ≥16 位纯数字的文本会被当公式/数值处理 —— **改完不报错，值悄悄变了**。
  `to_kdocs_payload.py` **默认给这类值前置一个单引号**：引擎把它当"文本标记"消费掉、
  **回读值不含引号**（实测无损），并列出命中项与原因；
  确需按公式写入时才用 `--no-escape-risky-text`（会附带告警说明数据将不一致）。
  逐条实测（`0012`→12、`=A1`→被求真值、`+86`→86 …）见 `references/workflow-notes.md`「云文档纯文本风险」。

### 云文档接口约束（实测，写错即失败）

- **工作表参数名是 `worksheet_id`**，不是 `sheetId`：`get_sheets_info` / `get_range_data` / `update_range_data`
  都是这个键名。
- **单次 `update_range_data` 最多 100 条 `rangeData`**：超出报 `rangeData length 110 exceeds limit 100`。
  `to_kdocs_payload.py` 已**自动分批**，用输出的 `calls` 逐批调用。
- **op 数 ≈ 值单元格数，压不下去**：每条 `formula` op 只能写一个值，N 个不同值就是 N 条 op；
  能合并的只有 `format`/`merge`/`picture` 这类**区域**操作。所以 6 行 × 18 列 ≈ 107 条 op + 几条日期格式、
  分 2 批调用是**正常量级**，不是脚本啰嗦。
- **限频**：连接器有 `429001`/`429002` 熔断。不要逐格写、不要密集重试；命中限频要等响应给的时间。
- **区域保护**：文档若设了区域权限（`sheet.list_protection_ranges`），写入会失败，先让用户解除。
- **`.ksheet` 智能表格**：字段有类型（日期/单选等），写入形态与 `.xlsx` 不同，先确认文档类型。

## 岗位复用（`rules.reuse_position_column`，默认开）

按行序向上找同一提出人**最近的岗位值**填补空缺，**只填空缺、不覆盖更具体值**（`--override` 才覆盖）。

### 常规流程（内置在 `final`，一轮写入完成）

`final` 算出追加起始行后，从**同一个表格来源**读历史区 `[表头行+1, 起始行-1]`（**只读**），
把本批空岗位补齐，再写 `final_rows.csv` 与 `payload.json`。因此**不再需要**"先写一遍、再跑
`position_reuse` 补一遍"，`final` 的 CSV 与真正写进表格的内容**永远一致**（复用发生在写 CSV 之前）。

- **关闭**：`--no-reuse-position`，或 `rules.reuse_position_column=false`（脚本自检后整步跳过）。
- **历史区读到 0 个「人+岗位」对 → 显式告警并跳过**，并列出三种可能（行没读全 / 该列历史里确实没填过 / 只给了 `--start-row-excel`）。
- **只给 `--start-row-excel` 而没给 `--workbook`/`--snapshot`** → 同样**告警跳过**，不静默。
- **两条通道都自动读历史**：本地 openpyxl（零额外成本）；云文档**走快照** ——
  前提是「第 2 趟」把关键列**按行读全**（见「WPS 通道读表」），否则 `final` 会直接报错。
- **孤儿尾行按预期保留**：若末数据行只在非探针列（如项目编号）有值，它会被**当作有效数据行保留**，
  本批从它下方开始追加（宁留一个空行，也不覆盖）。要复用这类行需显式 `--start-row`。

**云文档为什么也能表内复用**：第 2 趟读表读的关键列**本来就包含** `提出人` 与 `提出人岗位`
（见 `sheet_snapshot.KEY_COLUMNS`），读进来的行又全在快照里 —— 所以只要**行读全**，快照天然含历史区，
`final --snapshot` 就与本地通道走同一条复用逻辑（同一个 `common.reuse_position_fill`）。

> 早期版本写的是"云文档拿不到便宜的历史岗位、一律手写 `--history`"，**那个结论已作废**：
> 它按"每格约 450 B ⇒ 200 KB 上下文"估算成本，漏掉了**大回包会被宿主落盘**这条路径
> （落盘文件可直接 `--raw` 喂 `build`，实测 196 格 / 87 KB 全量取回）。
> `--history` 只保留作**连接器不可用 / 落盘失败**时的降级，且它是**手写**的 —— 见下。

### `--history`（降级通道，需手写）

**仓里没有任何脚本会生成这个文件** —— 格式 `{"positions": {"人名": "岗位"}}` 只被**读**，
你要自己写（`position_reuse.py --out` 写的是 payload，不是它）：

```bash
python scripts/build_matrix_rows.py final --preview merged_preview.csv --start-row-excel 192 \
       --history history_positions.json     # {"positions": {"张三": "客服", "李四": "财务"}}
```

只在"快照确实读不到历史、而你又确有把握"时用。**多数情况下更省的做法**：像"备注无岗位、历史也没岗位"
的首次提法人，表内复用本来救不了（岗位信息根本不在数据里）—— 直接在 `merged_preview.csv` 的
「提出人岗位」列**裁决时补一个值**即可，成本为零且不依赖历史。
**不要把"补满岗位"当成必须达成的目标。**

### 独立补跑（`position_reuse.py`）

保留用于事后补漏或自定义区间，与 `final` **共用同一实现**（`common.reuse_position_fill`）。

**必须显式说明"哪些行算本批新行"，否则直接报错退出** —— 判据是"追加起始行 = 末数据行 + 1"，
从那里向下扫恒为空，所以默认模式永远得 0 条（表格尾部若还有空的"带格式行"，连提示都不打）：

- `--new-rows <final_rows.csv>`（推荐）：从 `final` 产物取新行，Excel 行号 = 起始行 + 序号，
  **不要求新行已落表**；行号由 `--workbook`/`--snapshot` 推算，或显式 `--start-row`。
- `--start-row <本批首行号>`：显式指定扫描起点（**会扫该行之后的全部行，含历史**）。
- 两者都不给 → **报错退出**。
- 历史行永远只读。

## 判定规则（对方提出 → 我方答复）

生效方式：`merge_cross_session` / `similarity_threshold` / `near_days` 由 `build_matrix_rows` 代码读取；
`ignore_if_rejected` / `ignore_if_no_reply` / `ignore_vague_complaints` / `data_change_default_done`
是**判定提示词开关**，由执行代理按 `references/agent-prompt-zh.txt` 的【开关渲染表】注入（不入 pipeline 脚本）；
`reuse_position_column` 控制岗位复用。把"忽略"类开关设为 `false`，会把被拒/无回复/笼统抱怨的需求
也登记为待裁决行（`outcome` 分别为 `rejected` / `no_reply` / `vague`）。

- **只记**对方提出的、针对系统的需求；寒暄/通知/闲聊不算。需求**必须可交付** —— 能落成一句
  "改什么 / 修什么 / 加什么"。
- **我方明确拒绝**（不能做/做不了/不支持/你们自己改…）→ 忽略（`rules.ignore_if_rejected`）。
- **提出后无实质回复** → 忽略（`rules.ignore_if_no_reply`）。
- **笼统抱怨不算需求**（`rules.ignore_vague_complaints`）：只表达体感/情绪、**指不出具体对象**的反馈
  （"系统特别卡 / 太慢了 / 又出问题了 / 不好用"），**且我方未明确答复已解决** → 忽略。
  反之，给出了可定位对象或量化现象（哪个功能 / 哪一步操作 / 哪个单据，或"搜一个箱号要等15秒"
  "审批后卡住"）→ 仍按 bug 记。
- **明确答复完成**（好了/已完成/改好了/搞定/你试试…）→ 记行；计划时间 = 完成时间 = 答复日期。
  注意**运维性动作与安抚不算完成**（重启系统 / 稍后看 / 换个浏览器 / 清缓存 / 再看看 / 我看看），
  不得据此记为"完成"。
- **数据修改类特例**（改费用/改单号/录补调数据/导数据…）：没答完成也没拒绝 → 默认已完成，
  时间取提出日期（`rules.data_change_default_done`）。
- **同一需求跨会话出现** → 合并一行，提出人取最早提出者（`rules.merge_cross_session`）。
- **提出时间取首次提出日**（不是我方答复日）；一条消息含多个需求 → 拆多行。
- **解决人 / 责任人** = `people.handler`。
- **子代理的 `note` 字段不会落进台账**：它出现在 `merged_preview.csv` 的「备注」列（供你裁决判断，
  关态开关下还会写"待人工裁决"），但 `final` 生成 payload 前会把备注**清空** ——
  台账单元格只放业务结果。别指望 `note` 出现在表里。

### 「需求归类」（Q 列）的取值域

**只能取这 5 个值**，照抄不要连写：

| 值 | 用于 |
|---|---|
| `数据处理` | 改数据（改费用/改单号/录补调/导数据…） |
| `需求` | 功能 / 界面 / 流程的**新增** |
| `优化` | 已有功能 / 界面的**改进** |
| `bug` | 系统出错 |
| `答疑` | 仅查证问询 |

- `需求` 与 `优化` **是并存的两种值**，不是「需求或优化」这种连写值 —— 后者台账里从不存在，
  由 `build_matrix_rows.QMAP` 归一成 `需求`；提示词规则文本已改成明确二选一，不再写"需求或优化"。
- `final` 遇到白名单外的值会**显式告警并列出序号**（不静默写入，也不擅自改值）。
- 取值域来自线上台账「需求归类」列 **196 格实测**：数据处理 104 / bug 31 / 答疑 27 / 优化 24 / 需求 9。
  `Q_VALUES` / 提示词枚举 / 本表三处由 `smoke_test.smoke_q_values` 守着，改一处漏一处会红。

### 改判定规则：必须做行为验证

改 `rules.*` 开关语义、或"什么算完成 / 什么算需求"的措辞之后，**光跑 `smoke_test` 不算验证** ——
它只能守住枚举与开关三处一致这类**形状**，守不住规则本身对不对。

```bash
python scripts/rules_check.py plan   --config config.json --out <工作目录>
#   ↑ 渲染提示词（占位符 + 依 config.rules 生成的开关覆盖表）+ 为每条夹具生成**盲评 case**
#   一个 case 派一个**独立**子代理判定 → 产出 case_<id>.json（case 文件里没有期望答案）
python scripts/rules_check.py verify --config config.json --results <工作目录>
```

- 差异分三类：**误记**（该忽略却记了）/ **漏记**（该记却没输出 —— 比误记更危险）/ **结果不符**；
  另有「未判定」（缺结果文件，不算通过）。
- 夹具在 `references/rules-fixtures.json`（**18 条**），每条都标了它守的 **SKILL.md 原文锚点** ——
  改规则时据此判断哪些样本受影响。样本用 `[我方]`/`[对方]` token，渲染时才注入 config 里的真实值。
- **规则没定义清楚的地方不许假装有答案**：这类夹具标 `open: true` 并登记进 `_open_questions`，
  `verify` 只报告不计失败。当前有 **1 条**：`reply-with-ops-action`（对方给了可定位对象、
  我方只回运维动作）该落哪个 `outcome`，规则从未定义。
- CI **不跑判定**（要模型、非确定性），只校夹具完备性：`smoke_test.smoke_rules_fixtures`（7 项）。

## 微信库结构（领域事实）

> 只列**会改变你怎么做**的部分。完整机制（消息表分片与 `Name2Id`、`local_type` 码表、
> `WCDB_CT_message_content` 取值、导出侧的三形态处理表）见 `references/workflow-notes.md`
> 「环境事实」与「消息正文解码」——那里是唯一权威，本节不复制。

1. **库位置**：`~/Documents/xwechat_files/<账户目录>/db_storage/`，SQLCipher 加密。
   一台机器可能有**多个账户目录**，**必须让用户选**（活跃的那个看 `db_storage` 下文件的 mtime）。
2. **正文可能是压缩的，导出侧必须能解**：`WCDB_CT_message_content=4` 的正文是 ZSTD 压缩 bytes，
   长文本/图片/引用/合并转发/通话基本都走这条路 —— **"压缩"不等于"没有需求"**。
   解不开时**输出明确标记**（绝不 `errors="replace"` 吐乱码，那会把整段会话变成乱码并漏掉需求），
   并在导出末尾**统计未解码条数**。
3. **`lt=49` / `lt=57` 的正文常藏在 `<refermsg>` 里**：对方用「引用回复」提需求时，`<title>` 通常只是
   一句「话头」，真正的需求在引用原文 `<refermsg><content>`；只输出 `[链接/文件]` 会**整条丢掉需求**。
   提取顺序：`title` + 引用正文 + `des`，都没有时再兜底取外层 `<content>` 或 `<url>`。
   **只取有语义的字段**，不做"去标签取全文"——那会把 `appid`/`fromusername` 之类 ID 混进转录。
4. **日期写入用 Excel 序列号**（1899-12-30 起），并对 `date_columns` 设日期格式；
   **子表名/列位置都靠表头与关键字动态匹配**（如 `sheet_match: "运维"`），不要写死年份或列字母。

## 速查

### 脚本清单（`scripts/`）

| 脚本 | 作用 |
|---|---|
| `smoke_test.py` | 自检：import / 纯函数 / CLI 守卫契约 / 离线端到端 / 仓库卫生 / 编码（改完代码先跑它；CI 入口） |
| `probe.py` | 探测 `accounts` / `sessions` / `workbook`（表头→列映射、末行、处理人候选） |
| `sheet_snapshot.py` | 云文档通道：`plan`（读表调用体）/ `build`（稀疏→密集快照，并记下 `coverage` 覆盖范围）/ `inspect` |
| `to_kdocs_payload.py` | 云文档通道：payload → `sheet.update_range_data` 的 `rangeData`（自动分批 + 日期格式 + 风险值转义） |
| `decrypt.py` | 封装 wcdb-key-tool 的 extract/decrypt（密钥按账号隔离，显式 `--db-storage`） |
| `export_conversations.py` | 按 config 导出转录；含消息正文解码（ZSTD 解压 / appmsg 引用提取 / 系统消息归一化） |
| `split_for_agents.py` | 转录按**文件大小均衡**分成 N 份，输出 `agent_N.txt` 清单 |
| `build_matrix_rows.py` | `preview`（**先校验子代理产出契约**，不合格即报错；再标跨会话重复）/ `final`（按列映射生成 payload，**内含岗位复用**） |
| `position_reuse.py` | 岗位列向上复用**独立补跑**（与 `final` 共用 `common.reuse_position_fill`；历史只读、只动新增行） |
| `rules_check.py` | 判定规则**行为验证**：`plan`（渲染提示词 + 生成盲评 case）/ `verify`（比对标注，三类差异） |
| `common.py` | 配置加载、路径识别、日期/列工具、云文档快照抽象层（`GridWorkbook`/`GridSheet`/`sparse_to_grid`）、控制台编码、输入路径守卫 |
| `pipeline.py` | `probe`（探测确认）/ `run`（执行） |

### 参数速查（常用命令）

下表只列常用命令。**`--config` 是全局选项，必须写在子命令之前** ——
`pipeline.py probe --config x.json` 会被 argparse 拒绝（`unrecognized arguments`），
正确写法是 `pipeline.py --config x.json probe`。
（**探测阶段可以先不给 `--config`**：账户与表头映射都不需要它。但**会话清单的位置要看它** ——
不给 config 时 `probe` 只找默认的 `./wechat_pilot/output/decrypted`，解密输出在别处就得给
`--config` 指到 `account.decrypted`；一旦要给，位置就必须对。）

**探测有两个入口，实现是同一份**，别以为是两套逻辑：
`pipeline.py probe` 是**编排入口**（内部跑 `probe.py` 的 JSON 子命令再汇总打印，给人读）；
`probe.py accounts|sessions|workbook --json` 是**脚本级入口**（输出 JSON，用来生成/粘贴
`excel.column_mapping`）。两者不重复解密、不重复导出。注意 **`probe.py` 没有 `--config` 选项**，
传了同样被 argparse 拒绝。

另外三个脚本通常由 `pipeline.py run` 代调，单独用时看各自 `--help` ——
其中这几个参数是**给用户自己调窗口用**的，别漏：
`decrypt.py --db-storage/--tool/--output/--reextract`；
`export_conversations.py --since/--until/--name-filter/--session/--group`（可重复给 `--session`/`--group` 精确点名）；
`split_for_agents.py --transcripts/--n/--out`。

| 命令 | 参数 |
|---|---|
| `pipeline.py --config <json> probe` | `(--workbook <xlsx>` \| `--snapshot <json>) [--since] [--keyword] [--dump]` |
| `pipeline.py --config <json> run` | `[--since] [--skip-decrypt] [--reextract]` |
| `probe.py` | `accounts --hint [--json]`；`sessions --decrypted <dir> [--since] [--keyword] [--json]`；`workbook (--workbook <xlsx> \| --snapshot <json>) [--sheet <关键字>] [--json]` |
| `sheet_snapshot.py` | `plan --file-id <id> --worksheet-id <n> [--rows] [--cols] [--letters "L,M,O"]`；`build --raw <f>… --out <json> [--sheet <名>] [--worksheet-id] [--file-id] [--drive-id] [--name]`；`inspect --snapshot <json>` |
| `build_matrix_rows.py` | `preview --src <_out> --out <csv> [--config]`（子代理产出不合契约即报错，完整清单落 `--out` 同目录的 `validation_report.txt`）；`final --preview <csv> [--remove "1,3"] [--merge "a:b"] [--out-dir] (--workbook <xlsx> \| --snapshot <json> \| --start-row-excel N) [--reuse-position \| --no-reuse-position] [--history <json>] [--config]` |
| `position_reuse.py` | `(--workbook <xlsx> \| --snapshot <json> \| --history <json>) --out <json> (--new-rows <final_rows.csv> \| --start-row N) [--override] [--end-row N] [--config]` —— 必须给 `--new-rows` 或 `--start-row`，两者都不给直接报错 |
| `to_kdocs_payload.py` | `--payload <json> [--out <json>] [--file-id] [--worksheet-id] [--date-cols "M,V,W"] [--date-numfmt yyyy-mm-dd] [--no-date-format] [--no-escape-risky-text] [--batch-size 100] [--config]` |
| `rules_check.py` | `plan --config <json> [--out <dir>] [--only id1,id2] [--fixtures <json>] [--allow-unfilled]`；`verify --config <json> --results <dir> [--fixtures <json>]` |

### 可配置项（见 `config.example.json`）

| 类别 | 变量 | 说明 |
|---|---|---|
| 账户 | `account.db_storage` / `decrypted`；`dir`（可选，仅人眼识别） | 多账号由用户选 |
| 表格来源 | `excel.source`（`local` \| `kdocs`） | 决定读表/回填走哪条通道 |
| 本地表格 | `excel.workbook` / `sheet_match` / `header_row` / `column_mapping` / `date_columns` / `start_row` | 列映射由 probe 依表头生成，勿手填 |
| 云文档 | `excel.kdocs.link` / `file_id` / `drive_id` / `name` / `worksheet_id` / `date_numfmt` | 用户给的定位形式三选一，**由你解析、代码不读 `link`**：link（推荐，你调 `get_share_info` 换成 `file_id`/`drive_id`）> `file_id` > `name`（搜索，可能同名）。脚本真正读的只有 `file_id` / `drive_id` / `worksheet_id` / `date_numfmt`，`worksheet_id` 写入必填 |
| 人员 | `people.handler` / `my_identifiers` / `counterparty_keyword` / `service_object` | 处理人、我方标识、对方筛选关键字、服务对象 |
| 范围 | `scope.since` / `until` / `sessions` / `name_filter` / `groups` | 时间窗与会话清单 |
| 规则 | `rules.*` | 见「判定规则」；`reuse_position_column` 控制岗位复用 |
| 默认值 | `defaults.*` | 项目编号/名称/是否收费/需求类型/子系统/产生阶段/状态 等固定列填充 |
| 输出 | `output.dir` / `agents` | 产物目录、并行子代理路数 |
| 词表 | `post_words` | 岗位词表（从"提出人"文本剥离岗位）。**按客户补全**，漏词会剥离不出岗位 |
| 提示词上下文 | `excel.business_system_description` | 渲染进 `{{业务系统描述}}` |
