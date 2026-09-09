---
name: wechat-worklog-matrix
description: >-
  通用「微信聊天记录 → 需求跟踪矩阵 Excel」流水线。扫描本机微信账户、由用户选择账户/会话/时间范围/Excel/处理人，
  解密导出聊天，LLM 识别对方提出的系统需求并判定答复结果，合并跨会话重复，按**表头**映射到 Excel 追加行。
  适用于任何"客户通过微信提需求、运维方需登记成台账/矩阵"的场景，不绑定具体客户或项目。
  触发场景：更新需求跟踪矩阵 / 从微信聊天记录整理需求台账 / 把微信里的需求登记进 Excel。
agent_created: true
---

# 微信记录 → 需求跟踪矩阵（通用流水线）

把微信 PC 聊天记录整理成需求跟踪矩阵台账。服务对象、项目、人员**全部由用户配置**，skill 本身不内置任何具体客户/项目/姓名。

## 何时使用
用户要"根据微信聊天记录更新/补充需求跟踪矩阵"、要读微信 4.x 本地记录、或要把微信里的需求登记进 Excel 台账。

## 环境依赖（首次使用前）
- Python 3.9+，需 `pip install openpyxl`（probe workbook / build final / position_reuse 均依赖）。
- Excel 须为 `.xlsx`（`.xlsm` 也支持）；旧版 `.xls`/`.xlt` 会被 `probe` 直接拒绝，请先用 Excel/WPS「另存为 .xlsx」再运行。
- 解密用第三方 wcdb-key-tool（见"关键事实"第 2 条），skill 内不含。
- 修改 scripts 后跑一次自检：`python scripts/smoke_test.py`（纯 import + 纯函数断言，不碰真实数据；加 `--full` 会额外校验 openpyxl 依赖项）。

## 核心原则（勿重蹈覆辙）
1. **不写死任何具体项目/人员/客户** —— 一切从 `config.json` 读。
2. **列位置靠"表头名"识别**，不写死 A~AA（换模板也能用）。
3. **账户/会话/时间范围由用户选**，不内置筛选条件。
4. **先探测 → 一次性确认 → 再执行**，不要逐个追问、不要替用户假设。
5. **单元格只写业务结果**，禁止把判定依据/分析过程写进任何单元格（尤其备注列）。

## 交互式流程（推荐）

### 第 1 步：探测（agent 自动执行，一次跑完）
```
python pipeline.py probe [--workbook <用户给的xlsx>] [--since <日期>] [--keyword <可选筛选>] [--dump <探测结果json>]
```
一次探测出三件事并汇总打印：
- **①本机微信账户列表**（可能有多个账号，让用户选）
- **②Excel 表头 → 列映射**（读表头名自动识别；报告未识别列、末数据行、追加起始行、处理人候选）
- **③可选会话清单**（名称、是否群、消息数、首末消息时间；需先有解密库，没有则先解密再重跑）

### 第 2 步：一次性确认（把探测结果一起呈现给用户）
用 AskUserQuestion（一次最多 4 问）或直接列出，让用户一次性确认/修正：
1. **用哪个微信账户**（探测①的列表）
2. **用哪个 Excel**（探测②；路径由用户提供，不写死）
3. **查哪些会话**（全部 / 按关键字筛选 / 指定编号，来自探测③）
4. **时间范围**（起止日期）

另外需用户提供：**处理人姓名**（可从探测②的"处理人候选"推断后请其确认）。
会话很多时，把清单写成文件让用户勾选，不要逐条追问。

### 第 3 步：写 config.json
从 `config.example.json` 复制，填入上述确认结果（account / excel / people / scope / defaults）。

### 第 4 步：执行
```
python pipeline.py run                    # 解密 → 导出 → 分包（产物在 output.dir/transcripts/_out）
python scripts/build_matrix_rows.py preview --src <output.dir>/transcripts/_out --out merged_preview.csv
python scripts/build_matrix_rows.py final --preview merged_preview.csv --remove "..." --merge "a:b"
python scripts/position_reuse.py --workbook <xlsx> --out payload_n.json   # rules.reuse_position_column=false 时跳过
```
再用 tencent-local-office-edit 回填 payload。

**派发判定子代理前**：按 config.json 渲染 `references/agent-prompt-zh.txt` 的全部占位符后使用——`{{处理人}}`←people.handler、`{{我方标识}}`←people.my_identifiers、`{{服务对象}}`←people.service_object、`{{业务系统描述}}`←excel.business_system_description、`{{对方关键字}}`←people.counterparty_keyword；判定开关 `rules.*`（ignore_if_rejected / ignore_if_no_reply / data_change_default_done）按模板【开关渲染表】注入（开=现文，关=替换为表中替代句）。**不要发未注入的裸模板**。
`position_reuse.py` 默认只填空缺、不覆盖；加 `--override` 才把空缺写为上方最近岗位。`rules.reuse_position_column=false` 时跳过（脚本读取 config 自检退出）。
（`<output.dir>/transcripts/_out` 即 run 产物目录；`output.dir` 见 config，默认 `./wechat_pilot`）

### 可选参数速查（与 argparse 一致）
| 命令 | 参数 |
|---|---|
| `pipeline.py probe` | `--workbook <xlsx>` `--since <日期>` `--keyword <串>` `--dump <json>` |
| `probe.py`（细粒度） | `accounts --hint [--json]`；`sessions --decrypted <dir> [--since] [--keyword] [--json]`；`workbook --workbook <xlsx> [--sheet <关键字>] [--json]` |
| `pipeline.py run` | `[--since] [--skip-decrypt] [--reextract]` |
| `build_matrix_rows.py preview` | `--src <_out目录> --out <csv> [--config]` |
| `build_matrix_rows.py final` | `--preview <csv> [--remove "1,3"] [--merge "a:b"] [--out-dir] [--workbook] [--start-row-excel] [--config]` |
| `position_reuse.py` | `--workbook <xlsx> --out <json> [--override] [--start-row] [--end-row] [--config]` |

## 可配置项（都可变量化，见 config.example.json）
| 类别 | 变量 | 说明 |
|---|---|---|
| 账户 | `account.db_storage` / `decrypted`；`dir`(可选，仅人眼识别) | 多账号由用户选 |
| Excel | `workbook` / `sheet_match` / `header_row` / `column_mapping` / `date_columns` / `start_row` | 列映射由 probe 依表头生成 |
| 人员 | `people.handler` / `my_identifiers` / `counterparty_keyword` | 处理人、我方标识(→渲染进 {{我方标识}})、对方筛选关键字 |
| 范围 | `scope.since` / `until` / `sessions` / `name_filter` / `groups` | 时间窗与会话清单 |
| 规则 | `rules.*` | ignore_if_rejected / ignore_if_no_reply / data_change_default_done → 派发判定子代理时注入提示词(见 agent-prompt-zh【开关渲染表】)；merge_cross_session / similarity_threshold / near_days → build_matrix_rows 代码读取；reuse_position_column → 是否执行 position_reuse |
| 默认值 | `defaults.*` | 项目编号/名称/是否收费/需求类型/子系统/产生阶段/状态 等固定列填充 |
| 输出 | `output.dir` / `agents` | 产物目录、并行子代理路数 |
| 词表 | `post_words` | 岗位词表（从"提出人"文本剥离岗位用） |
| 提示词上下文 | `people.service_object` / `excel.business_system_description` | 渲染进判定提示词的 `{{服务对象}}` / `{{业务系统描述}}`（可选，缺失时由执行代理据用户描述补充） |

## 判定规则（对方提出 → 我方答复）
> 生效方式：merge_cross_session / similarity_threshold / near_days 由 build_matrix_rows 代码读取；ignore_if_rejected / ignore_if_no_reply / data_change_default_done 是"判定提示词开关"，由执行代理在派发子代理时按 references/agent-prompt-zh.txt 的【开关渲染表】注入生效（不入 pipeline 脚本）；reuse_position_column=false 时跳过岗位复用命令（position_reuse 读取 config 自检退出）。把"忽略"类开关设为 false 会把被拒/无回复的需求也登记为待人工裁决行。
- 只记**对方提出的、针对系统的需求**；寒暄/通知/闲聊不算。
- 我方明确拒绝（不能做/做不了/不支持/你们自己改…）→ 忽略（`rules.ignore_if_rejected`）。
- 提出后无实质回复 → 忽略（`rules.ignore_if_no_reply`）。
- 明确答复完成（好了/已完成/改好了/搞定/你试试…）→ 记行；计划时间=完成时间=答复日期。
- **数据修改类特例**（改费用/改单号/录补调数据/导数据…）：没答完成也没拒绝 → 默认已完成，时间取提出日期（`rules.data_change_default_done`）。
- 同一需求跨会话出现 → 合并一行，提出人取最早提出者（`rules.merge_cross_session`）。
- 解决人/责任人 = `people.handler`。
- **岗位列复用**：同一提出人向上找最近一次岗位值填补空缺，**只填空缺、不覆盖更具体值**（`rules.reuse_position_column`）。

## 脚本清单（scripts/）
- `smoke_test.py` — 自检：import 全模块 + 纯函数断言（改完代码先跑它）
- `probe.py` — 探测 `accounts` / `sessions` / `workbook`（表头→列映射、末行、处理人候选）
- `decrypt.py` — 封装 wcdb-key-tool 的 extract/decrypt（密钥按账号隔离）
- `export_conversations.py` — 按 config 导出指定会话转录
- `split_for_agents.py` — 转录按**文件大小均衡**分成 N 份，输出 agent_N.txt 清单
- `build_matrix_rows.py` — preview（去重标记+裁决）/ final（按列映射生成 payload）
- `position_reuse.py` — 岗位列向上复用（依赖 openpyxl）
- `common.py` — 配置加载、路径自动识别、日期/列工具
- `pipeline.py` — `probe`（探测确认）/ `run`（执行）

## 关键事实（踩坑记录）
1. 微信 4.x（Weixin 4.1+）本地库 `~/Documents/xwechat_files/<账户目录>/db_storage/`，SQLCipher 加密；**一台机器可能有多个账户目录**，必须让用户选。
2. 解密用第三方 **wcdb-key-tool**（不在 skill 内，灰色工具，只读本机，使用前须经用户确认）：
   `git clone https://github.com/TANGandXUE/wcdb-key-tool scripts/tools/wcdb-key-tool` 或设环境变量 `WCDB_KEY_TOOL`。
   口令缓存 `~/.wcdb-key-tool/wechat-passphrase.json` 有效时无需重登；缓存失效需 `extract`（微信须前台）。
3. 消息表 `Msg_<md5(wxid)>` 分布在 `message_1/2/3.db`（_2 最老 → _3 最新）；发送者 id = 分片内 `Name2Id.rowid`（跨分片会变）；`create_time` 秒级；`local_type` 低 32 位为类型（49=链接/文件）。
4. 日期写入用 **Excel 序列号**（1899-12-30 起），并对 `date_columns` 设 `yyyy-mm-dd` 格式。
5. 子表名**按关键字动态匹配**（默认"运维"），不要写死年份。

## 易错点
- 会话很多时别用 AskUserQuestion 逐条问；探测后写清单交用户勾选，或按关键字过滤。
- 子代理判定出入较大：merged_preview.csv 必须交用户裁决后再回填。
- 文件被打开导致 "Export file is occupied"：save_file 到临时路径 → close_file(force) → 二进制覆写原文件。
- 引擎保存会把 .xls 写成 xlsx：**保存后确认/改为 .xlsx**。
- 详细坑位见 `references/workflow-notes.md`。
