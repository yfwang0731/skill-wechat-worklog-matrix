# 机制说明与实测记录（微信 4.x → 需求跟踪矩阵） （随 skill 发布｜禁写本机路径、账号、uid、用户数据、运行期结果）

> 本文件是 `SKILL.md` 的**详情层**：`SKILL.md` 讲"怎么做"，这里讲"**为什么这样** + 逐条实测依据"。
> 不记录任何具体客户/人员/项目信息，全部以占位符或配置键描述。
> 标 **【真机】** 的条目是在 Windows + 微信 4.1 + WPS 云文档上实跑踩出来的。

## 本文件与 `SKILL.md` 的分工（改文档前先读这段）

| | `SKILL.md` | 本文件 |
|---|---|---|
| 内容 | **规则与指令**：做什么 / 不许做什么 / 怎么执行 / 参数与命令 | **依据**：为什么这么定、实测数字、踩坑经过、历史案例 |
| 读者 | 执行任务的 agent（**激活时必读**） | 需要深究、或改规则的人（按需读） |

- **规则本体只在 `SKILL.md` 写一份，本文件不复制。** 两个文件若在「云文档接口」「纯文本风险」
  「岗位复用」「判定规则」这类主题上各写一份、措辞还不同 —— 双份真相必然漂移。
  现已收敛：本文件遇到这些主题只写**实测与理由**，规则本身指向 `SKILL.md` 对应小节。
- **两处不一致时，以 `SKILL.md` 为准**，并回头把本文件改成依据、不要再抄一遍规则。
- **`references/kdocs-channel.md` 是第三层**：`excel.source = kdocs` 那一端的**操作详情**（读表两趟、
  回填请求体怎么造、接口约束）。它与本文件的分工是「**操作详情 vs 实测依据**」—— 该文件不写"为什么"，
  实测数字与踩坑经过一律落在本文件对应小节；规则本体仍在 `SKILL.md`。

**章节索引**：环境事实 · 解密 · 消息正文解码 · 列映射 · 表格来源两条通道 · 云文档接口约束 ·
云文档纯文本风险 · 岗位复用 · 判定规则（仅真机案例）· 自检与 CI · 易错点

上述主题里，「环境事实」「解密」「消息正文解码」三节见 `references/wechat-db-notes.md`；
「岗位复用」「判定规则」两节见 `references/reuse-judgement-notes.md` —— 本文件只留「列映射」及其后的
通道 / 接口约束 / 纯文本风险 / 自检与 CI / 易错点诸节，是三份记忆文件里的**主文件与索引**。

## 分册索引

| 分册 | 负责的主题 |
|---|---|
| `references/wechat-db-notes.md` | 环境事实 / 解密 / 消息正文解码 |
| `references/reuse-judgement-notes.md` | 岗位复用 / 判定规则 |

**本目录文件与发布属性**（全部随 skill 发布）：

| 文件 | 发布属性 |
|---|---|
| `references/workflow-notes.md` | 随包发布（记忆文件，单文件不超过 8 千字符） |
| `references/wechat-db-notes.md` | 随包发布（记忆文件，单文件不超过 8 千字符） |
| `references/reuse-judgement-notes.md` | 随包发布（记忆文件，单文件不超过 8 千字符） |
| `references/kdocs-channel.md` | 随包发布（记忆文件，单文件不超过 8 千字符） |
| `references/agent-contract.json` | 随包发布（契约类，非记忆文件，不受单文件上限约束） |
| `references/rules-fixtures.json` | 随包发布（夹具类，非记忆文件，不受单文件上限约束） |
| `references/agent-prompt-zh.txt` | 随包发布（模板类，非记忆文件，不受单文件上限约束） |

判据与盲区见 `scripts/smoke_test.py` 的守卫注释（`smoke_doc_layering` / `smoke_doc_refs` /
`smoke_doc_structure` / `smoke_doc_claims`）—— 它们按符号锚定、不靠行号，改章节或搬文件后漏同步会当场报红；本文件不复制判据，复制就会漂移。

## 列映射

- **列位置靠"表头名"识别**：`probe.py workbook` 读表头，按 `COLUMN_ALIASES`（逻辑列名 → 可能的表头名）匹配，
  产出 `column_mapping`（如 `"需求描述"→"L"`）。换模板也能用，不再写死 A~AA。
- 匹配不到的表头列不写入（会打印未识别清单与 `warnings`）。
- **表头名带前导/尾随空格很常见**（真机遇到 `" 提出时间"`），已统一 strip 后再匹配。
- **表头行怎么定**：取"前 5 行里命中 `COLUMN_ALIASES` 别名数最多"的一行，非空单元格数只作次级判据。
  只按"非空最多"在**很宽的表**上会被数据行骗过（数据行的需求描述/备注往往更长更满 →
  误判后列映射为空，整表读不了；虽然有守卫兜底不写脏数据，但等于白跑）。
- **`mapping` 为空时必须拒绝给追加行号**：末数据行是靠"已映射列有值"判定的，
  空映射会让 `any()` 恒 False、末行停在第 1 行、`next_append_row` 变成 2 → 据此写入会**覆盖已有数据**。
  现在返回 `next_append_row: null` + warning，`build_matrix_rows final` 也会拒绝执行。
- **「末数据行」只能有唯一判据**（`common.last_data_row_ws`，`probe` 与 `final`/`position_reuse` 共用）。
  两处各写一套曾造成分歧：probe 用**全部映射列**、`next_append_row_ws` 只用
  "需求描述/提出时间/提出人" 3 个探针列；尾部行若只在**别的**映射列（如 项目编号）有值，
  `final` 就算出**更小**的起始行 → **静默覆盖**尾行（实测：probe=5 vs final=4）。
  已抽成同一函数，并用 smoke 断言锁死"两处同值 + `header_row` 必须透传"。
- `date_columns` 里的列以 **Excel 序列号**写入，并设日期格式（**以目标表最新行的 `numFormat` 为准**）。

## 两种表格来源（`excel.source`：local / kdocs）

两条通道算法完全一致（列映射 / 判定规则 / payload 结构 / 岗位复用算法都不变），只换读表与回填的 I/O；
差别只在**取数成本**：本地 openpyxl 读表是免费的；云文档读表要经 agent 取数（大回包会落盘，
见下「岗位复用」一节的成本修正），且**第 2 趟必须把关键列按行读全**，否则岗位复用不生效。

### 通道 A：local（本地 .xlsx）

- 读：`probe.py workbook --workbook <xlsx>`（openpyxl）。
- 追加起始行 + **岗位复用**：`build_matrix_rows.py final --workbook <xlsx>`（一次读表两用；复用零额外成本）。
- 回填：editor_sdk（见下节）。

### 回填（通道 A：editor_sdk / tencent-local-office-edit 技能）

1. 先 Skill 加载 tencent-local-office-edit 拿 edsdk.py 路径；读 sheet.md。
2. open_file → 按 `excel.sheet_match` 关键字匹配子表（不要写死年份）→ sheet_get_used_range 探边界；
   追加起始行由 `common.next_append_row_ws(ws, mapping, header_row)` 算出（`build_matrix_rows final`
   与 `position_reuse` 共用同一判据，见「列映射」）。
3. 追加：sheet_insert_dimension(row, 末行+1, N) → sheet_set_range_value（每格必带 row/col/value_type
   + string_value|number_value）→ 对 date_columns 设 number_format_pattern → save_file。
4. ⚠️ 引擎保存会把旧 .xls 内容写成 xlsx：**保存后确认/改名为 .xlsx**，否则打开弹"格式与扩展名不匹配"。
5. ⚠️ 文件被宿主预览/用户打开（"Export file is occupied"）：save_file 到临时路径 → close_file(force)
   → 用 python 二进制覆写原文件内容（delete/rename 会被锁拒绝，写共享通常允许）；最后清理临时文件。
6. 保存后不要主动 close_file（用户可能在看）；需改名/重开时再 close。

### 通道 B：kdocs（WPS 云文档）

之所以必须多一层「快照」：连接器的 MCP 工具**只有 agent 能调，Python 脚本调不到**。所以：

- 读：agent 调 `sheet.get_range_data` → 原始返回**原样**落 `raw_*.json` → `sheet_snapshot.py build`
  重建密集网格快照 → 脚本读快照。
- 写：脚本产出 payload → `to_kdocs_payload.py` 转 `rangeData` → agent 调 `sheet.update_range_data`。

读表**分两趟窄读**，别把整表搬进上下文：

1. `sheet.get_sheets_info` 拿工作表清单 / `worksheet_id` / 已用区域上限。
2. 第 1 趟只读前 3 行 × 全部列 → 认表头（目的是拿 `column_mapping`）。
3. 第 2 趟按映射只读关键列（需求描述/提出时间/提出人/解决人…）→ 算末数据行。
4. 只想知道**末数据行**时，`sheet.get_typed_value`（A1 记法，返回紧凑 type+value）很省 ——
   【真机】就是用它在 2 次调用内定位到末行的。
   ⚠️ **但它会"跳过空单元格"**（8 格只回 7 值），**行列位置对不上**，
   所以**绝不能**用它建"逐行的人→岗位"映射（真机实测：`N190:O193` 4 行 × 2 列只回 4 个值）。
   需要行级定位时必须用 `get_range_data`（每条自带 `rowFrom/colFrom`）。

## 云文档接口约束（实测依据）

> **规则本体见 `references/kdocs-channel.md`「云文档接口约束」**（`worksheet_id` / 必须用 `update_range_data` /
> 写后必须回读 / 限频怎么等）。本节只补**怎么测出来的**、以及踩坑的现场。

- 选区坐标 camelCase：`rowFrom/rowTo/colFrom/colTo` + `opType`；日期格式走 `xf.numfmt`。
- **单次 `rangeData` 最多 100 条**，超出报 `length N exceeds limit 100` → `to_kdocs_payload.py` 自动分批。
- **每条 `formula` op 只能写一个值**，N 个不同值就是 N 条 op，无法靠合并收敛；能合并的只有
  `format`/`merge`/`picture` 这类区域操作。所以"6 行 × 18 列 ≈ 107 条 op"是**正常量级**，不是脚本啰嗦。
- `get_range_data` 返回**稀疏**数组（行列 0-based，只含有值的单元格）；range 覆盖多格时按"合并单元格"
  语义整块填同一值。`build` 会重建成密集网格并裁掉尾部/右侧空行列。
- 限频现场的教训：**别逐格写、别密集重试**。命中的响应里给了恢复时间，照它等；立刻重试只会再撞一次。
- `.ksheet` 智能表格有字段类型（日期/单选等），写入形态与 `.xlsx` 不同。
- 文档若设区域保护（`sheet.list_protection_ranges`），写入会失败。

## 云文档纯文本风险（实测依据）

值一律以 `opType:"formula"` 写入 —— 这是**引擎的写入契约**，不是本项目的选择，所以才会有下面的静默改写。

- 实测（`create_file_with_content` 与 `sheet.update_range_data` 两条路径行为一致）：
  `0012` → 数值 **12**（丢前导零）；`=A1` → **被当真公式求值**，值变成 A1 单元格的内容（**不报错**）；
  `12345678901234567` → **丢精度**；`+86` → **86**。全部是**静默改写**。
- 现有对策：`to_kdocs_payload.py` 默认给这类值**前置一个单引号**，引擎把它当"文本标记"消费掉、
  **回读值不含引号**（实测无损），并列出命中项与原因；确需按公式写入才用 `--no-escape-risky-text`。

## 自检与 CI（依据）

命令、以及"改完代码要跑什么"写在 `SKILL.md`；这里只记**矩阵为什么这么切**。

- **只有 Windows 一轴**：主链（微信 PC 客户端 + `wcdb-key-tool`）只在 Windows 上有意义，
  多摆一个 Linux 轴只会产出"永远绿、又永远不代表真实条件"的格子。
  另有几条**大小写敏感**的文件名判据（文档里把文件名大小写写错时，Windows / macOS 会判为"存在"）
  靠逐级 `os.listdir` 精确名比对实现 —— 这条本地与 CI 都跑得到，不需要额外轴。
- **依赖分两档**：`minimal` 守的是"零依赖也能跑 + 缺库给可执行提示"这条路本身；
  本地表格通道读表与 ZSTD 正文解压**只有 `full` 那档跑得到** —— 只留 minimal 等于这两块没人测。
- **Python 跨两代（3.9 / 3.13）**：3.9 是声明的最低支持版本。只测本机那一代，
  要等用户装不上库时才发现。
- **故意不设 `PYTHONUTF8` / `PYTHONIOENCODING`**：设了就把 Windows 的 cp1252 缺陷盖住，
  那一格等于白跑。脚本必须靠 `common.ensure_utf8_stdio()` 自己切标准流 ——
  这条**只有在没设这两个变量的环境里**才验得到，所以别为了"输出好看"给 CI 加上它们。
- **本地与 CI 共用一个入口**（不另写一套），否则必然出现"CI 绿、本地跑不到"或反之。
- **那条"正文不写编年"的仓库约定**：判据（三条臂）、豁免窗口与**已知盲区**都写在
  `scripts/smoke_test.py` 的 `smoke_doc_claims` 注释里 —— 这里不复制，复制就会漂移。
- **分块跑**：`python -B scripts/smoke_test.py --list-blocks` 列出所有块；
  `--only <块[,块…]>` 只跑选中的块（改动后按受影响面选块，不必每次全量）。
  `--only` 会跳过"项数 vs README"的元检查 —— 核项数只能用默认档。
- **并行跑**：`--jobs <N>`（N≥2）按块起子进程并行；输出与串行**逐字节一致**
  （这一条本身是断言式验收，防"并行丢展示行"）。

## 易错点（跨主题）

- **绝不静默改写已有数据**：拿不到"追加起始行"就报错退出。【真机】暴露过**四条**这类路径 ——
  空表头映射、`final` 兜底第 2 行、`position_reuse` 默认扫整表、以及「末数据行」判据双实现分歧
  （见上「列映射」）；现全部改为显式报错或显式 opt-in。
- 名称重复联系人：contact 表一个 display 可能对应多个历史 username，取消息量最大者（export 已处理）。
- 用户在 config 里点名的会话，若时间窗内无消息要**如实报告**（【真机】遇到点名的群两周没发言），
  别默默产出空结果；让用户决定是否扩窗。
- 子代理判定出入可能很大：`merged_preview.csv` 必须给用户裁决（删行/合并）后再回填。
- 大批量写值别逐格调；payload JSON 用 `ensure_ascii=False` 保持中文可读。
