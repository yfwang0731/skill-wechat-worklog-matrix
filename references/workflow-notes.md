# 机制说明与实测记录（微信 4.x → 需求跟踪矩阵）

> 本文件是 `SKILL.md` 的**详情层**：`SKILL.md` 讲"怎么做"，这里讲"**为什么这样** + 逐条实测依据"。
> 不记录任何具体客户/人员/项目信息，全部以占位符或配置键描述。
> 标 **【真机】** 的条目是 2026-09 在 Windows + 微信 4.1 + WPS 云文档上实跑踩出来的。

## 本文件与 `SKILL.md` 的分工（改文档前先读这段）

| | `SKILL.md` | 本文件 |
|---|---|---|
| 内容 | **规则与指令**：做什么 / 不许做什么 / 怎么执行 / 参数与命令 | **依据**：为什么这么定、实测数字、踩坑经过、历史案例 |
| 读者 | 执行任务的 agent（**激活时必读**） | 需要深究、或改规则的人（按需读） |

- **规则本体只在 `SKILL.md` 写一份，本文件不复制。** 早期两个文件在「云文档接口」「纯文本风险」
  「岗位复用」「判定规则」等**六个主题上各写一份**、措辞还不同 —— 双份真相必然漂移。
  现已收敛：本文件遇到这些主题只写**实测与理由**，规则本身指向 `SKILL.md` 对应小节。
- **两处不一致时，以 `SKILL.md` 为准**，并回头把本文件改成依据、不要再抄一遍规则。

**章节索引**：环境事实 · 解密 · 消息正文解码 · 列映射 · 表格来源两条通道 · 云文档接口约束 ·
云文档纯文本风险 · 岗位复用 · 判定规则（仅真机案例）· 易错点

## 环境事实

- 微信 4.1.x；进程 Weixin.exe。本地库在 `~/Documents/xwechat_files/<账户目录>/db_storage/`，SQLCipher 加密。
- **一台机器可能有多个账户目录**，必须由用户选择。判断"哪个在用"的快捷办法：看各账户 `db_storage`
  下文件的 mtime，活跃账户会持续更新；停用账户的 mtime 会停在很久以前。
- 消息库分片：`db_storage\message\message_1/2/3.db`（时间序 `_2` 最老 → `_3` 最新）；每会话消息表 `Msg_<md5(wxid)>`。
- 发送者数字 id = **分片内** `Name2Id.rowid`（**跨分片会变**），不要跨分片复用。
- `create_time` 秒级时间戳；`local_type` **低 32 位**为消息类型
  （1=文字，3=图片，34=语音，43=视频，47=表情，49=链接/文件，50=通话，57=引用回复）。高位还有别的位，必须先掩码。
- 会话范围/时间范围/处理人/对方关键字全部来自 `config.json`，不在代码里内置。
- 验数据是否"到今"：`SELECT MAX(create_time)` 扫各分片的 `Msg_*` 表；只看某一会话容易误判。

## 解密（wcdb-key-tool，第三方灰色工具，不在 skill 内）

- 准备：`git clone https://github.com/TANGandXUE/wcdb-key-tool scripts/tools/wcdb-key-tool`
  或设环境变量 `WCDB_KEY_TOOL` 指向 `wcdb_key_tool_windows.py`；`common.find_wcdb_tool` 会自动找
  （含家目录深度≤3 搜索）。**注意**：工具放在更深的目录里就找不到，此时用环境变量显式指定最稳。
- ⚠️ 供应链注意：第三方灰色工具会随上游更新改变行为。clone 后记录所用 commit
  （`git -C scripts/tools/wcdb-key-tool rev-parse HEAD`），后续解密异常先核对是否工具版本漂移。
- **多账户机器必须显式指定库目录**：`decrypt.py` 的参数是 `--db-storage <目录>`
  （`--db-dir` 是 wcdb-key-tool 自己的参数、由脚本内部透传，**不是** `decrypt.py` 的）；
  不指定则自动探测可能取到停用的那个账号（【真机】踩到）。
- 口令缓存 `~/.wcdb-key-tool/wechat-passphrase.json`：有效时无需重登；缓存丢失则先 `extract`（**微信须前台**）。
- 密钥缓存按账号隔离：`decrypt.py` 输出 `output/keys_<账号目录>.json`。
- 解密结果里 `.factory\...\*.db` 这类是**历史备份副本**，拿不到密钥属正常（会 SKIP），不影响主库。
- 合规：只读本机、数据不出电脑，使用前向用户确认。

## 消息正文解码（【真机】最容易被忽略的一环）

同一条 `Msg_<md5>` 表里，正文有三种形态，取决于 `WCDB_CT_message_content`：

| `WCDB_CT_message_content` | `message_content` 类型 | 处理 |
|---|---|---|
| `0` | `str` | 明文，直接用 |
| `4` | `bytes`，**ZSTD 压缩帧**（魔数 `28 B5 2F FD`） | 需 zstd 解压 |
| 其它/空 | `bytes` 或空 | 按文本尝试解码；乱码率 >30% 判为二进制 |

- **哪些消息会被压缩**：图片/视频/语音的描述、`appmsg`（链接/文件/引用/合并转发）、通话记录，
  以及**较长的文本**。所以"压缩"不等于"没有需求"——真机上解压出来的 8 条消息里有 4 条含需求正文。
- 缺 `zstandard` 时的行为：**输出明确标记**（`[压缩内容未解码: 需 pip install zstandard]`），
  并在导出末尾统计条数；**早期版本直接 `errors="replace"` 解成乱码**，导致整段会话不可读、需求漏判。
- 解压后往往还是 XML，需要再提取：
  - `appmsg` + `type=57`（**引用回复**）：`<title>` 只是"话头"，**真正的需求在
    `<refermsg><content>`**。只输出 `[链接/文件]` 会整条丢需求（真机上漏过一条）。
  - `appmsg` + `type=19`（合并转发的聊天记录）：`<des>` 里是整段对话，要截断。
- `appmsg` 三者（title / 引用正文 / des）**都取不到**时，再兜底取**外层 `<content>` 或 `<url>`** ——
  否则只剩一个 `[链接/文件]`，纯文本型 appmsg 的需求会整条丢掉。只取有语义的字段，
  不做"去标签取全文"（那会把 `appid`/`fromusername` 之类 ID 混进转录）。
- `sysmsg`：系统通知；`type=revokemsg` → `[撤回了一条消息]`，其它 → `[系统消息]`。
- `voipmsg`：通话记录 → `[通话]`。
- 结论：导出脚本 `render_content()` 是唯一的正文出口，新增消息类型时**只改这里**，并补 smoke 断言。

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
  `final` 就算出**更小**的起始行 → **静默覆盖**尾行（2026-09-11 实跑复现：probe=5 vs final=4）。
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

> **规则本体见 `SKILL.md`「云文档接口约束」**（`worksheet_id` / 必须用 `update_range_data` /
> 写后必须回读 / 限频怎么等）。本节只补**怎么测出来的**、以及踩坑当时的现场。

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

- 实测（2026-09-11，`create_file_with_content` 与 `sheet.update_range_data` 两条路径行为一致）：
  `0012` → 数值 **12**（丢前导零）；`=A1` → **被当真公式求值**，值变成 A1 单元格的内容（**不报错**）；
  `12345678901234567` → **丢精度**；`+86` → **86**。全部是**静默改写**。
- 现有对策：`to_kdocs_payload.py` 默认给这类值**前置一个单引号**，引擎把它当"文本标记"消费掉、
  **回读值不含引号**（实测无损），并列出命中项与原因；确需按公式写入才用 `--no-escape-risky-text`。

## 岗位复用

> **规则本体见 `SKILL.md`「岗位复用」**（算法 / 作用范围 / 何时关闭 / `--history` 怎么用）。
> 本节写依据：为什么云文档也能复用、以及那条"太贵所以不可行"的结论是怎么被推翻的。

### 两条通道都做表内复用（2026-09-18 **修正**）

**结论**：云文档**可以**表内复用，前提是按 `SKILL.md`「WPS 通道读表」的第 2 趟把关键列**按行读全**。
下文保留的四条"贵"的实测，其中 `get_range_data` 那条的**成本核算已被证伪**（见本节末）。

**为什么可行**：第 2 趟读的关键列（`sheet_snapshot.KEY_COLUMNS`）**本来就含** `提出人` 与 `提出人岗位`，
读进来的行全在快照里 ⇒ 只要行读全，快照天然含历史区，`final --snapshot` 走的就是与本地通道
**同一个** `common.reuse_position_fill`。实测（合成 raw，走真实 build/final）：
历史第 2 行 `提出人=张三 / 岗位=客服` → `final` 输出 `[reuse] R3 张三: '' -> '客服'`，payload 里 D 列写入「客服」。

`build` 会给每张子表写 `coverage`（**裁边之前**的原始极值）供 `final` 判两条：
行够不到历史区 → 报错；列没覆盖到 提出人/提出人岗位 → 告警（两可：没读到 / 该列历史区本全空）。
**回包不带请求范围**（只有稀疏的 `rangeData`），所以 `coverage` 只说明"看到了哪里"。

#### 四条路的原始实测（`get_range_data` 那条的**成本结论已作废**）

| 路径 | 实测结果 |
|---|---|
| `drive.download_file` → curl | 返回 `url` 但**需登录态**，直接 curl 得 `403 {"result":"userNotLogin"}`。**云文档无法下载到本地再走本地通道**（← 这条仍然成立） |
| `sheet.get_typed_value` | **跳过空单元格** → 行归属错位（见上）（← 仍然成立：建「人→岗位」映射**必须**用 `get_range_data`） |
| `sheet.get_range_data` | 每条带 alignment/fonts/`cell_background_color`/`hasBorder`/`numFormat`，约 **450 B/格**；⚠ **原结论"≈200 KB 上下文所以不可行"已作废** —— 大回包会被宿主**自动落盘**，用落盘文件当 `--raw` 喂 `build` 即可（实测见下） |
| `read_file(sheet_range=…)` | 返回**与 `get_range_data` 完全相同的带样式格式**（不是更精简的 markdown），一样贵。官方文档还给 xlsx 的 markdown 抽取标了「禁止」 |

另有 `sheet.find_range_data`（带 `filter`/`option_cols`）看起来像"按人名定点查"，
但 **`filter` 的结构未公开**：传错会被**静默忽略并返回整段**（比不查更贵）。
它顺带返回的 `option_col[].texts` 是**按列的值计数**（例：岗位列 14 格为空、「某提出人」在提出人列出现 1 次），
可作廉价的存在性判断，但**不给行级配对**。

#### 成本结论被证伪的过程（**可复用的判据**）

原结论把"450 B/格 × 382 格 ≈ 200 KB"当成**上下文成本**。但真机读 Q 列（196 格）时回包 **87 KB
被工具自动落盘**到 `tool-results/mcp-kdocs-*.txt`，全程没进上下文；把那个落盘文件直接
`--raw` 喂 `build` 实测成功（196 格 → 快照 196 行×17 列）。
⇒ **"某接口太贵"的判断，必须区分"回包大小"与"进上下文的大小"** —— 宿主对大回包的落盘是
一条常被漏掉的廉价路径。（注意：落盘是**宿主行为**、不是连接器保证，所以 `--history` 仍保留作降级。）

### 时序与历史（为什么现在是这个形态）

- **2026-09 起内置在 `final`**：历史区 = `[表头行+1, 起始行-1]`，复用发生在写 `final_rows.csv`
  **之前** ⇒ 一轮写入即可完成，且 CSV 与真正落表的内容**永远一致**；
  `position_reuse.py` 随之降级为"独立补跑/自定义区间"工具，两者共用 `common.reuse_position_fill`。
- **独立补跑为什么必须显式说明"本批新行"**：默认模式无从判断 —— 追加起始行 = 末数据行 + 1，
  从那里向下扫恒为空，只会得到 0 条；表格尾部若还有空的"带格式行"，**旧版本连提示都不打**（静默无操作）。
  现已改为默认报错退出。这是「绝不静默」在岗位复用上的具体形态。
- **孤儿尾行为什么"保留"而不是跳过**：只在非探针列（如项目编号）有值的尾行仍算有效数据行，
  本批从它**下方**追加 —— 宁可留一个空行，也绝不覆盖。要复用这类行需显式 `--start-row`。

## 判定规则（本文件只留真机案例，规则见 `SKILL.md`）

> **规则本体逐条在 `SKILL.md`「判定规则」**，本节**不再抄一遍** —— 早期两处各写一份、措辞不同，
> 是这套文档最容易漂移的地方。这里只记录**规则是怎么被打出来的**。

- **「系统特别卡」被记成 bug/完成**（【真机】）：群里只说了体感、我方只回"稍等 重启系统"。
  ⇒ 直接成因是两条规则：① 笼统抱怨指不出对象不记；② **运维性动作与安抚不算完成**。
- **提出时间被记成答复日**：把 09-09 的提出记成了 09-10 的答复日。
  ⇒ 成因是"首次提出日"这条规则此前没写死。
- **引用回复里的需求整条丢掉**：对方用「引用回复」提需求时，`<title>` 只是话头，
  真正的需求在 `<refermsg><content>`（见「消息正文解码」）。
- **「一条消息含多个需求」没拆行**、**跨会话重复没合并**：都是在真机裁决阶段才被人工发现的，
  现已写进规则。

## 易错点（跨主题）

- **绝不静默改写已有数据**：拿不到"追加起始行"就报错退出。【真机】暴露过**四条**这类路径 ——
  空表头映射、`final` 兜底第 2 行、`position_reuse` 默认扫整表、以及「末数据行」判据双实现分歧
  （见上「列映射」）；现全部改为显式报错或显式 opt-in。
- 名称重复联系人：contact 表一个 display 可能对应多个历史 username，取消息量最大者（export 已处理）。
- 用户在 config 里点名的会话，若时间窗内无消息要**如实报告**（【真机】遇到点名的群两周没发言），
  别默默产出空结果；让用户决定是否扩窗。
- 子代理判定出入可能很大：`merged_preview.csv` 必须给用户裁决（删行/合并）后再回填。
- 大批量写值别逐格调；payload JSON 用 `ensure_ascii=False` 保持中文可读。
