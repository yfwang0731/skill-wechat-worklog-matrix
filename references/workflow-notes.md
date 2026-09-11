# 工作流要点与坑位备忘（微信 4.x → 需求跟踪矩阵，通用）

> 本文件不记录任何具体客户/人员/项目信息，全部以占位符或配置键描述。
> 标 **【真机】** 的条目是 2026-09 在 Windows + 微信 4.1 + WPS 云文档上实跑踩出来的。

## 环境事实
- 微信 4.1.x；进程 Weixin.exe。本地库在 `~/Documents/xwechat_files/<账户目录>/db_storage/`，SQLCipher 加密。
- **一台机器可能有多个账户目录**，必须由用户选择。判断"哪个在用"的快捷办法：看各账户 `db_storage`
  下文件的 mtime，活跃账户会持续更新；停用账户的 mtime 会停在很久以前。
- 消息库分片：`db_storage\message\message_1/2/3.db`（时间序 `_2` 最老 → `_3` 最新）；每会话消息表 `Msg_<md5(wxid)>`。
- 发送者数字 id = **分片内** `Name2Id.rowid`（**跨分片会变**），不要跨分片复用。
- `create_time` 秒级时间戳；`local_type` **低 32 位**为消息类型
  （1=文字，3=图片，34=语音，43=视频，47=表情，49=链接/文件，50=通话）。高位还有别的位，必须先掩码。
- 会话范围/时间范围/处理人/对方关键字全部来自 `config.json`，不在代码里内置。
- 验数据是否"到今"：`SELECT MAX(create_time)` 扫各分片的 `Msg_*` 表；只看某一会话容易误判。

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
  - `sysmsg`：系统通知；`type=revokemsg` → `[撤回了一条消息]`，其它 → `[系统消息]`。
  - `voipmsg`：通话记录 → `[通话]`。
- 结论：导出脚本 `render_content()` 是唯一的正文出口，新增消息类型时**只改这里**，并补 smoke 断言。

## 列映射（关键改进）
- **列位置靠"表头名"识别**：`probe.py workbook` 读表头，按 `COLUMN_ALIASES`（逻辑列名 → 可能的表头名）匹配，
  产出 `column_mapping`（如 `"需求描述"→"L"`）。换模板也能用，不再写死 A~AA。
- 匹配不到的表头列不写入（会打印未识别清单与 `warnings`）。
- **表头名带前导/尾随空格很常见**（真机遇到 `" 提出时间"`），已统一 strip 后再匹配。
- **`mapping` 为空时必须拒绝给追加行号**：末数据行是靠"已映射列有值"判定的，
  空映射会让 `any()` 恒 False、末行停在第 1 行、`next_append_row` 变成 2 → 据此写入会**覆盖已有数据**。
  现在返回 `next_append_row: null` + warning，`build_matrix_rows final` 也会拒绝执行。
- `date_columns` 里的列以 **Excel 序列号**写入，并设日期格式（**以目标表最新行的 `numFormat` 为准**）。

## 判定规则（用户口径，可由 config.rules 开关）
- 对方提出系统需求才算；寒暄/通知不算。拒绝或无实质回复 → 不写行。
- 明确答"完成" → 计划时间=完成时间=答复日期。
- **提出时间 = 首次提出日**，不是我方答复日（真机曾把 09-09 的提出记成 09-10 的答复日）。
- 数据修改类：没答完成也没拒绝 → 默认完成，时间取提出日期。
- 同一需求多个会话都出现 → 合并一行，提出人取最早提出者。
- 岗位复用：按行序向上找同提出人最近一次岗位值；**历史行只读**，填充只发生在起始行起的行。
  2026-09 起**内置在 `final`**（历史区 = `[表头行+1, 起始行-1]`，复用发生在写 `final_rows.csv` 之前），
  所以一轮写入即可完成；`position_reuse.py` 降级为独立补跑工具，两者共用 `common.reuse_position_fill`。
- 单元格只放业务结果：**备注等列一律不写判定依据/分析过程**。

## 解密（wcdb-key-tool，第三方灰色工具，不在 skill 内）
- 准备：`git clone https://github.com/TANGandXUE/wcdb-key-tool scripts/tools/wcdb-key-tool`
  或设环境变量 `WCDB_KEY_TOOL` 指向 `wcdb_key_tool_windows.py`；`common.find_wcdb_tool` 会自动找
  （含家目录深度≤3 搜索）。**注意**：工具放在更深的目录里就找不到，此时用环境变量显式指定最稳。
- ⚠️ 供应链注意：第三方灰色工具会随上游更新改变行为。clone 后记录所用 commit
  （`git -C scripts/tools/wcdb-key-tool rev-parse HEAD`），后续解密异常先核对是否工具版本漂移。
- **`extract` 必须显式传 `--db-dir`**：多账户机器上自动探测可能取到停用的那个账号（【真机】踩到）。
- 口令缓存 `~/.wcdb-key-tool/wechat-passphrase.json`：有效时无需重登；缓存丢失则先 `extract`（**微信须前台**）。
- 密钥缓存按账号隔离：`decrypt.py` 输出 `output/keys_<账号目录>.json`。
- 解密结果里 `.factory\...\*.db` 这类是**历史备份副本**，拿不到密钥属正常（会 SKIP），不影响主库。
- 合规：只读本机、数据不出电脑，使用前向用户确认。

## 两种表格来源（`excel.source`：local / kdocs）
两条通道算法完全一致（列映射 / 判定规则 / payload 结构 / 岗位复用算法都不变），只换读表与回填的 I/O；
差别只在**取数成本**：本地 openpyxl 读历史是免费的，云文档读历史太贵（见下「云文档读历史没有便宜路径」）。

### 通道 A：local（本地 .xlsx）
- 读：`probe.py workbook --workbook <xlsx>`（openpyxl）。
- 追加起始行 + **岗位复用**：`build_matrix_rows.py final --workbook <xlsx>`（一次读表两用；复用零额外成本）。
- 回填：editor_sdk（见下节）。

### 通道 B：kdocs（WPS 云文档）
**为什么必须多一层"快照"**：连接器的 MCP 工具**只有 agent 能调，Python 脚本调不到**。所以：
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

### ⚠️ 云文档"读历史列"没有便宜路径（2026-09 四条路逐条实测）
岗位复用需要读 `提出人 + 提出人岗位` 两列的历史。四条路都实测过，结论是**都不可行**：

| 路径 | 实测结果 |
|---|---|
| `drive.download_file` → curl | 返回 `url` 但**需登录态**，直接 curl 得 `403 {"result":"userNotLogin"}`。**云文档无法下载到本地再走本地通道** |
| `sheet.get_typed_value` | **跳过空单元格** → 行归属错位（见上） |
| `sheet.get_range_data` | 每条带 alignment/fonts/`cell_background_color`/`hasBorder`/`numFormat`，约 **450 B/格**；读 N+O 全列（≈382 格）≈ **200 KB** 上下文，且**没有关掉样式的参数** |
| `read_file(sheet_range=…)` | 返回**与 `get_range_data` 完全相同的带样式格式**（不是更精简的 markdown），一样贵。官方文档还给 xlsx 的 markdown 抽取标了「禁止」 |

另有 `sheet.find_range_data`（带 `filter`/`option_cols`）看起来像"按人名定点查"，
但 **`filter` 的结构未公开**：传错会被**静默忽略并返回整段**（比不查更贵）。
它顺带返回的 `option_col[].texts` 是**按列的值计数**（例：N 列 14 格为空、O 列"韩宁"出现 1 次），
可作廉价的存在性判断，但**不给行级配对**。

→ 因此**云文档通道不做表内岗位复用**，改由 `final --history <json>` 手动喂入（agent 确有把握时才读）。

接口契约（【真机】逐条实测）：
- 工作表参数名是 **`worksheet_id`**（不是 `sheetId`）。
- 选区坐标 camelCase：`rowFrom/rowTo/colFrom/colTo` + `opType`；日期格式走 `xf.numfmt`。
- **单次 `rangeData` 最多 100 条**，超出报 `length N exceeds limit 100` → `to_kdocs_payload.py` 自动分批。
- **每条 `formula` op 只能写一个值**，N 个不同值就是 N 条 op，无法靠合并收敛；能合并的只有
  `format`/`merge`/`picture` 这类区域操作。所以"6 行 × 18 列 ≈ 107 条 op"是正常量级。
- `get_range_data` 返回**稀疏**数组（行列 0-based，只含有值的单元格）；range 覆盖多格时按"合并单元格"
  语义整块填同一值。`build` 会重建成密集网格并裁掉尾部/右侧空行列。
- **写入必须用 `update_range_data`（幂等、按坐标），不要用 `sheet.add_row`**（非幂等，重试会插多行脏数据）。
- 写后**必须** `get_range_data` 回读核对，不能只信 `code: 0`。
- 限频 `429001`/`429002`：批量写要合并（脚本已压同列连续行），命中限频按响应给的恢复时间等待，别立刻重试。
- `.ksheet` 智能表格有字段类型（日期/单选等），写入形态与 `.xlsx` 不同。
- 文档若设区域保护（`sheet.list_protection_ranges`），写入会失败。
- 定位优先用**链接**（`get_share_info(link_id)`）；只给文件名要走 `search_files` 并**让用户确认**（同名风险）。
- `raw_*.json` / `kdocs_update*.json` 是中间产物，流程结束要清理。

## 表格回填（本地通道：editor_sdk / tencent-local-office-edit 技能）
1. 先 Skill 加载 tencent-local-office-edit 拿 edsdk.py 路径；读 sheet.md。
2. open_file → 按 `excel.sheet_match` 关键字匹配子表（不要写死年份）→ sheet_get_used_range 探边界；
   追加起始行用 `common.next_append_row(<xlsx>, sheet_hint, mapping)`。
3. 追加：sheet_insert_dimension(row, 末行+1, N) → sheet_set_range_value（每格必带 row/col/value_type
   + string_value|number_value）→ 对 date_columns 设 number_format_pattern → save_file。
4. ⚠️ 引擎保存会把旧 .xls 内容写成 xlsx：**保存后确认/改名为 .xlsx**，否则打开弹"格式与扩展名不匹配"。
5. ⚠️ 文件被宿主预览/用户打开（"Export file is occupied"）：save_file 到临时路径 → close_file(force)
   → 用 python 二进制覆写原文件内容（delete/rename 会被锁拒绝，写共享通常允许）；最后清理临时文件。
6. 保存后不要主动 close_file（用户可能在看）；需改名/重开时再 close。

## 易错点
- **绝不静默改写已有数据**：拿不到"追加起始行"就报错退出。【真机】暴露过三条这类路径
  （空表头映射、`final` 兜底第 2 行、`position_reuse` 默认扫整表），现全部改为显式报错或显式 opt-in。
- 名称重复联系人：contact 表一个 display 可能对应多个历史 username，取消息量最大者（export 已处理）。
- 用户在 config 里点名的会话，若 `--since` 后无消息要**如实报告**（【真机】遇到点名的群两周没发言），
  别默默产出空结果；让用户决定是否扩窗。
- 子代理判定出入可能很大：`merged_preview.csv` 必须给用户裁决（删行/合并）后再回填。
- 子代理输出 JSON 可能出现 `null` / 空串字段；脚本已做兜底，但裁决时也要扫一眼。
- 大批量写值别逐格调；payload JSON 用 `ensure_ascii=False` 保持中文可读。
- 会话很多时不要用 AskUserQuestion 逐条问；把清单写文件交用户勾选，或按关键字过滤。
