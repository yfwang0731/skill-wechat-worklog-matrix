# 工作流要点与坑位备忘（微信 4.x → 需求跟踪矩阵，通用）

> 本文件不记录任何具体客户/人员/项目信息，全部以占位符或配置键描述。

## 环境事实
- 微信 4.1.x；进程 Weixin.exe。本地库在 `~/Documents/xwechat_files/<账户目录>/db_storage/`，SQLCipher 加密。
- **一台机器可能有多个账户目录**，必须由用户选择；`common.find_db_storage(preferred, hint)` 支持按账号目录特征串优选，`--db-storage` 可显式指定。
- 消息库分片：`db_storage\message\message_1/2/3.db`（时间排序 m2 < m1 < m3）；每会话消息表 `Msg_<md5(wxid)>`。
- 发送者数字 id = 分片内 `Name2Id.rowid`（**跨分片会变**），不要跨分片复用。
- `create_time` 秒级时间戳；`local_type` 低 32 位为消息类型（1=文字，3=图片，34=语音，43=视频，47=表情，49=链接/文件）。
- 会话范围/时间范围/处理人/对方关键字全部来自 `config.json`，不在代码里内置。

## 列映射（关键改进）
- **列位置靠"表头名"识别**：`probe.py workbook` 读表头，按 `COLUMN_ALIASES`（逻辑列名 → 可能的表头名）匹配，
  产出 `column_mapping`（如 "需求描述"→"L"）。换模板也能用，不再写死 A~AA。
- 匹配不到的表头列不写入（会打印未识别清单，不影响结果）。
- `date_columns` 里的列以 Excel 序列号写入，并需设 `yyyy-mm-dd` 格式。

## 判定规则（用户口径，可由 config.rules 开关）
- 对方提出系统需求才算；寒暄/通知不算。拒绝或无实质回复 → 不写行。
- 明确答"完成" → 计划时间=完成时间=答复日期。数据修改类：没答完成也没拒绝 → 默认完成，时间取提出日期。
- 同一需求多个会话都出现 → 合并一行，提出人取最早提出者。
- 岗位复用：按行序向上找同提出人最近一次岗位值；有则复用，**只填空缺、不覆盖更具体值**。
  用 `scripts/position_reuse.py --workbook <xlsx>` 生成回填 payload。
- 单元格只放业务结果：**备注等列一律不写判定依据/分析过程**。

## 解密（wcdb-key-tool，第三方灰色工具，不在 skill 内）
- 准备：`git clone https://github.com/TANGandXUE/wcdb-key-tool scripts/tools/wcdb-key-tool`
  或设环境变量 `WCDB_KEY_TOOL` 指向 `wcdb_key_tool_windows.py`；`common.find_wcdb_tool` 会自动找（含家目录深度≤3 搜索）。
- ⚠️ 供应链注意：第三方灰色工具会随上游更新改变行为。clone 后记录所用 commit 到 config 或本地备注
  （`git -C scripts/tools/wcdb-key-tool rev-parse HEAD`），后续解密异常先核对是否工具版本漂移。
- 口令缓存 `~/.wcdb-key-tool/wechat-passphrase.json`：缓存有效时无需重登微信；缓存丢失则先 `extract`（微信须前台）。
- 密钥缓存已按账号隔离：`decrypt.py` 输出 `output/keys_<账号目录>.json`，切换账号不会复用错 keys；
  若手工调用 wcdb-key-tool 请同样按账号区分 keys 文件。
- 封装脚本 `scripts/decrypt.py`：extract → decrypt；详见其 docstring。
- 合规：只读本机、数据不出电脑，使用前向用户确认。

## 两种表格来源（`excel.source`：local / kdocs）
**两条通道算法完全一致**（列映射 / 判定规则 / payload 结构 / 岗位复用逻辑都不变），只换读表与回填的 I/O。用户可能给本地文件，也可能只说"改 WPS 上那个文档"。

### 通道 A：local（本地 .xlsx）
- 读：`probe.py workbook --workbook <xlsx>`、`position_reuse.py --workbook <xlsx>`（openpyxl）。
- 追加起始行：`build_matrix_rows.py final --workbook <xlsx>`。
- 回填：editor_sdk（见下节）。

### 通道 B：kdocs（WPS 云文档）
**为什么必须多一层"快照"**：连接器的 MCP 工具**只有 agent 能调，Python 脚本调不到**。所以：
- 读：agent 调 `sheet.get_range_data` → 原始返回**原样**落 `raw_*.json` → `sheet_snapshot.py build` 重建密集网格快照 → 脚本读快照。
- 写：脚本产出 payload → `to_kdocs_payload.py` 转 `rangeData` → agent 调 `sheet.update_range_data`。

读表**分两趟窄读**，别把整表搬进上下文：
1. 先 `sheet.get_sheets_info` 拿工作表清单 / `sheetId` / 已用区域上限。
2. 第 1 趟只读前 3 行 × 全部列 → 识别表头（目的是拿到 column_mapping）。
3. 第 2 趟按映射只读关键列（需求描述/提出时间/提出人/提出人岗位/解决人…）整表 → 算末数据行与岗位复用。

坑位：
- `sheet.get_range_data` 返回**稀疏**数组（行列 0-based，只含有值的单元格）；`build` 会重建成密集网格并裁掉尾部/右侧空行列。range 覆盖多格时按"合并单元格"语义整块填同一值。
- 单元格取值优先 `cellText`，公式单元格退回 `originalCellValue`；`numFormat` 会存进快照的 `num_formats` 备用。
- **写入必须用 `sheet.update_range_data`（幂等、按坐标），不要用 `sheet.add_row`**——add_row 非幂等，重试/重复调用会插多行脏数据。
- 写后**必须** `sheet.get_range_data` 回读同一区域核对，不能只信 `code: 0`。
- 限频 `429001`/`429002`：批量写要合并成一次请求（`to_kdocs_payload.py` 已把同列连续行压成一段），命中限频按响应给的恢复时间等待，别立刻重试。
- 日期：与本地通道一致写 Excel 序列号，并额外补 `format` op（`numfmt=yyyy-mm-dd`，见 `excel.kdocs.date_numfmt`）；`--no-date-format` 可只写值。
- `.ksheet` 智能表格有字段类型（日期/单选等），写入形态与 `.xlsx` 不同，需先确认文档类型。
- 文档若设区域保护（`sheet.list_protection_ranges`），写入会失败，先让用户解除。
- 定位优先用**链接**（`get_share_info(link_id)`）；只给文件名要走 `search_files` 并**让用户确认**（同名风险）。
- `raw_*.json` / `kdocs_update*.json` 是中间产物，流程结束要清理。

## 表格回填（本地通道：editor_sdk / tencent-local-office-edit 技能）
1. 先 Skill 加载 tencent-local-office-edit 拿 edsdk.py 路径；读 sheet.md。
2. open_file → 按 `excel.sheet_match` 关键字匹配子表（不要写死年份）→ sheet_get_used_range 探边界；
   追加起始行用 `common.next_append_row(<xlsx>, sheet_hint, mapping)`。
3. 追加：sheet_insert_dimension(row, 末行+1, N) → sheet_set_range_value（每格必带 row/col/value_type
   + string_value|number_value）→ 对 date_columns 设 number_format_pattern=yyyy-mm-dd → save_file。
4. ⚠️ 引擎保存会把旧 .xls 内容写成 xlsx：**保存后确认/改名为 .xlsx**，否则打开弹"格式与扩展名不匹配"。
5. ⚠️ 文件被宿主预览/用户打开（"Export file is occupied"）：save_file 到临时路径 → close_file(force)
   → 用 python 二进制覆写原文件内容（delete/rename 会被锁拒绝，写共享通常允许）；最后清理临时文件。
6. 保存后不要主动 close_file（用户可能在看）；需改名/重开时再 close。

## 易错点
- 名称重复联系人：contact 表一个 display 可能对应多个历史 username，取消息量最大者（export 已处理）。
- 子代理判定出入可能很大：merged_preview.csv 必须给用户裁决（删行/合并）后再回填。
- 大批量写值别逐格调；payload JSON 用 ensure_ascii=False 保持中文可读。
- 会话很多时不要用 AskUserQuestion 逐条问；把清单写文件交用户勾选，或按关键字过滤。
