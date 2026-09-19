# WPS 云文档通道（读表 / 回填 / 接口约束）

`config.excel.source = kdocs` 时的**详情层**：读表分两趟取数、回填请求体怎么造、以及写错就失败的接口约束。
两条通道共用同一套列映射、判定规则与 payload 格式，只有「读表 / 回填」两端的 I/O 不同 ——
**规则与流程见 `SKILL.md`「表格来源：两条通道」**，本文只写这一端的做法与实测依据
（机制与逐条实测另见 `references/workflow-notes.md`）。

> 云文档之所以要**经过「快照」**：连接器的 `sheet.*` 都是**只有 agent 能调的 MCP 工具，Python 脚本调不到**。
> 所以读表由 agent 取数落盘、脚本再读快照；写表由脚本产出请求体、agent 去调工具。
> **不要把连接器当成脚本可 import 的库**，也不要在脚本里直连它。

## WPS 通道读表（分两趟，避免整表进上下文）

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

## WPS 通道回填

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

## 云文档接口约束（实测，写错即失败）

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
