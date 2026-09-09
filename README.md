# wechat-worklog-matrix

通用「微信 PC 聊天记录 → 需求跟踪矩阵 Excel」流水线（Weixin 4.x / WeChat 4.1+）。

把客户/用户在微信里提出的系统需求，自动整理登记成需求跟踪矩阵台账：
**探测 → 一次性确认 → 解密导出 → LLM 识别判定 → 人工裁决 → 回填 Excel**。

不绑定任何具体客户、项目或人员——服务对象、项目、会话、处理人全部由使用者通过 `config.json` 配置。

## 特性

- **配置驱动，零内置假设**：账户 / Excel / 会话 / 时间范围 / 处理人全部由你选择；skill 内不出现任何真实账号、姓名、客户与项目标识。
- **列位置靠表头识别**：`probe` 读取目标 Excel 表头自动生成「表头 → 列」映射，换模板也能用，不写死 A~AA。
- **人在回路**：先探测一次性确认 → 子代理判定结果必须交人工裁决（删行/合并）后才回填。
- **单元格只写业务结果**：备注等列一律不写判定依据/分析过程。
- **合规边界**：解密使用第三方灰色工具 wcdb-key-tool（只读本机、数据不出电脑、使用前向用户确认）；聊天明文与探测中间产物全部留在本地并被 `.gitignore` 排除。

## 快速开始

```bash
# 0) 环境依赖
pip install openpyxl            # probe workbook / build final / n_reuse 依赖

# 1) 探测（账户 / 会话 / Excel 表头映射，一次跑完供一次性确认）
python pipeline.py probe --workbook <你的需求矩阵.xlsx> --dump probe.json

# 2) 确认探测结果，复制 config.example.json 为 config.json 并填写
#    （account / excel / people / scope / defaults）

# 3) 执行（解密 → 导出会话转录 → 分包给子代理）
python pipeline.py run

# 4) LLM 识别（用 references/agent-prompt-zh.txt 模板 + 各 agent_N.txt 分派给并行子代理）

# 5) 预览 → 人工裁决删行/合并 → 生成回填 payload
python scripts/build_matrix_rows.py preview --src <_out> --out merged_preview.csv
python scripts/build_matrix_rows.py final --preview merged_preview.csv --remove "..." --merge "a:b"
python scripts/position_reuse.py --workbook <xlsx> --out payload_n.json
```

解密工具 wcdb-key-tool（第三方，不在本仓库内）：

```bash
git clone https://github.com/TANGandXUE/wcdb-key-tool scripts/tools/wcdb-key-tool
# 或设置环境变量 WCDB_KEY_TOOL 指向 wcdb_key_tool_windows.py
```

## 目录结构

```
wechat-worklog-matrix/
├── SKILL.md                    # 技能主文档：流程/判定规则/关键事实/易错点
├── config.example.json         # 配置模板（复制为 config.json 后填写）
├── pipeline.py                 # 编排入口：probe（探测）/ run（解密→导出→分包）
├── scripts/
│   ├── probe.py                #   探测账户/会话/Excel 表头映射（支持 --json）
│   ├── decrypt.py              #   封装 wcdb-key-tool 解密（密钥按账号隔离）
│   ├── export_conversations.py #   按 config 导出指定会话转录
│   ├── split_for_agents.py     #   转录按文件大小均衡分 N 份给子代理
│   ├── build_matrix_rows.py    #   preview（去重标记+裁决）/ final（生成回填 payload）
│   ├── position_reuse.py       #   岗位列向上复用计划（原 n_reuse_plan.py）
│   ├── common.py               #   配置加载 / 路径识别 / 日期与列工具
│   └── smoke_test.py           #   自检：import + 纯函数断言
└── references/
    ├── agent-prompt-zh.txt     # 需求识别子代理提示词模板
    └── workflow-notes.md       # 工作流要点与坑位备忘
```

## 自检

修改代码后运行：

```bash
python scripts/smoke_test.py        # import 全模块 + 纯函数断言
python scripts/smoke_test.py --full # 额外校验依赖 openpyxl 的模块
```

## 许可

本项目采用 MIT 许可，详见 [LICENSE](LICENSE)。

> 注意：本 skill 依赖的解密工具 wcdb-key-tool 为第三方项目，不在本仓库内，其授权与合规性请自行评估。

## 合规与隐私说明

- 本 skill 不采集、不上传任何聊天数据；探测/导出/解密产物均在本机，且已被 `.gitignore` 排除（`wechat_pilot/`、`*.db`、`keys_*.json`、`config.json`）。
- 仓库内不包含任何真实账号目录、姓名、微信号、客户名或项目号。
- wcdb-key-tool 属第三方灰色工具：使用前需向用户确认，只读本机、数据不出电脑。
