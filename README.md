# wechat-worklog-matrix

通用「微信 PC 聊天记录 → 需求跟踪矩阵」流水线（Weixin 4.x / WeChat 4.1+）。

把客户/用户在微信里提出的系统需求，自动整理登记成需求跟踪矩阵台账：
**探测 → 一次性确认 → 解密导出 → LLM 识别判定 → 人工裁决 → 回填台账**。

不绑定任何具体客户、项目或人员——服务对象、项目、会话、处理人全部由使用者通过 `config.json` 配置。

当前版本 [`v1.2.0`](https://github.com/yfwang0731/skill-wechat-worklog-matrix/releases/tag/v1.2.0)　·　变更历史见 [`CHANGELOG.md`](CHANGELOG.md)

[![selftest](https://github.com/yfwang0731/skill-wechat-worklog-matrix/actions/workflows/selftest.yml/badge.svg)](https://github.com/yfwang0731/skill-wechat-worklog-matrix/actions/workflows/selftest.yml)

## 特性

- **配置驱动，零内置假设**：账户 / 表格 / 会话 / 时间范围 / 处理人全部由你选择；仓库内不出现任何真实账号、姓名、客户与项目标识。
- **两种表格来源**（`excel.source`）：`local` 本地 `.xlsx`（openpyxl 读、editor_sdk 回填）；
  `kdocs` **WPS 云文档**（只用给文档链接/名字，agent 直接读写，不必把文件下到本地）。
  两条通道共用同一套列映射、判定规则与 payload 结构，只换 I/O 两端。
- **列位置靠表头识别**：`probe` 读目标表头自动生成「表头 → 列」映射，换模板也能用，不写死 A~AA。
- **消息正文解码到位**：微信把长文本/图片/引用/合并转发等以 ZSTD 压缩存储，
  装 `zstandard` 可完整还原（缺库时给出明确标记而非乱码）；`[链接/文件]` 会提取标题与**引用原文**，
  对方用「引用回复」提需求时不再丢内容。
- **合规边界**：解密用第三方灰色工具 wcdb-key-tool（只读本机、数据不出电脑、使用前向用户确认）；
  聊天明文与探测中间产物全部留在本地并被 `.gitignore` 排除。

> 运行纪律（拿不到关键输入为何必须报错退出、人工裁决为何不能跳过、单元格里只许放什么）见「核心原则」（`SKILL.md`）——那三条都是历史上真栽过的地方。

## 安装

本技能是 WorkBuddy 的 skill，装到技能目录即可被识别：

```bash
# 方式 A：clone 到用户级技能目录（所有项目都能用）
#   Windows 的技能目录：%USERPROFILE%\.workbuddy\skills\
git clone https://github.com/yfwang0731/skill-wechat-worklog-matrix.git \
  ~/.workbuddy/skills/wechat-worklog-matrix

# 方式 B：网页右上 Code → Download ZIP，解压到同一位置
```

- **只给单个项目用**：放到 `<项目根>/.workbuddy/skills/`；**目录名请用 `wechat-worklog-matrix`**
  （与 `SKILL.md` 里的 `name` 一致），不要保留 `skill-` 前缀、也不要多套一层目录。
- 装好后确认路径形如 `…/skills/wechat-worklog-matrix/SKILL.md`；更新已安装的技能：
  `git -C ~/.workbuddy/skills/wechat-worklog-matrix pull`。

**解密工具** wcdb-key-tool（第三方，不在本仓库内；授权与合规性请自行评估）：

```bash
git clone https://github.com/TANGandXUE/wcdb-key-tool scripts/tools/wcdb-key-tool
# 或设置环境变量 WCDB_KEY_TOOL 指向 wcdb_key_tool_windows.py
```

## 快速开始

```bash
pip install zstandard openpyxl   # zstandard 建议装（缺库会漏掉压缩正文里的需求）；openpyxl 仅本地表格通道需要
python pipeline.py --config config.json probe --dump probe.json   # 探测账户 / 表头映射
python pipeline.py --config config.json run                       # 解密 → 导出转录 → 分包
python scripts/build_matrix_rows.py preview --src <output.dir>/transcripts/_out --out merged_preview.csv
python scripts/build_matrix_rows.py final --preview merged_preview.csv --start-row-excel <起始行>
```

- **配置**：复制 `config.example.json` 为 `config.json` 后填写；`excel.column_mapping` 与 `excel.header_row`
  必须用 `probe` 的 `--json` 输出粘进去，**不要手填**（`final` 靠它们拿列位）。
- **会话清单要等解密之后才有** → 先解密、再探一次（见「交互式流程」第 1 步）。
- `preview` 的产出**必须先交用户裁决**（删行/合并）才能 `final`；回填按 `config.excel.source` 分本地 / 云文档两路。
- **完整的五步流程、两条通道的读表与回填命令、参数速查**见「交互式流程」（`SKILL.md`）；
  云文档通道的取数、回填与接口约束另见 [`references/kdocs-channel.md`](references/kdocs-channel.md)。

## 开发与仓库

### 目录结构

```
wechat-worklog-matrix/
├── SKILL.md
├── CHANGELOG.md
├── config.example.json
├── pipeline.py
├── test-prompts.json
├── LICENSE
├── .gitignore
├── .gitattributes
├── .github/workflows/
│   └── selftest.yml
├── scripts/
│   ├── probe.py
│   ├── sheet_snapshot.py
│   ├── to_kdocs_payload.py
│   ├── decrypt.py
│   ├── export_conversations.py
│   ├── split_for_agents.py
│   ├── build_matrix_rows.py
│   ├── position_reuse.py
│   ├── rules_check.py
│   ├── common.py
│   └── smoke_test.py
└── references/
    ├── agent-contract.json
    ├── agent-prompt-zh.txt
    ├── kdocs-channel.md
    ├── rules-fixtures.json
    └── workflow-notes.md
```

各脚本干什么见「脚本清单」（`SKILL.md`）；`references/workflow-notes.md` 是机制说明与实测依据层。

### 自检

```bash
python scripts/smoke_test.py        # 33 项：import / 纯函数 / CLI 契约 / 输入路径守卫 / 落表前复核 / 子代理产出契约 / 快照覆盖与表内复用 / frontmatter 额度 / 文档分层与引用结构 / 离线端到端 / 仓库卫生 / 临时目录清理 / 编码
python scripts/smoke_test.py --full # 再加 3 项依赖 openpyxl、zstandard 的检查
```

全程**不碰真实微信数据与网络**，可随时执行；失败以退出码 1 结束。覆盖重点在那些
"**不会报错、却会改错数据**"的路径：缺关键输入必须报错退出、列映射靠表头识别、跨会话去重、
岗位列只填空缺、云文档请求体分批与键名、Windows 非 UTF-8 控制台下不能崩。
2026-09 在一台 Windows + 微信 4.1 + WPS 云文档上完整实跑过，修的就是这一批路径。

CI（[`.github/workflows/selftest.yml`](.github/workflows/selftest.yml)）跑的就是它 —— 本地与 CI
**同一个入口**，避免"CI 绿、本地跑不到"或反之，共 4 格：

| 轴 | 取值 | 为什么 |
|---|---|---|
| Python | 3.9 / 3.13 | 3.9 是声明的最低支持版本，跨代测避免"你本地能跑、用户装不上库" |
| 依赖 | minimal / **full** | minimal 守"零依赖也能跑 + 缺库给可执行提示"；full 才跑得到本地通道读表与 ZSTD 解压 |

> **矩阵只有 Windows 一轴** —— 主链（微信 PC 客户端 + `wcdb-key-tool`）本就只在 Windows 上有意义，
> 留一个 ubuntu 格子等于对外承诺"支持 Linux"，而那条路根本跑不通。它原先唯一独有的覆盖是
> **大小写敏感文件系统**（文档里文件名写错大小写，Windows / macOS 会判为存在），现已下沉到自检内部，
> 本地也跑得到 —— 这比只在 CI 那一格可跑更强。
>
> CI 里还**故意不设** `PYTHONUTF8` / `PYTHONIOENCODING` —— 设了就把 Windows 的 cp1252 缺陷盖住，
> 那格等于白跑。脚本靠 `common.ensure_utf8_stdio()` 自己把标准流切到 UTF-8。

## 许可与合规

本项目采用 MIT 许可，详见 [LICENSE](LICENSE)。

- 本 skill 不采集、不上传任何聊天数据；探测 / 导出 / 解密 / 回填的产物**全部留在本机**，
  且已被 `.gitignore` 排除。**完整的忽略清单以 [`.gitignore`](.gitignore) 为准**（它会随新功能更新），
  含聊天明文与个人标识的几类（`wechat_pilot/`、`*.db`、`keys_*.json`、`config.json`、云文档中间物等）都在其中。
- 仓库内不包含任何真实账号目录、姓名、微信号、客户名或项目号。
- wcdb-key-tool 属第三方灰色工具，不在本仓库内：使用前需向用户确认，只读本机、数据不出电脑；
  其授权与合规性请自行评估。
