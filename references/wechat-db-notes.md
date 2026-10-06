# 微信库与消息解码（环境事实 / 解密 / 消息正文解码） （随 skill 发布｜禁写本机路径、账号、uid、用户数据、运行期结果）

> 本文件是依据层：写**为什么这样**（实测数字、踩坑经过）。规则本体逐条在 `SKILL.md`；
> 通道 I/O 依据在 `references/workflow-notes.md`。本文件只负责其主题 —— 微信库结构、解密与消息正文解码。

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
