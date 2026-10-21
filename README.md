> **本插件全程由 DeepSeek Harness（DSH）+ DeepSeek V4.1 Flash 开发与维护：100% 纯 AI 生成，人工代码 0 添加。**

# dsh-doc-review

DSH 侧栏批注插件。在 Word/PDF 预览中选中文字写修改意见，可连续添加多条，一次提交给 agent；agent 通过附带的 `docx_edit.py` 完成定位、备份、修改、回读与原子替换。Excel 支持按单元格地址添加批注。

批注只描述“改什么”，不直接编辑文件；格式保真由原文件编辑脚本负责。

## 功能

- **Word / PDF**：右侧栏预览中选中文字，浮出“添加批注”，输入意见；可连续添加多条，最后一并提交。
- **Excel**：手填工作表、单元格地址（如 `B3:F8`）和意见；可附当前值，供 agent 用“地址 + 当前值 + 相邻单元格”定位。
- 批注按文件绝对路径保存在浏览器 `localStorage`；重开标签或刷新后恢复列表，并按页码 + 引文重新定位高亮。
- 提交时把全部批注拼成一条消息，通过当前会话的 `session.prompt` 入队。
- agent 修改原文件后，预览监听文件变化自动刷新，不需要重新打开。
- 选中范围用 CSS Custom Highlight 标黄，不改动预览 DOM。

## 工作原理

1. 插件在预览头部注册“批注”按钮；只有文件已在侧边栏打开时才会出现。
2. 提交时把批注拼成一条带任务协议的消息（当前 `v1.6`），协议限定在操作层，内容与文风仍交给当前模式。
3. agent 按扩展名选择 helper：docx 用 `docx_edit.py`，xlsx/xlsm 用 `xlsx_edit.py`，一次完成 hash、备份、扫描、修改、回读与原子替换；helper 不支持的结构直接返回 `unsupported`，不猜。
4. 修改完成后预览自动刷新；消息声明预览已打开，因此默认不调用 `present`。

协议与模式无关：standard、文稿或自定义 preset 都能用，不依赖任何技能或人格。

## 安装

### 要求

- DSH 0.2.x（Web 或桌面版 profile）
- Python 3.10+，并安装 `python-docx`（`python3 -m pip install python-docx`）
- macOS。预览依赖 DSH 内置的 PDF/Office 预览；其他平台未验证

### 安装到 Web profile

```bash
DSH_PROFILE_DIR="$HOME/.dsh/profiles/web" bash install.sh
```

### 安装到桌面版 profile

```bash
DSH_PROFILE_DIR="$HOME/.dsh/profiles/desktop" bash install.sh
```

脚本会：

- 把本目录链接进 profile 的 `node_modules`；
- 在 `package.json` 的 `dsh.profile.bundles` 末尾追加 `dsh-doc-review`（改前备份）；
- 把 `docx_edit.py`、`doc_review_log.py` 安装到 `~/.dsh/doc-review/bin/`。

安装后需要重启对应的 DSH（Web 版重启 `dsh web`，桌面版退出重开），client bundle 只在启动时加载。

## 使用

1. 在右侧栏打开 docx/pdf：Files 标签、交付卡片或对话里的文件链接都可以。
2. 预览头部出现“批注”。
3. Word/PDF：点“批注”→ 在预览中选中文字 → 浮出“添加批注” → 输入意见 → 保存。
4. Excel：点“批注”→ 填工作表、单元格、当前值（可选）、意见 → 加入批注。
5. 面板里核对多条意见，点“提交 N 条”。

提交的消息形如：

```
【批注修改任务 · dsh-doc-review v1.6】
文件：/path/to/文件.docx
批注：2 条

1. [第 3 页]
   引文：“……”
   意见：改为……

预览：提交时已在右侧栏打开，修改后自动刷新；完成后不要调用 present，报告里给出文件路径即可。

执行约定（仅约束本次文件修改的操作方式；内容与文风仍遵循当前模式，安全、权限与审批遵循宿主）：
0. 需求已由批注确认，不重复提纲、方案确认或先例检索。
1. 只改批注涉及的内容，不顺手优化其他段落；只读目标文件……
```

## docx_edit.py

声明式修改 docx，替代一次性脚本。

```bash
docx_edit.py --file "目标.docx" --ops '[{"op":"replace","find":"一、建设要点","replace":"一、建设目标"}]'
```

| op | 参数 |
|---|---|
| `replace` | `find`、`replace`、`all`、`occurrence` |
| `delete` | `find`、`all`、`occurrence` |
| `set_font` | `find`、`font`、`size`、`runs`（`all`/`matched`） |
| `set_color` | `find`、`color`（RRGGBB）、`runs` |
| `set_paragraph_style` | `find`、`style` |

内部完成 hash、备份、结构扫描、定位、修改、回读校验与原子替换；只允许正文段落与表格，页眉页脚、文本框、修订痕迹、域、内容控件一律拒绝。退出码：0 成功，2 未找到/歧义，3 不支持，4 校验失败，5 文件/异常。

## xlsx_edit.py

声明式修改 xlsx/xlsm，替代一次性脚本。

```bash
xlsx_edit.py --file "台账.xlsx" --ops '[{"op":"set_cell","sheet":"数据","cell":"B2","value":"999","expect":"100"}]'
```

| op | 参数 |
|---|---|
| `set_cell` | `sheet`、`cell`、`value`、`expect`（现值哨兵）、`value_type`（`auto`/`string`/`number`/`bool`/`formula`/`date`） |
| `replace_text` | `sheet`、`find`、`replace`、`range`、`all`、`occurrence` |
| `set_format` | `sheet`、`cell` 或 `range`、`font`、`size`、`color`、`bold`、`italic`、`underline`、`number_format`、`align`、`wrap` |

`value_type: auto` 的规则：`=` 开头按公式；`12%` 按数值 0.12；纯数字按数值（前导零如 `001` 保持文本）；`TRUE/FALSE` 按布尔；日期格式单元格 + ISO 日期按日期；其余按文本。

内部完成预检、hash、备份、定位、修改、回读校验、zip 部件守恒检查与原子替换：

- 除被操作单元格外，所有公式必须逐一保持不变，否则回滚；
- 存盘后任何部件消失（透视表、切片器、VBA 等）即放弃写入；
- 合并非左上角、数组公式、加密或损坏文件直接拒绝，并给出可用表名或左上角地址；
- 受保护工作表默认拒绝，返回 `sheet_protected`；用户确认后才能加 `--allow-protected` 重试。

退出码与 docx helper 一致：0 成功，2 未找到/歧义/现值不符，3 不支持/受保护，4 校验失败，5 文件/异常。

## 失败与回退日志

辅助脚本与 agent 的失败、回退、人工介入会写入：

```
~/.dsh/doc-review/logs/events-YYYY-MM.jsonl
```

沙箱禁止写 `~/.dsh` 时，自动回退到当前工作区的 `.dsh-doc-review/logs/`。

事件类型：`helper_ok`、`helper_fail`、`fallback_script`、`manual_intervention`、`verify_fail`。字段含时间、会话、文件、hash、阶段、错误码、原因、操作摘要；不写文档全文，引文截断 120 字。

```bash
python3 ~/.dsh/doc-review/bin/doc_review_log.py summary
python3 ~/.dsh/doc-review/bin/doc_review_log.py tail --n 20
```

## 数据位置

- 插件：profile 的 `node_modules/dsh-doc-review`（指向本仓库）；
- 辅助脚本：`~/.dsh/doc-review/bin/`；
- 日志：`~/.dsh/doc-review/logs/`，或工作区 `.dsh-doc-review/logs/`；
- 批注草稿：浏览器 `localStorage`，键为 `dsh-doc-review:v1:<绝对路径>`。

## 限制

- 只处理 docx 正文段落与表格；页眉页脚、文本框、修订、域、内容控件拒绝修改。
- Excel 只做单元格值、单元格内文本替换与格式；行列增删、跨表引用迁移不支持。
- openpyxl 存盘会清空公式缓存值（已设 `fullCalcOnLoad`，Excel/WPS 打开时重算）；预览里公式可能显示空白。
- 含图表/图片的工作簿会被 openpyxl 重新序列化绘图部件，部件守恒但版式可能有细微变化。
- `.xlsm` 以 `keep_vba=True` 处理；`.xls`、`.csv` 需先另存为 xlsx。
- Excel 需要手填单元格地址，预览的选中状态尚未暴露给插件。
- 扫描版 PDF 没有文字层，选不中。
- 批注按绝对路径持久化，跨文件不迁移；文件内容大改后引文可能定位失败。
- 不跟踪“已解决”状态，以 agent 的修改报告为准。
- 依赖预览插件的 DOM 标记（`data-pdf-preview`、`data-pdf-page`、`data-textpreview-body`），上游大改时需要同步。
- 默认沙箱下 central 日志可能不可写，日志会分散在各工作区的回退目录。

## 测试

```bash
python3 -m unittest discover -s tests -v
```

`tests/make_fixtures.py` 生成含公式、图表、图片、合并、数据验证、条件格式、批注、表格、定义名称、保护与隐藏工作表的工作簿；`tests/test_xlsx_edit.py` 覆盖写入、现值不符、歧义、合并区、格式、受保护表、dry-run、无变化、日志与部件守恒。

## 卸载

```bash
cp ~/.dsh/profiles/web/package.json.bak-<时间戳> ~/.dsh/profiles/web/package.json
rm ~/.dsh/profiles/web/node_modules/dsh-doc-review
```

桌面版同理，替换为 `profiles/desktop`。重启 DSH 后生效；`~/.dsh/doc-review` 可按需保留或删除。

## 版本与变更

- 插件版本：`0.2.0`
- 批注任务协议：`v1.6`
- 变更记录：GitHub Releases

## 许可证

MIT
