---
name: officecli
description: 当任务以 Word、Excel 或 PowerPoint 文件作为输入或输出，或需要将 CSV、TSV 数据导入 Excel 时使用；支持创建、读取、编辑、排版、诊断和验证 .docx、.docm、.doc、.xlsx、.xlsm、.xls、.pptx、.pptm、.ppt 文件。
---

# OfficeCLI 文件处理

## 选择程序

根据运行平台直接使用本技能目录中的对应程序：

| 平台 | 程序 |
|---|---|
| Apple Silicon Mac | `<本技能绝对目录>/bin/officecli-mac-arm64` |
| Intel Mac | `<本技能绝对目录>/bin/officecli-mac-x64` |
| Linux ARM64（glibc） | `<本技能绝对目录>/bin/officecli-linux-arm64` |
| Linux x64（glibc） | `<本技能绝对目录>/bin/officecli-linux-x64` |
| Windows ARM64 | `<本技能绝对目录>\bin\officecli-win-arm64.exe` |
| Windows x64 | `<本技能绝对目录>\bin\officecli-win-x64.exe` |

macOS 和 Linux 根据 `uname -m` 选择程序：`arm64`、`aarch64` 使用 ARM64，`x86_64`、`amd64` 使用 x64。Linux 程序用于 glibc 环境，不用于 Alpine 等 musl 环境。将程序绝对路径赋给 `OCLI`。

Windows PowerShell 使用 `[System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture` 判断 `Arm64` 或 `X64`，将对应程序的绝对路径赋给 `$OCLI`，使用 `& $OCLI ...` 执行。其他平台当前不支持。

运行前设置以下环境变量，并确保它们对本次任务的每一条 OfficeCLI 命令生效：

```bash
export OFFICECLI_SKIP_UPDATE=1 OFFICECLI_NO_AUTO_INSTALL=1 OFFICECLI_NO_AUTO_RESIDENT=1
```

```powershell
$env:OFFICECLI_SKIP_UPDATE = "1"
$env:OFFICECLI_NO_AUTO_INSTALL = "1"
$env:OFFICECLI_NO_AUTO_RESIDENT = "1"
```

`OFFICECLI_NO_AUTO_RESIDENT=1` 只禁止自动创建会话，不会关闭显式执行 `open` 创建的会话。程序不存在或不可执行时停止并报告具体路径。

## 读取格式规则

根据任务涉及的格式读取对应文件，再开始操作：

- `.docx`、`.docm`：读取 `references/docx.md`；涉及可填写字段、内容控件、邮件合并或文档保护时，同时读取 `references/docx-forms.md`。
- `.xlsx`、`.xlsm`、CSV、TSV：读取 `references/xlsx.md`。
- `.pptx`、`.pptm`：读取 `references/pptx.md`。
- `.doc`、`.xls`、`.ppt`：先读取 `references/legacy-conversion.md` 完成转换，再按转换结果和任务内容读取上述对应文件。
- 同时处理多种格式：读取所有对应文件。

## 宏启用文件

- 对已有 `.docm`、`.xlsm`、`.pptm` 直接执行读取或原位编辑命令，并保留原扩展名。不要对这些格式使用 `create`、`dump` 或 `merge`，避免重建文件时丢失 VBA 工程。
- 需要生成另一份输出时，先复制整个源文件并保持相同扩展名，再编辑该副本；不要创建空白文件后搬运内容。
- 不执行宏，不修改 `vbaProject` 或相关部件。完成后执行 `close` 和 `validate`，并在目标 Office 应用中确认文件、宏和签名状态符合要求；结构校验不能证明宏可正常运行或数字签名仍然有效。

## 工作流程

1. 明确输入、输出、受众、语言、版式和目标查看器。
2. 编辑现有文件前使用 `view`、`get` 或 `query` 检查结构和内容。
3. 属性不确定时运行 `"$OCLI" help <格式> <元素>`；需要确认某个命令支持哪些属性时运行 `"$OCLI" help <格式> <命令> <元素>`。以帮助输出为准，不猜测属性。
4. 按“语义化 DOM 操作 → 必要时原始 XML”的顺序修改。
5. 先用少量操作验证结构，再使用 `batch` 执行大量同类操作；检查退出状态、JSON `success`、`warnings`、失败项和回滚状态，不能只统计成功项。
6. 默认不调用 `open`；直接执行修改命令或使用 `batch`，让每条命令自行保存并释放文件。
7. 最后一项修改完成后先执行 `close`，再执行结构、格式和视觉检查；修复问题后重新关闭并执行完整检查。

## 会话规则

- 将非驻留模式作为默认方式。只有大量连续操作存在明确性能收益时才调用 `open`。
- 一旦 `open` 成功，无论任务正常完成、命令失败还是提前停止，都必须对同一文件执行 `close`。
- `save` 只刷新磁盘并保留会话，不能作为任务的收尾命令。需要中途生成外部检查快照时才使用 `save`。
- 对每个创建、修改或显式打开的文件，在最终检查前无条件执行一次 `"$OCLI" close "$FILE"`。没有活动会话时 `close` 是安全的幂等操作。
- 同时处理多个文件时逐一关闭，不得因为其中一个文件失败而跳过其他文件。
- `save` 或 `close` 失败，或程序报告驻留进程遗留了未刷新修改时，立即停止继续写入，报告文件路径和错误；重新读取磁盘文件确认状态前，不得声称修改已保存。
- 驻留进程存活但命令无法送达时会返回退出码 `3`。先重试一次；仍失败则执行 `close`。`close` 失败时不得绕过该会话直接启动另一写入流程，以免并发覆盖。

## 批处理规则

- `batch` 默认原子执行：各项仍会返回结果，但任意一项失败都会回滚整批修改。使用 `--json` 并检查失败项的 `code` 和汇总中的 `atomicRolledBack`。
- 每个批处理项都可能带有自己的 `warnings`；退出状态为 `0` 也不能代替逐项检查。批处理中的 `query`、`get` 等读取项需要不同输出深度或格式时，在各项中分别设置相应选项。
- 只有任务明确允许保留部分成功结果时才使用 `--best-effort`。需要遇到首个错误立即停止且保留此前成功项时，同时使用 `--best-effort --stop-on-error`。
- 单独使用 `--stop-on-error` 只会提前停止，默认仍会整体回滚。
- `--force` 只用于绕过 Word 文档保护，不是批处理错误恢复选项。
- 批处理失败后重新读取目标结构，确认文件处于预期状态，再决定修正命令并重试或改用 `--best-effort`。

## 退出状态

- `0`：操作成功落地；仍需检查 Batch 子项的 `warnings`。
- `1`：操作未完成。读取 JSON `error.code`、`suggestion` 和 `validValues`，修正后再执行。
- `2`：命令已运行但存在警告，例如不支持的属性被丢弃或校验器发现问题。它既可能是 `success=true` 的部分成功，也可能是 `success=false` 的无修改；必须读取 JSON `success` 和 `warnings`，再读取目标确认实际状态。不得把退出码 `2` 当作成功交付。
- `3`：显式驻留会话仍存在但无法接收命令，按会话规则处理。

出现 `unsupported_property` 时先查询对应元素或命令帮助。若部分属性已经落地，先读取目标，再用正确属性补齐；不要在未知状态下盲目重放整条命令。

## 常用命令

| 命令 | 功能 | 典型用途 |
|---|---|---|
| `help <格式> [命令] <元素>` | 查看元素或命令可用的属性、取值格式和示例 | 属性或语法不确定时先查询 |
| `create <文件>` | 根据扩展名创建空白 Office 文件 | 创建 `.docx`、`.xlsx` 或 `.pptx` |
| `open <文件>` | 打开文件并保持处理会话 | 仅用于大量连续操作；必须与 `close` 配对 |
| `view <文件> <模式>` | 查看大纲、文本、问题或渲染结果 | 快速了解文件和执行视觉检查 |
| `get <文件> <路径>` | 读取指定元素及其子节点 | 检查单个段落、单元格、形状或结构 |
| `query <文件> <选择器>` | 按条件查找多个元素 | 查找标题、公式、空段落或缺少替代文本的图片 |
| `set <文件> <路径>` | 修改现有元素的内容或属性 | 修改文本、格式、公式和位置 |
| `add <文件> <父路径>` | 在父节点下添加元素 | 添加段落、工作表、幻灯片、形状或图表 |
| `remove <文件> <路径>` | 删除指定元素 | 删除多余内容或结构 |
| `move <文件> <路径>` | 调整元素顺序 | 移动段落、幻灯片或其他节点 |
| `batch <文件>` | 默认以原子方式执行多项操作 | 批量写入单元格或重复结构 |
| `import <文件> <工作表> <数据源>` | 将 CSV 或 TSV 导入 XLSX | 导入区域化表格数据、设置表头和起始单元格 |
| `save <文件>` | 写入磁盘并保留处理会话 | 仅生成中途快照，不能结束任务 |
| `close <文件>` | 写入磁盘并结束处理会话 | 每个创建、修改或打开的文件都必须执行 |
| `validate <文件>` | 检查 OpenXML 包和结构 | 发现格式包、关系和 XML 错误；不能代替视觉检查 |

`view` 常用模式包括：`outline` 查看结构，`text` 提取文本，`issues` 检查内容和格式问题，`html` 查看静态渲染，`screenshot` 生成视觉检查图片。`get --depth N` 用于展开 N 层子节点；需要机器读取结果时添加 `--json`。

## 操作规则

- 路径索引从 1 开始；`--index` 通常从 0 开始，Excel 行列插入例外。
- 始终引用 shell 中带 `[]` 的路径。
- 使用不带 `#` 的十六进制颜色值。
- 元素属性统一使用重复的 `--prop key=value`；不要使用 `--props`。
- 分别处理 shell、JSON 和 OfficeCLI 三层中的 `$`、反斜杠和换行转义。
- 使用 `data=` 填充 Word 或 PowerPoint 表格时遵循 CSV 引号规则：含逗号、行分隔符或双引号的单元格用双引号包裹，内部双引号写成两个双引号。
- 不能编造用户未提供的数据、来源、业务结果或文档事实。

## 交付检查

对每个文件依次执行 `close`、`validate`、`view issues`、结构检查和视觉检查。检查后若再次修改，必须重新执行这一完整顺序。版式重要时使用 `view screenshot --grid auto` 或逐页截图。无法渲染时明确说明“未完成视觉验证”，不得把结构验证等同于视觉验证；`close` 未成功时不得交付。
