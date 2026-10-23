# Word 文档处理

处理 `.docx`、`.docm` 文件时遵循本文件；`.docm` 同时遵循主技能中的宏启用文件规则。

## 开始前

1. 确认文档类型、受众、语言、纸张、页数和交付格式。
2. 编辑现有文件前运行 `view outline`、`view text` 和 `get / --depth 2`。
3. 属性不确定时运行 `"$OCLI" help docx <元素>`。
4. 先定义标题、Heading 1/2/3、正文和题注等样式，再写内容。

## 排版规则

- 建立“标题 → 一级标题 → 二级标题 → 正文”的清晰层级，不使用加粗和字号模拟标题。
- 统一页边距、字体、段距、行距、颜色和表格风格。
- 使用真实 `PAGE` 字段生成页码，使用真实 TOC 字段生成目录，不手写页码。
- 存在三个以上标题时检查是否需要目录；目录来源必须是 Heading 样式或带 `outlineLvl` 的样式。
- 将模板占位符替换为实际内容或真实 MERGEFIELD，不交付 `{{name}}` 等内部标记。
- 为图片添加替代文本，避免表格跨页断裂和标题单独悬挂。
- 在段落级 `text` 中使用 `\n` 创建新段落，使用 `\v` 创建同一段落内的软换行，使用 `\t` 创建制表符。各新段落继承源段落样式。
- `run` 不能跨段落；对 `run` 设置的 `\n` 会降级为软换行。需要创建真实段落时操作段落或正文容器。表格行的 `c1`…`cN` 快捷属性也是例外，其中 `\n` 表示单元格内软换行。
- 命令行 `--prop text=` 会解释两个字符组成的 `\n`、`\v`、`\t`。Batch JSON 不解释字符串中的 `\\n` 或 `\\v`：使用 JSON `"\n"` 传入真实换行，使用 `"\u000b"` 传入软换行。
- 分别检查 shell、JSON 和 OfficeCLI 的转义层，使用 `view text` 确认正文中没有意外残留的反斜杠转义文本。

## 构建示例

目标：创建包含文档标题、一级标题和正文的最小 Word 文档，并使用语义化样式表达层级。

```bash
"$OCLI" create "$FILE"
"$OCLI" add "$FILE" /styles --type style \
  --prop id=Title --prop name='标题' --prop type=paragraph \
  --prop basedOn=Normal --prop size=26pt --prop bold=true
"$OCLI" add "$FILE" /styles --type style \
  --prop id=Heading1 --prop name='标题 1' --prop type=paragraph \
  --prop basedOn=Normal --prop outlineLvl=0 --prop size=18pt --prop bold=true
"$OCLI" add "$FILE" /body --type paragraph --prop text='标题' --prop style=Title
"$OCLI" add "$FILE" /body --type paragraph --prop text='第一章' --prop style=Heading1
"$OCLI" add "$FILE" /body --type paragraph --prop text='正文内容' --prop style=Normal
"$OCLI" close "$FILE"
```

预期结果：文档包含一个 `Title` 段落、一个 `Heading1` 段落和一个 `Normal` 正文段落，文件已保存且没有活动处理会话。

优先使用语义化元素处理段落、表格、图片、页眉页脚、脚注、批注、书签、编号和字段；只有高层接口无法表达时才使用 `raw-set`。

## Markdown 快速构建

当内容已经是 Markdown，或需要一次生成多级标题、段落、列表、代码块、引用和简单表格时，使用 `markdown` 将其展开为可继续编辑的 Word 原生元素。多行内容优先通过 `src` 读取本地 `.md` 文件，避免 shell 改变换行和反斜杠。导入前使用 `get /styles --depth 1` 检查 Markdown 会用到的 `Heading1` 至 `Heading6`；缺少时先按 `help docx style` 创建对应样式，不忽略样式缺失警告。

```bash
"$OCLI" create "$FILE"
"$OCLI" add "$FILE" /body --type markdown --prop "src=$SOURCE_MD"
"$OCLI" get "$FILE" /body --depth 2
"$OCLI" close "$FILE"
"$OCLI" validate "$FILE"
```

预期结果：Markdown 的标题、列表、表格等块结构被转换为普通 Word 元素，后续通过返回路径或 `@paraId` 修改，不存在可再次寻址的 `markdown` 节点。

内联输入必须包含真实换行；在 bash 或 zsh 中可使用 `--prop markdown=$'# 标题\n\n- 第一项\n- 第二项'`，其他 shell 优先使用 `src`。链接和图片只保留可见文本，不会创建超链接关系或嵌入图片；表格单元格中的行内格式不会保留，嵌套引用会展平。需要这些结构时，在展开后使用 `hyperlink`、`picture` 或相应原生元素补充。

## 高级结构

- 公式：使用 `equation`，明确选择 `inline` 或 `display`，不要把 LaTeX 源码当正文写入。
- 脚注与尾注：将 `footnote` 或 `endnote` 添加到目标段落，由 OfficeCLI 创建引用和对应部件。
- 书签与交叉引用：用 `bookmark` 建立稳定目标，再用 `fieldType=ref` 或 `pageref` 引用。
- 自动编号：用 `fieldType=seq` 和唯一序列名生成图、表等编号，不硬编码序号。
- 字段缓存：全部 `SEQ` 字段完成后执行 `recalcFields=seq`；PAGE、PAGEREF、TOC 等依赖分页的字段设置 `updateFields=true`，并在目标查看器中更新和检查。
- 结构操作会改变位置索引。继续编辑时使用命令返回的 `@paraId`、书签名或其他稳定路径，不猜测后续 `p[N]`。
- 涉及可填写字段、内容控件、邮件合并或保护时，同时读取 `docx-forms.md`。

```bash
"$OCLI" add "$FILE" /body --type equation --prop mode=display --prop formula='E = mc^2'
"$OCLI" add "$FILE" '/body/p[1]' --type footnote --prop text='脚注内容'
"$OCLI" add "$FILE" '/body/p[1]' --type bookmark --prop name=section1
"$OCLI" add "$FILE" '/body/p[2]' --type field --prop fieldType=seq --prop id=Figure
"$OCLI" add "$FILE" '/body/p[3]' --type field --prop fieldType=pageref --prop name=section1 --prop hyperlink=true
"$OCLI" set "$FILE" / --prop recalcFields=seq --prop updateFields=true
"$OCLI" close "$FILE"
```

示例中的段落路径必须按实际结构替换。使用 `query equation`、`query footnote`、`query endnote`、`query bookmark` 和 `query field` 核对元素数量、目标及 `evaluated` 状态。

## 格式检查

```bash
"$OCLI" close "$FILE"
"$OCLI" validate "$FILE"
"$OCLI" view "$FILE" outline  # 检查标题层级、章节和文档结构
"$OCLI" view "$FILE" text     # 检查正文、占位符和转义字面量
"$OCLI" view "$FILE" html -o /tmp/docx-preview.html  # 检查静态版式和批注标记
"$OCLI" view "$FILE" screenshot --grid auto -o /tmp/docx-contact-sheet.png  # 生成全页缩略图
```

确认不存在占位符、空标题、孤行、异常分页、截断和低对比度；确认字段结构存在且目录来源有效。涉及批注时同时使用 `query comment --json` 核对作者和正文，并在 HTML 预览中检查高亮与标记。缩略图仅用于定位，细节使用逐页全尺寸图像复核。
