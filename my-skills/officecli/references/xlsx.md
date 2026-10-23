# Excel 文件处理

处理 `.xlsx`、`.xlsm` 文件或导入 CSV、TSV 数据时遵循本文件；`.xlsm` 同时遵循主技能中的宏启用文件规则。

## 数据与公式规则

- 能由其他单元格计算的值使用公式，不硬编码计算结果。
- 清晰区分输入、计算和输出区域；专业模型使用蓝色输入、黑色公式、绿色跨表引用、黄色假设底色。
- 将年份按文本处理，为百分比、货币、负数和零设置一致的数字格式。
- 导入大数据时先分块写入纯值，再添加公式、格式、表格和图表。
- 使用完整的单元格属性名，例如 `font.color`、`font.size` 和 `fill`。
- 跨表公式优先通过 JSON batch 或安全引用写入，防止 `!` 被 shell 破坏。
- 批量公式和数据写入默认使用原子 `batch`；任一项失败时修正命令后整批重试，不把已回滚的成功项当作已写入。
- 同时检查公式字符串、`evaluated`、`cachedValue` 和 `computedValue`，不能只检查显示结果。
- 修改上游值后重新读取所有关键下游公式；需要可靠缓存时，重新提交原公式，再核对 `cachedValue` 与 `computedValue`。
- 两个值同时存在但不一致时按缓存过期处理；`evaluated=false`、空结果或错误值不得交付。
- 只有业务确实需要循环引用时才启用 `calc.iterate`，并验证迭代次数、收敛阈值和最终结果。
- 不把“目标查看器会自动重算”或 `fullCalcOnLoad` 标记当作当前结果正确的证明。

跨工作表公式使用不展开 shell 内容的 JSON：

```bash
"$OCLI" batch "$FILE" --json <<'JSON'
[
  {"command":"set","path":"/Summary/B2","props":{"formula":"'Data Sheet'!B2"}}
]
JSON
"$OCLI" close "$FILE"
```

## CSV 与 TSV 导入

使用 `import` 将 UTF-8 CSV 或 TSV 写入已有 `.xlsx` 工作表。`import` 的目标仅支持 `.xlsx`；处理 `.xlsm` 时使用 `set` 或 `batch` 原位写入，不能通过导入命令重建文件。数据源必须是文本文件或标准输入，不能把 `.xlsx`、`.xls` 等工作簿作为 CSV 导入源。

```bash
"$OCLI" create "$FILE"
"$OCLI" import "$FILE" /Sheet1 "$SOURCE" \
  --delimiter ';' --decimal ',' --header --start-cell A1 --json
"$OCLI" get "$FILE" '/Sheet1/A1:B3' --depth 1 --json
"$OCLI" close "$FILE"
```

该示例导入以分号分列、逗号作为小数点的数据，并将首行设为表头。普通逗号 CSV 或制表符 TSV 应省略不需要的选项，让程序按文件内容、`--format` 或扩展名判断。

- 分隔符选择顺序为：显式 `--delimiter`、文件首行 `sep=`、`--format`、文件扩展名、默认逗号。`sep=` 声明不会作为数据导入。
- 小数逗号必须显式设置 `--decimal ','`，并使用不同的字段分隔符；不猜测含义不明确的 `1,234`。
- 保留标准 CSV 引号语义，允许带分隔符、换行或双引号的字段。导入后检查报告的行列数、表头、首尾数据和数字/日期类型；意外只得到一列时按警告修正分隔符后重新导入。
- 起始位置可以用工作表路径配合 `--start-cell B2` 指定，也可以直接使用 `/Sheet1/B2` 作为目标路径。使用一种方式明确位置，避免同时给出相互冲突的目标。

## 构建示例

目标：创建一个数据工作表，写入文本、数值和求和公式，形成最小的公式驱动工作簿。

```bash
"$OCLI" create "$FILE"
"$OCLI" add "$FILE" / --type sheet --prop name=Data
"$OCLI" set "$FILE" /Data/A1 --prop value='名称' --prop bold=true
"$OCLI" set "$FILE" /Data/B2 --prop value=100 --prop type=number
"$OCLI" set "$FILE" /Data/B10 --prop formula='SUM(B2:B9)'
"$OCLI" close "$FILE"
```

预期结果：工作簿包含 `Data` 工作表、加粗表头、数值单元格以及计算 `B2:B9` 合计的公式单元格。

使图表引用单元格范围，不复制静态数据。突出表头、冻结窗格并设置合理列宽；避免合并单元格破坏筛选。仅用条件格式表达业务含义。

## 行列与东亚文本

- 插入行或列并需要继承邻近区域的字体、填充和公式时，使用 `"$OCLI" add "$FILE" /Data --from '/Data/row[5]' --index 5` 或对应列路径；公式引用会随插入调整。只需要空白结构时使用 `add --type row` 或 `add --type col`。
- 需要在 Excel 中保存日文假名或其他 CJK 拼音指南时，对文本单元格设置 `phonetic=`，然后使用 `get --json` 确认读取结果。它是单元格共享字符串的注音属性，不应写入另一个可见单元格代替。
- 需要控制工作表显示或打印时，使用工作表属性 `gridlines`、`headings`、`visibility`、`orientation`、`paperSize`、`fitToPage`、`printArea` 和 `margin.*`；设置前运行 `"$OCLI" help xlsx set sheet`，并确保工作簿至少保留一个可见工作表。

## 单元格内图片

需要图片随单元格排序、筛选和移动，或需要图片本身成为单元格值时，使用单元格的 `image=` 属性。它不同于浮在工作表上的 `picture`：公式会把单元格内图片视为非空值。

```bash
"$OCLI" set "$FILE" /Data/C2 --prop "image=$IMAGE" --prop alt='产品照片'
"$OCLI" get "$FILE" /Data/C2 --json
"$OCLI" close "$FILE"
```

预期结果：`C2` 成为图片单元格，读取结果包含 `type=Image`、替代文本、媒体类型和文件大小。仅使用栅格图片，不使用 SVG；删除时设置 `image=none`。在支持单元格内图片的 Excel 中验证显示效果，目标查看器不支持该结构时可能显示 `#VALUE!`，此时根据交付环境改用浮动 `picture`。

## 图表与扫描辅助

- 图表必须引用工作表范围，具有描述性标题和明确的系列名称；不得保留 `Series1` 等默认名称。
- 图表类型必须匹配数据关系，坐标轴范围不得制造误导。源数据为空、隐藏或包含错误值时先修复数据。
- 只有数据存在顺序关系时才使用 Sparkline；横截面分类数据不使用趋势迷你图。
- 条件格式只表达业务含义，保持正向、负向和中性色一致，不用颜色替代真实数值。

```bash
"$OCLI" add "$FILE" /Data --type chart \
  --prop chartType=line --prop title='收入趋势' --prop series1.name='收入' \
  --prop series1.values='Data!B2:B13' --prop series1.categories='Data!A2:A13'
"$OCLI" add "$FILE" /Data --type sparkline \
  --prop type=line --prop dataRange=B2:B13 --prop location=C2
"$OCLI" add "$FILE" /Data --type conditionalformatting \
  --prop type=colorScale --prop ref=B2:B13 \
  --prop minColor=FFCDD2 --prop midColor=FFFFFF --prop maxColor=C8E6C9
"$OCLI" close "$FILE"
```

示例中的工作表和范围必须替换为实际路径。创建后使用 `get` 或 `query` 检查图表系列数、标题、范围、Sparkline 位置和条件格式范围。

## 格式检查

```bash
"$OCLI" close "$FILE"
"$OCLI" validate "$FILE"
"$OCLI" view "$FILE" issues
"$OCLI" view "$FILE" issues --type numeric_overflow
"$OCLI" view "$FILE" issues --type general_precision_loss
"$OCLI" view "$FILE" html  # 检查工作表渲染、列宽和图表布局
"$OCLI" view "$FILE" screenshot --grid auto -o /tmp/xlsx-contact-sheet.png  # 生成工作表缩略图
```

检查 `#REF!`、`#DIV/0!`、`#VALUE!`、`#NAME?` 等错误，核对公式缓存、年份格式、列宽、截断、图表标题和图例、打印范围、隐藏辅助表、活动工作表及跨表引用。`numeric_overflow` 表示明确数字格式的显示宽度不足，按建议加宽列、缩短格式或谨慎使用 `shrinkToFit`；`general_precision_loss` 表示 General 格式会以科学计数法显示较长数字，设置明确的 `numberformat`，加宽列不能解决。随机抽查至少 2–3 个关键公式，使用 `get --json` 核对公式、依赖来源、`evaluated`、`cachedValue` 和 `computedValue`。
