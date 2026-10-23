# Word 表单与字段

创建、读取或修改可填写 Word 表单、内容控件、邮件合并字段或文档保护时遵循本文件，并同时遵循 `docx.md`。

## 选择元素

| 需求 | 元素 |
|---|---|
| 文本、富文本、下拉框、组合框、日期、图片、分组或复选框 | `sdt` 内容控件 |
| 兼容旧版 Word 的文本、下拉框或复选框 | `formfield` |
| 邮件合并、条件、引用、页码或自动编号 | `field` |
| 限制非字段内容编辑 | 文档 `protection=forms` |

优先使用 `sdt`。只有兼容既有模板或旧版 Word 时才使用 `formfield`。字段属性不确定时分别运行 `help docx sdt`、`help docx formfield` 和 `help docx field`。

## 构建规则

- 每个 SDT 都设置唯一、稳定的 `tag` 和清晰的 `alias`；不要依赖位置索引作为业务标识。
- 使用真实控件，不用下划线、空格或方框字符模拟输入区域。
- 下拉框和组合框使用 `items=`；显示文本与存储值不同时使用 `显示文本|存储值`。
- SDT 复选框使用 `type=checkbox` 和 `checked=`；仅在兼容旧模板时使用 FormField 复选框。
- `MERGEFIELD` 属于 `field`，不会出现在 `view forms` 的控件清单中，必须另行 `query field`。
- 先完成静态文本、样式和布局，再设置 `protection=forms`。保护后不要无意绕过限制修改正文。
- 修改既有表单前先读取控件、字段和保护状态，保留未知控件及数据绑定结构。
- 表单属性被拒绝可能仍留下缺少该属性的控件。使用 `--json` 检查 `success` 和 `warnings`；出现 `WARNING: UNSUPPORTED` 或退出码 `2` 时，不得继续在该结果上构建。

## 最小示例

```bash
"$OCLI" create "$FILE"
"$OCLI" add "$FILE" /body --type paragraph --prop text='姓名：'
"$OCLI" add "$FILE" '/body/p[1]' --type sdt \
  --prop type=text --prop alias='姓名' --prop tag=name --prop text='请输入姓名'
"$OCLI" add "$FILE" /body --type sdt \
  --prop type=dropdown --prop alias='状态' --prop tag=status \
  --prop 'items=草稿|DRAFT,正式|FINAL'
"$OCLI" add "$FILE" /body --type sdt \
  --prop type=checkbox --prop alias='同意条款' --prop tag=agree --prop checked=false
"$OCLI" add "$FILE" /body --type field \
  --prop fieldType=mergefield --prop name=CustomerName
"$OCLI" set "$FILE" / --prop protection=forms
"$OCLI" close "$FILE"
```

将示例中的段落路径替换为实际稳定路径。设置保护后必须确认命令成功；不把控件存在等同于保护已经启用。

## 更新已有控件

先用 `query sdt --json` 获取 `@sdtId` 稳定路径，再修改控件。`items=` 会替换完整选项列表；更新列表后同时设置有效的当前值，避免保留已不存在的选项值。

```bash
"$OCLI" query "$FILE" sdt --json
"$OCLI" set "$FILE" '/body/sdt[@sdtId=1]' --prop checked=true
"$OCLI" set "$FILE" '/body/sdt[@sdtId=2]' \
  --prop 'items=待审|REVIEW,发布|PUBLISHED' \
  --prop dropDown.lastValue=PUBLISHED --prop text='发布'
"$OCLI" set "$FILE" '/body/sdt[@sdtId=3]' \
  --prop format=yyyy-MM-dd \
  --prop date.fullDate=2026-01-01T00:00:00Z --prop text='2026-01-01'
"$OCLI" close "$FILE"
```

根据控件类型选择属性：复选框使用 `checked`，下拉框使用 `dropDown.lastValue`，组合框使用 `comboBox.lastValue`，日期控件使用 `format`、`date.fullDate`、`date.calendar` 和 `date.lid`。当前值属性与可见内容相互独立，选择值或日期变化时同时设置匹配的 `text`。需要切换占位显示或引用占位构件时使用 `placeholder`、`placeholderText`。修改后重新查询，确认控件类型、选项、当前值和显示文本一致。

## 修改与锁定

- `lock=unlocked`：内容和控件均可编辑。
- `lock=sdtLocked`：内容可编辑，控件不可删除。
- `lock=contentLocked`：内容只读，控件可删除。
- `lock=sdtContentLocked`：内容和控件均锁定。

`protection=forms` 控制文档级编辑范围，`lock` 控制单个 SDT。两者作用不同，按用户的填写流程组合使用。

## 检查

```bash
"$OCLI" close "$FILE"
"$OCLI" validate "$FILE"
"$OCLI" view "$FILE" forms
"$OCLI" query "$FILE" sdt --json
"$OCLI" query "$FILE" formfield --json
"$OCLI" query "$FILE" field --json
"$OCLI" get "$FILE" / --depth 1
```

`view forms` 用于快速汇总表单控件，`query sdt --json` 用于核对完整属性。确认两者报告的控件类型和复选状态一致；确认所有 SDT 均有 `alias`、`tag` 和正确类型；下拉选项、当前值、日期格式、复选状态和锁定状态符合预期；MERGEFIELD 名称正确；根节点显示所需的 `protection` 和 `protectionEnforced=true`。最后在目标 Word 查看器中实际填写、切换选项和检查只读区域。
