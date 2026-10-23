# PowerPoint 演示文稿处理

处理 `.pptx`、`.pptm` 文件时遵循本文件；`.pptm` 同时遵循主技能中的宏启用文件规则。

## 规划

1. 明确受众、演讲时长、语气、品牌规范和目标查看器。
2. 创建形状前写出完整且有逻辑的幻灯片标题序列。
3. 每页只表达一个核心观点，使内容可在三秒内理解。
4. 选定一种视觉母题、一套配色和最多两种主要字体。

## 视觉规则

- 新建自由形状时使用 16:9 画布和 12 栏网格，并设置明确的 x/y/width/height。模板占位符可以继承版式或母版位置，不要为了满足显式坐标规则而破坏继承关系。
- 标题通常不小于 36pt，正文不小于 18pt，标题至少为正文的两倍。
- 深色背景使用高亮正文，投影环境不能依赖中灰色小字。
- 避免项目符号墙、装饰性数据、无意义图标和过密卡片。
- 使图表匹配数据关系、标签可读，关键业务趋势的坐标轴不得制造误导。
- 使用真正的 connector 绘制流程连接线，并明确箭头端点。
- 为内容页添加演讲者备注，为图片设置替代文本。
- 仅用动画辅助理解，确保静态画面也完整可读。

## 构建示例

目标：创建一张使用空白版式和自定义背景的幻灯片，并在明确坐标处添加标题形状。

```bash
"$OCLI" create "$FILE"
"$OCLI" add "$FILE" / --type slide --prop layout=blank --prop background=F7F4EF
"$OCLI" add "$FILE" '/slide[1]' --type shape --prop text='标题' --prop x=2cm --prop y=1.5cm --prop width=29cm --prop height=3cm --prop size=40pt
"$OCLI" get "$FILE" '/slide[1]' --depth 1
"$OCLI" close "$FILE"
```

预期结果：演示文稿包含一张 16:9 幻灯片，背景色为 `F7F4EF`，标题位于指定位置并使用 40pt 字号。

按照“背景 → 标题 → 主体 → 图表或图片 → 连接线 → 备注”的顺序逐页构建。每次结构操作后立即使用 `get` 检查 ID 和层级，批量操作前先验证一页。

## 版式与占位符

- `layout` 只指定版式关系，不会自动把版式中的所有占位符实例化到幻灯片。
- 创建幻灯片时可使用 `title=` 和 `text=` 生成标题、正文占位符；需要其他角色时添加 `placeholder` 并指定 `phType`。
- 编辑模板中的占位符时优先使用版式绑定 ID：`placeholder[@idx=K]`。存在相同 ID 的不同类型时使用 `placeholder[@type=body][@idx=K]` 等组合条件缩小范围。
- 修改前使用 `get '/slide[N]' --depth 1` 读取实际 `phType`、`phIndex`、名称和路径；不根据占位符在页面上的顺序猜测索引。
- 占位符没有直接的 `x`、`y`、`width`、`height` 时，读取 `effective.x`、`effective.y`、`effective.width`、`effective.height` 及对应的 `.src`，确认实际尺寸和来源。只有确实需要覆盖模板布局时才写入直接坐标。
- `get` 返回 `hidden=true` 的形状不会出现在预览中；不要把隐藏元素误判为渲染丢失，也不要依赖隐藏内容传递必须可见的信息。

以下示例用于创建带标题和正文占位符的幻灯片，并演示如何稳定修改模板中的正文占位符：

```bash
"$OCLI" create "$FILE"
"$OCLI" add "$FILE" / --type slide --prop layout='Title and Content' \
  --prop title='季度报告' --prop text='核心摘要'
"$OCLI" get "$FILE" '/slide[1]' --depth 1
"$OCLI" set "$FILE" '/slide[1]/placeholder[@type=body][@idx=1]' \
  --prop text='更新后的核心摘要'
"$OCLI" close "$FILE"
```

预期结果：幻灯片包含与版式关联的标题和正文占位符，正文通过绑定 ID 修改，不受其他形状插入或删除影响。若模板使用不同的绑定 ID，以 `get` 返回值为准。

## 转场、动画与 Morph

- 转场属于幻灯片，使用 `transition=`；可在同一值中指定方向、速度或毫秒时长，例如 `push-right-1500`、`fade-slow`。
- 动画属于形状或图表，明确设置 `effect`、`class`、`trigger` 和 `duration`。只在有助于理解时添加，静态页面也必须完整可读。
- 内容页使用 `notes` 保存演讲者备注；不要把讲稿堆在幻灯片正文中。
- 只有用户要求跨页连续运动时才使用 Morph。相邻页面参与配对的形状必须使用完全相同、以 `!!` 开头的名称，并产生明显的位置、尺寸或旋转变化。
- shell 中始终单引号引用包含 `!!` 的名称。完成后查询两页的同名形状，确认名称一致且各页仅有一个目标形状。
- Morph 必须在 PowerPoint、WPS 或 Keynote 中验证；LibreOffice及部分网页查看器可能只显示静态页面或淡入淡出。

以下示例假定文件中已经存在两张幻灯片，并且动画示例引用的第一页第一个形状已经存在。实际执行前先用 `get` 确认幻灯片数量和形状路径。

```bash
"$OCLI" set "$FILE" '/slide[2]' --prop transition=push-right-1500
"$OCLI" add "$FILE" '/slide[1]/shape[1]' --type animation \
  --prop effect=fade --prop class=entrance --prop trigger=onClick --prop duration=500
"$OCLI" add "$FILE" '/slide[1]' --type notes --prop text='说明本页重点。'

# Morph：相邻页面分别创建同名形状，并改变位置或尺寸
"$OCLI" add "$FILE" '/slide[1]' --type shape \
  --prop 'name=!!hero' --prop text='主题' --prop x=2cm --prop y=5cm --prop width=10cm --prop height=3cm
"$OCLI" set "$FILE" '/slide[2]' --prop transition=morph
"$OCLI" add "$FILE" '/slide[2]' --type shape \
  --prop 'name=!!hero' --prop text='主题' --prop x=20cm --prop y=2cm --prop width=12cm --prop height=4cm
"$OCLI" close "$FILE"
```

现有元素路径不确定时先使用 `get '/slide[N]' --depth 1`；普通形状使用返回的 `@id` 或 `@name`，占位符使用 `@idx` 与类型组合条件建立稳定路径。

## 格式检查

```bash
"$OCLI" close "$FILE"
"$OCLI" validate "$FILE"
"$OCLI" view "$FILE" issues  # 检查溢出、越界、遮挡、低对比度、图片拉伸和损坏关系
"$OCLI" view "$FILE" outline  # 检查幻灯片顺序、标题和元素数量
"$OCLI" view "$FILE" screenshot --grid auto -o slides-grid.png  # 总览整套幻灯片的布局节奏
"$OCLI" view "$FILE" screenshot --page 1 -o slide-1.png  # 渲染第一页进行视觉检查
```

`view issues` 返回任何问题时先按其路径修复并重复检查；可用 `--type low_contrast`、`--type picture_aspect_distorted` 或 `--type broken_part_ref` 定位特定问题。逐页检查重叠、截断、越界、低对比度、对齐、层级、空白节奏和图表可读性；使用 `query animation`、`query notes`、`query slide` 和同名形状查询检查效果结构。修复后重新关闭并执行完整检查。动画和 Morph 需要在目标演示软件中额外验证。
