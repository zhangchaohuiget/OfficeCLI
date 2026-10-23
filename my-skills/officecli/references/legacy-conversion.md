# 旧版 Office 文件转换

OfficeCLI 不直接处理 `.doc`、`.xls`、`.ppt`。先转换格式，再把转换结果交给 OfficeCLI：

`.docm`、`.xlsm`、`.pptm` 不属于本转换流程，按主技能中的宏启用文件规则直接读取或原位编辑。

| 输入 | 转换结果 | 后续规则 |
|---|---|---|
| `.doc` | `.docx` | `docx.md` |
| `.xls` | `.xlsx` | `xlsx.md` |
| `.ppt` | `.pptx` | `pptx.md` |

## 目录

- [通用规则](#通用规则)
- [工具顺序](#工具顺序)
- [Windows Microsoft Office](#windows-microsoft-office)
- [macOS textutil](#macos-textutil)
- [LibreOffice](#libreoffice)
- [转换后检查](#转换后检查)

## 通用规则

- 保留源文件，不得覆盖、移动或删除。
- 为每次转换选择不存在的输出路径；不同工具写入不同目录，不覆盖已有结果。
- 以只读方式打开源文件，禁用宏、事件、外部链接更新和交互提示。
- 密码保护文件无法安全打开时停止，请求用户提供已解密副本，不尝试绕过保护。
- `.doc`、`.xls`、`.ppt` 可能含宏；转换为 `.docx`、`.xlsx`、`.pptx` 会移除宏。任务依赖宏、ActiveX 或嵌入代码时停止，不得通过本流程转换。
- 确认输出存在且非空后，执行结构、内容和视觉检查。只有检查通过时才视为转换成功。
- 当前工具失败、校验失败或内容不完整时，不编辑、不交付该输出，继续尝试下一种工具。所有工具均失败时请求用户提供对应的 `.docx`、`.xlsx` 或 `.pptx`。

## 工具顺序

| 平台 | `.doc` | `.xls` | `.ppt` |
|---|---|---|---|
| Windows | Microsoft Word → LibreOffice | Microsoft Excel → LibreOffice | Microsoft PowerPoint → LibreOffice |
| macOS | `textutil` → LibreOffice | LibreOffice | LibreOffice |

使用系统中已经存在的工具。OfficeCLI 当前不支持的其他平台停止处理并报告。

## Windows Microsoft Office

在 PowerShell 中按扩展名调用对应的 Office 应用。格式代码 `16`、`51`、`24` 分别表示 DOCX、XLSX、PPTX：

```powershell
$InputFile = (Resolve-Path "C:\绝对路径\输入文件").Path
$OutputFile = [IO.Path]::GetFullPath("C:\绝对路径\输出文件")
if (Test-Path $OutputFile) { throw "输出文件已存在：$OutputFile" }

$ext = [IO.Path]::GetExtension($InputFile).ToLowerInvariant()
$app = $null
$item = $null
try {
    switch ($ext) {
        ".doc" {
            $app = New-Object -ComObject Word.Application
            $app.Visible = $false
            $app.DisplayAlerts = 0
            $app.AutomationSecurity = 3
            $item = $app.Documents.Open($InputFile, $false, $true)
            $item.SaveAs2($OutputFile, 16)
        }
        ".xls" {
            $app = New-Object -ComObject Excel.Application
            $app.Visible = $false
            $app.DisplayAlerts = $false
            $app.AutomationSecurity = 3
            $app.EnableEvents = $false
            $app.AskToUpdateLinks = $false
            $item = $app.Workbooks.Open($InputFile, 0, $true)
            $item.SaveAs($OutputFile, 51)
        }
        ".ppt" {
            $app = New-Object -ComObject PowerPoint.Application
            $app.DisplayAlerts = 1
            $app.AutomationSecurity = 3
            $item = $app.Presentations.Open($InputFile, $true, $false, $false)
            $item.SaveAs($OutputFile, 24)
        }
        default { throw "不支持的输入格式：$ext" }
    }
} finally {
    if ($item) {
        if ($ext -eq ".doc") { $item.Close(0) }
        elseif ($ext -eq ".xls") { $item.Close($false) }
        else { $item.Close() }
    }
    if ($app) { $app.Quit() }
}
if (-not (Test-Path $OutputFile) -or (Get-Item $OutputFile).Length -eq 0) {
    throw "格式转换失败：未生成有效输出"
}
```

Microsoft Office 不存在、自动化失败或输出未通过检查时，改用 LibreOffice。

## macOS textutil

`textutil` 仅用于 `.doc → .docx`：

```bash
INPUT='/绝对路径/输入.doc'
OUTPUT='/绝对路径/输出.docx'
test ! -e "$OUTPUT"
/usr/bin/textutil -convert docx -output "$OUTPUT" "$INPUT"
test -s "$OUTPUT"
```

`textutil` 可能生成 OfficeCLI 严格校验不接受的 `w:sz-cs` 元素。出现该错误或其他结构错误时，不猜测性修复 XML，改用 LibreOffice。

## LibreOffice

仅使用系统中可执行的 `soffice`。macOS 同时检查 `/Applications/LibreOffice.app/Contents/MacOS/soffice`；Windows 同时检查 LibreOffice 安装目录。每次转换使用独立的临时用户配置目录，避免复用进程、配置锁、可信位置和交互状态。根据输入格式设置 `TARGET`：

```bash
INPUT='/绝对路径/输入文件'
OUTDIR='/绝对路径/独立输出目录'
PROFILE_DIR='/绝对路径/不含空格的独立临时配置目录'
TARGET='docx'  # .doc 使用 docx；.xls 使用 xlsx；.ppt 使用 pptx
mkdir -p "$OUTDIR" "$PROFILE_DIR"
"$SOFFICE" "-env:UserInstallation=file://$PROFILE_DIR" \
  --headless --nologo --nodefault --norestore \
  --convert-to "$TARGET" --outdir "$OUTDIR" "$INPUT"
```

Windows PowerShell 使用同样的参数执行：

```powershell
$ProfileUri = "file:///C:/绝对路径/独立临时配置目录"
& $Soffice "-env:UserInstallation=$ProfileUri" `
  --headless --nologo --nodefault --norestore `
  --convert-to $Target --outdir $OutDir $InputFile
```

确认实际输出路径及扩展名正确且文件非空，不依赖进程退出码判断成功。

### DOCX 字体表兼容性

LibreOffice 可能在 `/word/fontTable.xml` 的 `w:charset` 中写入 OfficeCLI 严格校验不接受的 `w:characterSet`。仅当所有校验错误均指向该部件、该元素和该属性，并且 `"$OCLI" raw "$FILE" /fontTable` 显示目标节点只含 `w:val` 与 `w:characterSet` 时，才清理这个属性。

`raw-set remove` 不能直接删除属性。按实际出现的每个 `w:val` 分别替换元素并原样保留 `w:val`。例如：

```bash
"$OCLI" raw-set "$FILE" /fontTable \
  --xpath '/w:fonts/w:font/w:charset[@w:val="00"]' \
  --action replace --xml '<w:charset w:val="00"/>'
"$OCLI" raw-set "$FILE" /fontTable \
  --xpath '/w:fonts/w:font/w:charset[@w:val="01"]' \
  --action replace --xml '<w:charset w:val="01"/>'
"$OCLI" close "$FILE"
```

示例值不是固定值。若节点含其他属性、存在其他错误或无法确认修复影响，放弃该输出。修复后重新执行完整检查。

不要为 XLSX 或 PPTX 校验错误套用 DOCX 字体表修复。复杂图表、数据透视表、OLE 对象、媒体、动画、修订、域和分页可能在转换中发生变化；发现错误时优先更换转换工具或请求用户提供对应的 OpenXML 文件。

## 转换后检查

```bash
"$OCLI" close "$FILE"
"$OCLI" validate "$FILE"
"$OCLI" view "$FILE" text
"$OCLI" view "$FILE" outline
```

随后读取目标格式参考文件并执行其完整交付检查：

- DOCX：核对段落、表格、图片、页数、修订、域和分页。
- XLSX：核对工作表名称、使用区域、公式、合并单元格、数字与日期显示格式、空白格式化单元格、图表、数据透视表和计算结果。
- PPTX：核对幻灯片数量、文字、图片、形状、图表、媒体、动画和版式。

完成逐页或逐工作表视觉检查后再继续编辑，并向用户说明交付文件是转换后的格式。
