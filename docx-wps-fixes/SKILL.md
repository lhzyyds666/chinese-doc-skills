---
name: docx-wps-fixes
description: Use when Codex reads, writes, edits, converts, or extracts `.doc`/`.docx`/Word/WPS documents, especially Chinese thesis/report files, WPS-authored DOCX, OOXML/XML edits, Chinese font/字号 rules, OMML formulas, formula numbering with tab stops, m:oMathPara/m:eqArr issues, English Keywords formatting, Word COM validation, python-docx/docx-js/pandoc/LibreOffice workflows. MANDATORY companion to official docx/documents skill; load in parallel at first sign of Word work.
---

# DOCX 中文文档规则 & WPS 兼容

创建或编辑中文 `.docx` 文档时的补充规则，适用于所有工具链（docx-js、Python XML 操作、python-docx 等）。这些规则是实际踩坑后总结出来的，与官方 `docx` skill 独立——官方 skill 升级不会影响这些自定义规则。

## XML 结构安全

- **只改 `<w:t>` 内的文本，不动 XML 结构。** WPS 文档存在非标准嵌套，任何正则结构改动（删 `<w:r>`、移标签）都会级联破坏文档。结构性变更通过 Python 逐段重建文档实现。
- **WPS 非标准 XML 修复用 token-stream 栈式解析器**，不用计数器法或简单 tag 替换（均已验证失败）。
- **禁止用 `xml.etree.ElementTree` 序列化输出 Word XML。** ET 写出会丢失命名空间前缀（`<w:p>` 变 `<p>`），一律用字符串操作。

## unpack / pack 工作流陷阱（用 docx skill 的脚本时）

- **`pack.py` 在中文文件路径下必须加 `--validate false`。** 否则 GBK 解码错误会让脚本报 `Validation failed for ...`，但**实际 docx 已经成功生成**——属于"看似失败实际成功"。第二次跑加 `--validate false` 就过了。直接默认加这个标志省事。
- **`unpack.py` 默认会 merge 相邻 runs**（如 `merged 285 runs`），让整段文本落到单个 `<w:t>` 里。这对后续做字符串替换很关键——同一段中文不会被拆成 `<w:r>...</w:r><w:r>...</w:r>` 三段。**第二次 unpack 同一已合并文件时，merge 0 runs，是正常的**。
- **中文文件名导致 unpack/pack 输出乱码**（终端按 GBK 解码 UTF-8 文件名），但**实际操作正常**，文件已生成。看到 `Unpacked ../���Ʊ�ҵ...` 不要慌。
- **pandoc 转 markdown 不渲染 OMML 公式**（`<m:oMath>` 会显示为空白）。**不要据此判断公式丢失**，应该 grep `<m:oMath>` 计数对比 unpack 前后是否相等。Word 里打开看渲染才是 ground truth。

## 用 Edit 工具改 docx XML 时的注意事项

直接 unpack 后用 Edit 改 `word/document.xml`（避免写 Python 脚本）的几条铁律：

- **弯引号 `""''`（U+201C/D, U+2018/9）是字面字符**，old_string 必须包含完整字符。从 pandoc 输出贴回来的中文里都是 U+201C/D，不能写成 ASCII `"`。
- **长 `<w:t>` 块用 unique 子串替换，不要整块 `<w:t>...</w:t>` 替换**。摘要、长段落往往整段在一个 `<w:t>` 里，整块替换 old_string 太长容易 false negative。挑一句独有的中短子串切。
- **含引用编号的段落是多 run 拼接**——每个 `[N]` 引用是独立 `<w:hyperlink>` 块，前后各一个 `<w:r><w:t>...</w:t></w:r>`。改写要按 run 边界拆，先 Read 看 XML 结构再 Edit。

### 插入 / 保留引用编号（hyperlink XML 模板）

切分原 `<w:t>` 文本，关闭当前 run，插入 hyperlink 块，再开新 run（rPr 必须与原 run 一致）：

```xml
<w:t>原文前半</w:t>
      </w:r>
      <w:hyperlink w:anchor="ref_N" w:history="1">
        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>
          <w:sz w:val="24"/><w:vertAlign w:val="superscript"/></w:rPr>
          <w:t xml:space="preserve">[N]</w:t></w:r>
      </w:hyperlink>
      <w:r><w:rPr>...same as original run...</w:rPr>
        <w:t xml:space="preserve">原文后半</w:t>
```

### 代码片段里的方括号下标会被误识别为引用编号

学校论文模板有时会扫描所有 `[N]` 自动包成 `<w:hyperlink w:anchor="ref_N">`。结果 Python 代码里的 `coco_eval.stats[0]`、`stats[1]` 这种数组下标也被错误转成上标 hyperlink，要么指向不存在的 `ref_0`，要么误链到无关的 `ref_1`。

**修复**：把这两段 hyperlink 整体替换为普通 `<w:r>`，去掉 `<w:vertAlign w:val="superscript"/>`，让 `[0]` `[1]` 回到正常正文行内。**检测方法**：grep `w:anchor="ref_0"`（论文从 ref_1 起编号，ref_0 几乎一定是 bug）。

## python-docx 直接 XML 操作陷阱

绕过 `doc.add_paragraph()` 等 API、直接拼 XML 再插入 body 时，下面这几个坑都已经踩过。

### `body.append(p_elem)` 会把段落放到 `<w:sectPr>` 之后，Word 不渲染

`doc.element.body.append(p_elem)` 把段落粘到 body **末尾**——但 body 末尾通常是 `<w:sectPr>` 章节属性元素，**位于最后一个 sectPr 之后的段落 Word 当作章节外的孤儿，静默丢弃不渲染**（XML 里看得到，PDF/Word 里整段消失，无报错）。

正确写法：

```python
from docx.oxml.ns import qn

body = doc.element.body
last_child = body[-1] if len(body) else None
if last_child is not None and last_child.tag == qn("w:sectPr"):
    last_child.addprevious(p_elem)   # 插在最后的 sectPr 之前
else:
    body.append(p_elem)
```

`doc.add_paragraph()` 自带这个保护，问题只出在直接拼 XML + `body.append` 的代码上（典型场景：插入自定义 OMML 公式段落、绕过 python-docx 渲染特殊样式）。诊断方法：grep `<w:sectPr` 看出现位置，确认目标段落在最后那个之前。

### `parse_xml(fragment)` 要求 fragment 自带用到的命名空间声明

python-docx 的 `parse_xml` 用 lxml 解析，fragment 不带命名空间前缀声明会直接 `XMLSyntaxError`。例如要插入 `<m:oMath>...</m:oMath>` 必须先给 fragment 加 `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"`：

```python
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
omath = "<m:oMath>...</m:oMath>"   # 比如 pandoc 输出
if "xmlns:m=" not in omath.split(">", 1)[0]:
    omath = omath.replace("<m:oMath", f'<m:oMath xmlns:m="{M_NS}"', 1)
paragraph._element.append(parse_xml(omath))
```

挂到父节点后 lxml 会自动把冗余的 xmlns:m 上提到 document 根，最终 XML 干净，不用担心声明重复。

### markdown 表格行切分要识别 `$...$` 数学区域

`inner.split("|")` 切表格单元格时，公式 `$\sum |P|$` 里的 `|` 会被误当列分隔符把单元格切碎。写一个跟 `$` 状态的解析器，math 区内的 `|` 当字面量：

```python
def split_table_cells(inner):
    cells, buf, in_math = [], [], False
    for ch in inner:
        if ch == "$":
            in_math = not in_math
            buf.append(ch)
        elif ch == "|" and not in_math:
            cells.append("".join(buf).strip())
            buf = []
        else:
            buf.append(ch)
    cells.append("".join(buf).strip())
    return cells
```

## 样式和编号

- **样式 ID 必须从目标文档 `word/styles.xml` 读取，不能猜。** styleId="1" 可能是 Normal 而非 Heading 1；先解压 .docx 查 `<w:name w:val="heading 1">` 对应的实际 ID。
- **标题编号统一手动写，不用 `<w:numPr>` 自动编号。** numPr 的格式（如 `%1.%2`）可能输出 "1.1"/"1.2" 而非预期的 "1"/"2"，与手动编号的 H2 冲突。
- **英文关键词行按检测模板写成 `Keywords: `。** 冒号用半角 `:`，后面 1 个半角空格；关键词之间用 `, `。如果空格位于某个 `<w:t>` 末尾，必须写 `xml:space="preserve"`，否则 Word/WPS 可能吞掉尾随空格。

### 公式编号与制表符排版（OMML）

学校模板常要求公式居中、编号靠右，最稳的 OOXML 形态是：同一个 `<w:p>` 里依次放 `w:tab`、直接子节点 `<m:oMath>`、再放 `w:tab + <w:t>（2-1）</w:t>`。修公式编号时：

- **用用户确认过的公式段作模板，但不要复制书签/批注 ID。** 可复制该段的 `<w:pPr>`、首个 tab run、尾部编号 run；保留每条公式自己的 `<m:oMath>` 和编号文本。不要把 `bookmarkStart/bookmarkEnd/commentRange*` 从模板行复制到其它公式行。
- **把内嵌编号移出公式。** 形如 `<m:t>...#（2−2）</m:t>` 的编号会参与公式排版；删除 `#（...）`（注意 U+2212 `−`），把编号放到尾部普通 `<w:t>`。
- **不要留下显示公式外壳。** 顶层 `<m:oMathPara>` 和单行顶层 `<m:eqArr>` 都可能让公式本体独占一行，导致右侧编号被挤到下一行。单行 `eqArr` 的安全修法是：若 `<m:oMath>` 只有一个 `<m:eqArr>`，且其中只有一个 `<m:e>`，把这个 `<m:e>` 的公式子节点提升为 `<m:oMath>` 的直接子节点，并丢弃该层的 `<m:ctrlPr>`。
- **验证不能只看 XML。** 结构检查至少确认：目标公式段数量正确、每段有 1 个直接 `<m:oMath>`、无 `<m:oMathPara>`、无直接 `<m:oMath>/<m:eqArr>`、无 `#（...）` 残留、制表位和两个 tab run 与模板一致。若本机有 Word COM，再枚举 `Paragraph.Range.OMaths.Count > 0` 的段落，用 `Range.ComputeStatistics(1)` 确认公式编号段是 1 行；不要用普通 Find 搜索编号，因为会先命中正文里的“如式（2-2）”引用。

## 中文字体、字号、颜色

- **中文文档统一字体规则：中文用宋体，英文和数字用 Times New Roman，颜色统一用纯黑。** 在每个 run 的 `<w:rPr>` 中显式写出：

  ```xml
  <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>
  <w:color w:val="000000"/>
  ```

  不要使用 `w:themeColor`（主题色在不同模板下可能渲染成灰色或深蓝）。不要只写 `w:hint="eastAsia"` 而省略 `w:eastAsia="宋体"`——hint 只是提示字符归属，不强制字体族。

- **中文文档正文统一用小四号字（12pt）。** 在每个正文 run 的 `<w:rPr>` 中写 `<w:sz w:val="24"/><w:szCs w:val="24"/>`（Word 的 `w:sz` 单位是半磅，24 = 12pt = 小四）。不要用 `sz="21"`（五号 10.5pt）或 `sz="28"`（四号 14pt）作为正文字号。
- **标题字号按层级递减**：H1/H2 = 30（15pt），H3 = 28（14pt），H4 = 24（12pt，与正文同字号但加粗）。

## 脚本中的中文引号陷阱（Python & JavaScript 通用）

中文弯引号 `""`（U+201C / U+201D）在源码中会被解释为字符串定界符，导致语法错误。**所有语言都要处理。**

### Python

- 双引号字符串中出现 `"..."` 会 SyntaxError。用 `\u201c`/`\u201d` 代替，或改用单引号包裹字符串。

### JavaScript / docx-js

- **双引号字符串中的中文弯引号同样会导致 SyntaxError。** 原因：Node.js 会将 `"` (U+201C) 或 `"` (U+201D) 在某些上下文下误判为字符串结束符（尤其在 .cjs 文件中与其他 Unicode 混合时）。
- **解决方案（按推荐顺序）：**
  1. 所有含中文文本的字符串一律用 `\uXXXX` 转义写法，彻底避免源码中出现非 ASCII 字符
  2. 用模板字符串（反引号 `` ` ``）包裹，弯引号在模板字符串内安全
  3. 将弯引号提取为常量：`const LQ = "\u201C"; const RQ = "\u201D";`，正文用模板字符串拼接
- **踩坑记录：** 用 Python 脚本批量替换 `.cjs` 文件中的 `\u201c` → `\\u201C` 并不可靠——Python 的 `\u201c` 匹配的是 Unicode 字符本身，而文件中可能混合了字面 Unicode 和已转义的形式，替换后 Node.js 仍然会在 JS 解析阶段还原字符。**最可靠的方式是在写文件时就全量使用 `\uXXXX` 转义。**

### 通用原则

- 封面文字替换前先打印所有 `<w:t>` 内容确认覆盖范围。同一词可能以独立节点和长字符串子串两种形式出现，需分别替换。
