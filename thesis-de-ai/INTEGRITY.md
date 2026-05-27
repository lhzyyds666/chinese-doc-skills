# 定稿前完整性检查清单

降 AI 痕迹只是定稿的一部分。Word 模板里常残留下面这些非语言层面的问题，老师评审时同样容易扣分。这一步在每轮降 AI 之后、最终提交之前各做一次。

---

## 1. 交叉引用完整性

**检查**：

```bash
# 列出所有被引用的 ref_N
grep -oE 'w:anchor="ref_[0-9]+"' word/document.xml | sort -u

# 列出参考文献列表里所有 bookmark
grep -oE 'w:name="ref_[0-9]+"' word/document.xml | sort -u
```

**期望**：被引用的集合 ⊆ bookmark 集合，且没有 ref_0（除非确实有 [0] 编号）。

**常见 bug**：
- 文档模板把代码片段里的 `stats[0]` `coco.stats[1]` 这种 Python 数组下标错误识别为参考文献编号，自动包成 `<w:hyperlink w:anchor="ref_0">`，指向不存在或无关的文献
- 修复：把那段 hyperlink 块替换为普通 `<w:r>`，去掉 `<w:vertAlign w:val="superscript"/>`

---

## 2. 未引用文献

**检查**：bookmark 集合 \\ 被引用集合 = 未引用文献编号

**处理**（按用户偏好）：
- 选项 A：删除未引用文献，重新连号（高风险，要同步改下游所有引用编号）
- 选项 B：在正文里补上引用（推荐，工作量可控）
- 选项 C：留给导师判断

**补引用的位置启发**：
- benchmark 数据集（MOT16/MOT17/MOT20/UA-DETRAC）→ 评测指标体系章节
- 经典论文（Kalman 1960、Focal Loss、YOLOv3）→ 对应方法首次出现处
- 同类方法（JDE、TransTrack、QDTrack）→ 算法综述章节并列提及

**插入引用的 XML 模板**见 [SKILL.md](SKILL.md#插入保留引用编号hyperlink-写法)。

---

## 3. 模板批注清理

**检查**：

```bash
grep -c "w:author=" word/comments.xml
```

**典型情况**：从学校提供的论文模板继承的旧批注（往往是模板作者多年前留下的修改建议），作者名是陌生人。

**处理**：
- 简单方式：让用户在 Word 里"显示批注 → 全部删除"
- 脚本方式：清空 `comments.xml` 内 `<w:comment>` 节点，并在 `document.xml` 里删除所有 `<w:commentRangeStart>` `<w:commentRangeEnd>` `<w:commentReference>` 标记

---

## 4. 图占位符

**检查**：

```bash
grep -n "此处缺图\|TODO.*图\|FIGURES_TODO" word/document.xml
```

**处理**：
- 如有图：用户提供图片路径后，按 docx skill 的图像插入流程（见 docx skill 的"Editing → Images"）
- 暂无图：删除占位提示文本，保留图题段落（"图N-N ..."）等待后补

---

## 5. 封面模板默认值

**检查**：grep 模板默认的姓名/学号/导师值，例如：

```bash
grep -n "张  三\|李  四\|1707010101\|学生姓名：" word/document.xml
```

**处理**：从致谢段反推真实导师姓名（"感谢我的指导教师 X 老师"），从 docx 文件名反推真实学生姓名，学号必须由用户提供。

---

## 6. 公式完整性

**检查**：unpack 前后 `<m:oMath>` 计数应相等

```bash
grep -c "m:oMath" word/document.xml
```

**注意**：pandoc 转 markdown 不渲染 OMML 公式（显示为空白），不要据此判断公式丢失。以 Word 渲染或 `<m:oMath>` 计数为准。

**真要确认 Word 里能看：** 用 Word COM 转 PDF 看一眼

```python
import win32com.client
word = win32com.client.Dispatch('Word.Application')
doc = word.Documents.Open(r'<absolute-path>')
doc.SaveAs(r'<absolute-path>.pdf', FileFormat=17)
doc.Close(); word.Quit()
```

---

## 7. 参考文献格式（GB/T 7714）

**抽样检查**几条参考文献，确认格式一致。常见三类：

```
[N] AUTHOR A, AUTHOR B. Title[J]. Journal, year, vol(issue): pages.
[N] AUTHOR A, AUTHOR B. Title[C]//Conference Name. City: Publisher, year: pages.
[N] AUTHOR A. Title[J/OL]. arXiv preprint, year, arXiv:XXXX.XXXXX. URL.
```

**常见问题**：作者大小写不统一、期刊名缩写不统一、年份/页码标点不统一、URL 缺失。

---

## 8. 目录页码

**检查**：在 Word 里更新目录字段（F9 或右键"更新域"）。脚本无法验证，必须手动。

---

## 完整性扫描一键脚本（可选）

把 grep 命令封装成一个脚本：

```bash
#!/usr/bin/env bash
# thesis-integrity-scan.sh <unpacked_dir>
DIR="$1"
echo "=== 引用 vs bookmark 对比 ==="
diff <(grep -oE 'w:anchor="ref_[0-9]+"' "$DIR/word/document.xml" | sort -u) \
     <(grep -oE 'w:name="ref_[0-9]+"' "$DIR/word/document.xml" | sort -u | sed 's/w:name/w:anchor/')
echo "=== 模板批注作者 ==="
grep -oE 'w:author="[^"]+"' "$DIR/word/comments.xml" 2>/dev/null | sort -u
echo "=== 图占位 ==="
grep -n "此处缺图" "$DIR/word/document.xml"
echo "=== 封面模板默认值 ==="
grep -nE "张  三|李  四|1707010101" "$DIR/word/document.xml"
echo "=== OMML 公式数 ==="
grep -c "m:oMath" "$DIR/word/document.xml"
```
