---
name: thesis-de-ai
description: 中文学位论文（本科/硕士）降低 AI 写作痕迹的多轮修订工作流。识别 AI 模板化套话、工整三四段并列、过满总结、AI 归纳口吻、投诚式解释（"未做任何编造"等危险句）、过度像教材、隐性口语词、重复词共八类痕迹；给出改写原则、风险句清单、unpack-edit-repack 操作流程、引用 hyperlink 拆分写法，以及定稿前批注/封面/引用/公式/参考文献完整性检查清单。Use when 用户要降低中文论文的 AI 生成感、消除"第一/二/三/四，针对…设计了/构建了/提出了"类四段模板、改写"未做任何编造""抠出"类投诚句、处理改写后残留的"试出来的""跑下来""吐出""再压一压"等口语词、或核查交叉引用与未引用文献。
---

# 中文学位论文降 AI 痕迹

## 核心原则

1. **绝对不动**：所有数值、模块名、代码文件名、引用编号、公式编号、章节结构
2. **不拔高贡献**：不加不存在的实验、数据、对比或结论
3. **多轮迭代，每轮 5–10 句**：单次大改容易把"AI 模板"换成"降 AI 模板"，反而出现新痕迹（如"试出来的""减法版""跑下来"）。让用户阅读后再补改，3–5 轮收敛
4. **保留学生口吻**：可以有"开发过程中实际遇到过…加了对应分支"这种工程语感，比纯学术腔更自然
5. **降调而非删除**："证明了"→"实验结果表明"；"显著"→具体百分点；"为…提供支撑"→"与…有关"

## 工作流（每轮）

1. **扫描或接收清单**——若全文扫描，对照 [PATTERNS.md](PATTERNS.md) 八类痕迹 grep 全文，把 top 3–10 高风险段落+原句列给用户确认；若用户主动给出改写清单，直接进入第 2 步
2. **unpack docx**：`python C:/Users/User/.claude/skills/docx/scripts/office/unpack.py <input.docx> _ws/unpacked/`
3. **改写**：在 `_ws/unpacked/word/document.xml` 里只改 `<w:t>` 内文本，不动 XML 结构。含引用 hyperlink 的段落是多 run 拼接，需按 run 边界拆分；插入新引用见下方
4. **repack**：`python C:/Users/User/.claude/skills/docx/scripts/office/pack.py _ws/unpacked/ <output.docx> --original <input.docx> --validate false`（中文文件名导致 GBK 验证误报，但实际文件正常生成）
5. **验证**：
   - `<m:oMath>` 计数前后相等（公式数量未变）
   - `w:hyperlink w:anchor="ref_*"` 集合 ⊆ `w:name="ref_*"` 集合（无悬空引用）
   - `pandoc <output.docx> -o tmp.md` + grep 旧痕迹关键词应全部为空
6. **清理**：删除 `_ws/unpacked/` 与中间 .md，保持工作目录整洁

## 八类 AI 痕迹（详见 [PATTERNS.md](PATTERNS.md)）

| # | 类型 | 关键词 |
|---|------|--------|
| A | 模板化套话 | "针对…设计了…" "本文构建/提出/达到" "为…奠定基础" "充分说明" "全面" |
| B | 工整三/四段并列 | "第一/二/三/四，针对…" |
| C | 过满总结 | "各项均达到或超过预设阈值" "起到支撑作用" "证明了" "显著提升" "效果显著" |
| D | AI 归纳口吻 | "综合看" "可以看到" "整体上" "总的来说" "进一步" "横向看下来" |
| E | **投诚式解释 ⚠️** | "未做任何编造" "数据真实可靠" "抠出" — **必须删除**，比 AI 痕迹本身更危险 |
| F | 过度像教材 | 把"为什么本文这么做"写成方法本身的科普 |
| G | 隐性口语词 | "搬到" "跑下来" "吐出" "再压一压" "试出来的" "更值得相信" "行得通" "减法版" "看起来是合理的" |
| H | 重复词 | 同句中"在本文数据集上"出现两次等 |

## 插入/保留引用编号（hyperlink 写法）

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

## 中文 docx 编辑陷阱

- 弯引号 `""''`（U+201C/D, U+2018/9）在 Edit 工具中是字面字符，old_string 必须包含完整字符
- 长 `<w:t>` 段落用 unique 子串替换，不要整块替换
- 含引用编号的段落是 5+ runs 拼接，先 Read 看 XML 再 Edit
- pandoc 转 md 不渲染 OMML 公式（显示空白），不要据此判断公式丢失，应以 `<m:oMath>` 计数为准
- pack.py 在中文文件路径下必须加 `--validate false`，否则 GBK 编码错误会让脚本退出（但实际文件已成功生成）

## 定稿前完整性检查

详见 [INTEGRITY.md](INTEGRITY.md)。重点：
- 悬空引用 / 未引用文献 / 代码下标 `[0]` `[1]` 被误识别为引用
- 模板批注（`comments.xml` 里非本人作者的旧批注）
- 图占位 / 封面模板默认值 / 参考文献格式一致性
