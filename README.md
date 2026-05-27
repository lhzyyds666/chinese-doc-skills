# chinese-doc-skills

两个面向中文文档处理的 [Claude Code](https://docs.claude.com/en/docs/claude-code) / Agent Skills。

## 安装

把对应目录放到你的 skills 目录下（如 `~/.claude/skills/`）：

```bash
git clone https://github.com/lhzyyds666/chinese-doc-skills.git
cp -r chinese-doc-skills/docx-wps-fixes ~/.claude/skills/
cp -r chinese-doc-skills/thesis-de-ai   ~/.claude/skills/
```

## Skills

### `docx-wps-fixes`

创建或编辑中文 `.docx` 文档时的补充规则，是官方 `docx` skill 的**强制搭档**（需与 `docx` 在同一条消息里并行加载）。覆盖：

- 中文字体规则（宋体 + Times New Roman + 纯黑）、小四号正文字号
- JS/Python 源码里中文弯引号导致 SyntaxError 的转义
- WPS 非标准 XML 的安全处理
- 与官方 `docx` skill 不同的结构化编辑规则（`body.append` 与 `<w:sectPr>`、`parse_xml` 命名空间、用 Edit 工具改 XML 等踩坑总结）

文件：`SKILL.md`

### `thesis-de-ai`

中文学位论文（本科/硕士）降低 AI 写作痕迹的多轮修订工作流。识别 AI 模板化套话、工整三四段并列、过满总结、AI 归纳口吻、投诚式解释、过度像教材、隐性口语词、重复词共八类痕迹，并给出改写原则、风险句清单、unpack-edit-repack 操作流程，以及定稿前完整性检查清单。

文件：`SKILL.md`、`PATTERNS.md`（八类痕迹细则）、`INTEGRITY.md`（学术诚信边界）
