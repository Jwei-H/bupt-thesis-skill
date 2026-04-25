---
name: bupt-thesis-helper
description: >-
  Use this skill for BUPT thesis workflows: run structured Markdown checks,
  review heading trees against project conventions, and export any specified
  Markdown thesis file to DOCX with the bundled cover and integrity pages.
---

# bupt-thesis-helper

## Overview

用于北邮论文 Markdown 检查与 DOCX 导出。

适用场景：

- 检查任意论文 Markdown 文件的结构、标题、摘要、图表题注与编号
- 提取完整多级标题，供 LLM 按项目约定做语义复核
- 将指定 Markdown 文件导出为 DOCX，并拼接封面与诚信声明

当前工具链为 JS-only：检查、正文导出、封面组装均使用 Node.js 脚本完成。

## Agent Rules

1. **先检查，后导出。** 除非用户明确要求跳过，否则先运行 `check_markdown.js`。
2. **非必要不要阅读全文。** 优先使用检查脚本输出、标题树、局部行号与局部片段定位问题；只有在脚本结果不足以判断时，才阅读原文局部内容。
3. **多级标题的“含义约定”由 LLM 复核。** 检查脚本只能抽取和校验标题结构；对于“这个标题层级是否符合我们当前约定”的判断，必须结合 `headings` 输出由 LLM 继续复核。
4. **检查完先汇报，再询问是否修复。** 当脚本检查完成后，先总结错误/警告与标题树结论，再询问用户是否开始修复，不要默认直接改论文。
5. **始终显式指定文件路径。** 所有脚本必须明确传入输入的 Markdown 路径，系统不提供任何默认文件名回退。调用脚本前，必须首先确认要处理的 Markdown 文件路径。
6. **关注参考文献余量与补充 SOP。** 当用户显式要求补充参考文献，或者你观察到参考文献数量低于 20 个、近三年文献占比低于 30% 时，应主动询问用户是否需要帮忙补充文献。如果用户确认，在操作前**必须**前往查阅并遵循 \`bupt-thesis-helper/references/add-references.md\` 规范。
7. **目录页码异常先提示手动更新域。** 若用户反馈目录页码异常、但目录跳转仍正常，不要立刻判定导出逻辑错误；应先简要提示用户在 Word 中手动“更新整个目录”或全选后更新域重试，因为这类问题常与 Word 打开后的分页/域刷新时机有关。

## Dependencies

运行该 Skill 下的工具链需要配置 Node.js 并在当前上下文环境安装必要的 npm 包。作为 Agent，你应该自主判断或提示用户是否已安装这些包；如果环境缺失依赖，直接通过系统命令在恰当目录自行安装：

```bash
npm install docx jszip @xmldom/xmldom
```

依赖用途：

- `docx`：正文 DOCX 生成
- `jszip`：DOCX 包读写
- `@xmldom/xmldom`：封面注入与 XML 后处理

## Commands

### 1. 结构检查

```bash
node scripts/check_markdown.js <markdown-path>
```

需要结构化标题树与问题清单时：

```bash
node scripts/check_markdown.js <markdown-path> --json
```

### 2. 只生成正文 DOCX

```bash
node scripts/generate_thesis.js --input <markdown-path> --output <body-docx-path>
```

说明：

- `--input` 必填，不提供任何固定文件名回退
- `--output` 可省略，默认输出为“输入 Markdown 同目录下的同名 .docx”

### 3. 只组装封面与正文

```bash
node bupt-thesis-helper/scripts/compose_docx.js --cover <cover-docx-path> --body <body-docx-path> --output <final-docx-path> --cover-data <cover-json-path>
```

### 4. 一键导出最终 DOCX

```bash
node scripts/md2doc.js --input <markdown-path> --output <final-docx-path>
```

可选参数：

- `--cover <cover-docx-path>`
- `--cover-data <cover-json-path>`
- `--skip-check`
- `--force`

说明：

- 若未指定 `--output`，默认输出为“输入 Markdown 同目录下的同名 `.docx`”
- 若未指定 `--cover-data`，脚本会在输入 Markdown 同目录下自动补出 `<markdown-name>.cover.json` 模板

## What the Check Script Covers

`check_markdown.js` 会做这些事情：

1. 提取完整多级标题
2. 校验一级/二级/三级标题格式、层级与编号顺序
3. 检查二、三级标题后是否先出现正文说明文字
4. 检查中文摘要与英文摘要是否满足 `<br> + 关键词` 规则
5. 检查普通图片是否存在、是否有下方题注
6. 检查普通表格是否有上方题注
7. 检查表格单元格内嵌图片是否存在、是否带同单元格题注
8. 检查图表编号唯一性，以及编号章号与当前位置是否一致
9. 将 Mermaid / PlantUML 类代码块按图片对象纳入题注检查

**重要：** 标题“结构正确”不等于“层级语义正确”。拿到 `headings` 输出后，LLM 仍需结合当前项目约定判断：这些多级标题是否真的该放在这一层。

## Convention over Configuration (约定大于配置)

本技能的一切检查、导出逻辑与特定的论文模板排版，都强依赖于深度的“命名约定与格式约定”（如：特定的专用标题名、固定的图表题注格式等）。
作为 Agent，在执行涉及论文修改、查错、调整排版和结构生成的操作前，你必须前往研读 \`bupt-thesis-helper/references/markdown-writing-spec.md\`，这是所有约定逻辑的唯一真理源。

## Recommended Workflow

### 场景 A：快速自检

1. 运行 `check_markdown.js <markdown-path> --json`
2. 优先阅读 `issues` 和 `headings`
3. 由 LLM 复核标题层级是否符合当前约定
4. 向用户汇总问题，并询问“是否开始修复”
5. 只有在需要定位具体上下文时，再阅读原文局部片段

### 场景 B：准备导出

1. 运行检查
2. 确认无阻断错误，或用户明确允许 `--force`
3. 运行 `md2doc.js --input <markdown-path> --output <final-docx-path>`
4. 打开结果，重点检查目录、标题、图表题注、公式、封面填写区与诚信声明页

### 场景 C：只想复核标题树

1. 运行 `check_markdown.js <markdown-path> --json`
2. 直接读取 `headings`
3. 非必要不要通读整篇论文

## Resources

- `scripts/check_markdown.js`：结构化检查与标题树提取
- `scripts/generate_thesis.js`：正文 DOCX 生成
- `scripts/compose_docx.js`：封面/声明与正文组装
- `scripts/md2doc.js`：总入口，串联检查、正文生成、最终组装
- `references/markdown-writing-spec.md`：写作与导出规则说明
- `references/add-references.md`：参考文献补充 SOP
- `assets/论文封面+诚信声明.docx`：封面模板
- `assets/thesis.cover.example.json`：封面信息模板

## Boundaries

- 仅面向当前仓库的北邮论文导出链路，不作为通用 Markdown 论文工具
- 若检查存在错误，默认阻止导出；只有用户明确需要时才使用 `--force`
- 当前正文导出唯一来源为 skill 内的 `scripts/generate_thesis.js`
- 与论文内容语义相关的最终判断，优先依赖检查脚本输出 + LLM 复核，而不是直接通读全文
