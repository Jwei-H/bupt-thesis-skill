## bupt-thesis-skill

提供对于bupt本科毕设论文的 Markdown 自动化检查与 DOCX 渲染导出工具链（以 Agent Skill 形式提供）。

### 功能简述

- **结构化检查**: 对论文的标题层级、图表题注与编号、摘要格式等进行详尽的规范校验。
- **一键导出**: 将合规的 Markdown 正文渲染为排版良好的 DOCX 文件。
- **辅助论文修改**：辅助润色、调整论文格式，包括补充参考文献等。

### 目录结构

- `bupt-thesis-helper/scripts/`：核心 Node.js 脚本（包含结构检查、DOCX 渲染及封面组装等逻辑）。
- `bupt-thesis-helper/assets/`：封面及声明的 DOCX 模板和配置例子。
- `bupt-thesis-helper/references/`：关于排版、参考及各种写作规约的参考文档。
- `bupt-thesis-helper/SKILL.md`：详细描述了该 Skill 的工作流、调用指引以及前置条件，供 Agent 查阅参考。

### 快速使用

#### 1. 接入 Skill

将此技能仓库克隆或下载到本地，将整个目录配置至你的 Agent 环境中（具体方式视你使用的 Agent 产品而定，例如放入特定的技能文件夹、加入工作区上下文或通过指令导入）。
Agent 会自动读取 `bupt-thesis-helper/SKILL.md` 中的元数据、意图触发说明以及工作流设定。

#### 2. 对话调用

完成接入后，你只需用日常对话的方式向 Agent 下达处理指令。例如：
- “使用 bupt-thesis-helper 检查一下当前的 Markdown 论文格式是否有问题。”
- “把我的这篇论文导出为 DOCX 文档。”
- “提取本论文的多级标题树，检查一下大纲层级结构是否符合规范。”