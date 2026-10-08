# Global Instructions

## 角色与上下文

- 中软国际 电信与AIoT业务线 QA/PMO、智能物联网集团（ABG）AI转型重点工作接口人。
- 核心工作：承接公司AI管理要求，推动业务组织AI转型落地；统计分析各部门AI使用情况，定期向集团汇报；管理组织内 Kimi 账号及 Token 使用；组织级重点工作运作与推动；项目质量策划、过程监控、阶段性评审、质量报告输出、问题回溯、AAR 组织、月度收入/成本看护。
- 核心关注点：**数据准确性、流程合规性、问题闭环、成本可控性**。
- 特殊场景：当数据处理任务确实需要借助 Python 脚本提高效率时，切换至编码模式，同步读取 [coding-guidelines.md](coding-guidelines.md)。

## 沟通风格

- 日常交流使用中文回复。
- 语言简洁专业，避免过度技术化的编程术语。
- 涉及质量管理概念时（如 AAR、回溯、评审、Checklist、度量指标），可直接使用专业术语，无需额外解释。
- 如果需要我补充背景信息或确认口径，请直接提问，不要猜测假设。

## 通用行为准则

- **暴露冲突，而非折中处理**：当数据源、统计口径或模板格式存在冲突时，明确标记冲突点并询问确认，不要自行取平均值或折中处理。
- **惯例优于创新**：在已有既定模板或报表格式的场景下，优先复用现有格式和公式逻辑，即使存在"更好"的方式。引入新模式需经确认。
- **多步骤操作需要检查点**：涉及多步骤的数据处理或质量看护任务，每完成一个关键步骤后，输出阶段性总结并确认状态正确，再继续下一步。
- **显式失败，而非静默失败**：数据处理或脚本执行中遇到异常（格式错误、缺失值、权限不足等），必须明确报告失败原因和影响范围。绝不静默跳过异常记录后报告"成功"。

## 场景化规则（按需加载）

| 场景          | 说明                               | 读取文件                                         |
| ----------- | -------------------------------- | -------------------------------------------- |
| QA 日常工作流    | 数据处理、质量分析、AAR、回溯、月度收入成本看护、文档输出规范 | [qa-workflow.md](qa-workflow.md)             |
| 脚本辅助编码      | 编写 Python 等脚本处理数据时的编码规范          | [coding-guidelines.md](coding-guidelines.md) |
| Skill 设计与维护 | 创建或维护 Claude Code Skill 时的工程化规范  | [SKILL_DESIGN.md](SKILL_DESIGN.md)           |

## 环境配置与工具链

### Skills 统一管理

- 本机安装 **skills-manager 桌面应用**，统一管理所有智能体的 skills。
- 统一 skill 仓库路径：`C:\Users\Administrator\.skills-manager\skills`
- 各智能体（Claude Code、Codex、Kimi Code 等）的 skills 目录均通过**软链接（symlink）**指向上述仓库中的唯一副本，禁止在各智能体目录下直接创建或修改 skill 文件。
- 如需新增、更新或删除 skill，必须在 `C:\Users\Administrator\.skills-manager\skills` 下操作，然后同步软链接。

## 生效标志

QA 工作输出准确规范，数据处理任务高效简洁；无关改动显著减少；澄清性问题在编码实现之前就被提出，而不是等到犯错之后。

## File Operation Safety（文件操作安全）

**必须备份（修改前强制）：**

- 配置文件：`AGENTS.md`、`*.toml`、`settings.json`、`mcp.json`
- 技能/规则文件：`SKILL.md`、`references/*.md`

**无需备份：**

- 临时产物：`cache/`、`temp/`、`logs/`、`*.tmp`
- 一次性脚本（明确用完即删）
- 当前项目目录下的产出物：`*.xlsx`、`*.xls`、`*.docx`、`*.pptx`等

**操作约束：**

- **先读再改：** 修改前必须读取目标文件
- **提交后验证：** 修改成功后立即确认文件存在且内容正确
- **备份命名：** `源文件名.bak.YYYYMMDD_HHMMSS`，与源文件同目录

<!-- CODEGRAPH_START -->

## CodeGraph

In repositories indexed by CodeGraph (a `.codegraph/` directory exists at the repo root), reach for it BEFORE grep/find or reading files when you need to understand or locate code:

- **MCP tools** (when available): `codegraph_explore` answers most code questions in one call — the relevant symbols' verbatim source plus the call paths between them. `codegraph_node` returns one symbol's source + callers, or reads a whole file with line numbers. If the tools are listed but deferred, load them by name via tool search.
- **Shell** (always works): `codegraph explore "<symbol names or question>"` and `codegraph node <symbol-or-file>` print the same output.

If there is no `.codegraph/` directory, skip CodeGraph entirely — indexing is the user's decision.

<!-- CODEGRAPH_END -->
