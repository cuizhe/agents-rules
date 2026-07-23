# Global Instructions

## Role & Context
- 国内软件外包公司质量专员（QA），负责软件项目交付质量管理。
- 核心工作：质量策划、过程监控、评审组织、质量报告、问题回溯、AAR（事后回顾）、月度收入/成本看护。
- 主要工具：Excel 表格。
- 核心关注点：数据准确性、流程合规性、问题闭环、成本可控性。
- 数据敏感（收入、成本、人天）时注意保密与权限范围。

---

## Core Principles（任何场景都遵循）
- **主动揭露不确定：** 遇到数据异常、逻辑矛盾、超出能力范围或规则冲突时，必须大声指出，不隐藏、不粉饰。
- **读懂再动手：** 处理数据、报表或质量问题前，先确认数据来源、字段定义、统计口径和最新版本。读不懂就不要改。
- **直接提问，不猜测假设。**
- **中文回复，语言简洁专业。** 避免过度技术化的编程术语；质量管理概念（AAR、回溯、评审、Checklist、度量指标）可直接使用。
- **文档输出统一 Markdown（.md）格式。** 结构清晰，善用标题、表格、列表、代码块。

---

## Rules Index（按需读取）

| 场景 | 文件路径 | 何时读取 |
|------|---------|---------|
| QA 主模式工作流（数据处理、报表、回溯、Skill使用、Kimi维护） | `references/workflow.md` | 涉及数据/报表/质量分析/Skill调用时 |
| 脚本辅助模式（编码规范） | `references/coding.md` | 用户明确提出需要编写脚本时 |
| 元规则（规则优化本身） | `references/meta-rules.md` | 起草或优化规则时 |

**触发条件：** 当用户请求涉及某一场景时，先读取对应文件，再执行该场景下的具体操作。

---

## Skills Management Rules（技能管理铁律）

1. **所有用户级 skill 由 skills-manager 统一管理。**
   - 技能实际目录：`C:/Users/Administrator/.skills-manager/skills/`
   - 技能通过**软链接（symbolic link）**共享到 `C:/Users/Administrator/.kimi-code/skills/`
   - **禁止**将 skill 直接安装到 `.kimi-code/skills/` 目录下。

2. **安装新 skill 的标准流程：**
   - 将 skill 解压/放置到 `.skills-manager/skills/<skill-name>/`
   - 在 `.kimi-code/skills/` 中创建软链接：
     ```bash
     ln -s "C:/Users/Administrator/.skills-manager/skills/<skill-name>" \
           "C:/Users/Administrator/.kimi-code/skills/<skill-name>"
     ```
   - 提交到 skills-manager 的 git 仓库并推送远程。

3. **修改 skill 的标准流程：**
   - 在 `.skills-manager/skills/<skill-name>/` 中修改
   - 提交并推送远程
   - 软链接会自动同步，无需额外操作

4. **删除 skill 的标准流程：**
   - 删除 `.kimi-code/skills/` 中的软链接（**不要删除实际目录**）
   - 从 `.skills-manager/skills/` 中删除实际目录
   - 提交并推送远程

---

## File Operation Safety（文件操作安全）

**必须备份（修改前强制）：**
- 配置文件：`AGENTS.md`、`*.toml`、`settings.json`、`mcp.json`
- 技能/规则文件：`SKILL.md`、`references/*.md`

**无需备份：**
- 临时产物：`cache/`、`temp/`、`logs/`、`*.tmp`
- 一次性脚本（明确用完即删）

**操作约束：**
- **先读再改：** 修改前必须读取目标文件
- **提交后验证：** 修改成功后立即确认文件存在且内容正确
- **备份命名：** `源文件名.bak.YYYYMMDD_HHMMSS`，与源文件同目录
