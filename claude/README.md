# Claude Rules

## 配置文件路径

| 文件 | 本机路径 |
|------|---------|
| 全局 CLAUDE.md | `C:\Users\Administrator\.claude\CLAUDE.md` |

## 备份说明

- 修改 `~/.claude/CLAUDE.md` 后，手动复制到本目录并提交。
- 提交命令：
  ```bash
  cd ~/agents-rules
  git add claude/
  git commit -m "update: claude CLAUDE.md"
  git push origin main
  ```

## Claude Code 信息

- 安装方式：`npm install -g @anthropic-ai/claude-code`
- 更新方式：`claude update` 或 `npm update -g @anthropic-ai/claude-code`
- 配置文件：`~/.claude/CLAUDE.md`

---

> 当前状态：占位文件，待后续补充 Claude 的 CLAUDE.md 内容。
