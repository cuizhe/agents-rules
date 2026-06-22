# Codex Rules

## 配置文件路径

| 文件 | 本机路径 |
|------|---------|
| 全局 AGENTS.md | `C:\Users\Administrator\.codex\AGENTS.md` |

## 备份说明

- 修改 `~/.codex/AGENTS.md` 后，手动复制到本目录并提交。
- 提交命令：
  ```bash
  cd ~/agents-rules
  git add codex/
  git commit -m "update: codex AGENTS.md"
  git push origin main
  ```

## Codex CLI 信息

- 安装方式：`npm install -g @openai/codex`
- 更新方式：`npm update -g @openai/codex`
- 配置文件：`~/.codex/AGENTS.md`
