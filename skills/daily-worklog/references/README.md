# 自动化工作日志填报系统

## 项目结构

```
daily-worklog/
├── SKILL.md                  # Skill 主入口文件，含完整使用指南
├── scripts/
│   ├── fill_worklog.py       # 主控脚本：解析输入、浏览器自动化、表单填单
│   ├── engine.py             # 匹配引擎、聚合引擎、业务规则处理
│   ├── init_task_db.py       # 一次性工具：Excel → JSON 转换器
│   ├── launcher.ps1          # Chrome 启动脚本（带调试端口 9222）
│   ├── project_index.py      # 项目索引管理：活跃项目清单、简称解析
│   ├── update_active_projects.py  # 月度活跃项目清单更新脚本
│   ├── test_adjust.py        # 凑8小时逻辑单元测试
│   └── debug_trace.py        # 浏览器交互诊断工具
├── assets/
│   ├── config.yaml           # 运行时配置（Chrome 路径、URL、节假日等）
│   ├── task_db.json          # 本地历史任务库（运行时自动更新）
│   ├── workitem_weights.yaml # 随机补充权重配置
│   └── active_projects.json  # 月度活跃项目清单
├── references/
│   ├── plan.md               # 详细设计文档
│   ├── history_workitems.md  # 历史库工作项汇总
│   ├── mapping_worksheet.md  # 关键词→工作项映射表
│   ├── system_workitems.md   # 系统完整工作项列表
│   ├── workitem_comparison.md # 工作项对比分析
│   └── *.md                  # 复盘记录与参考文档
└── evals/
    └── evals.json            # Skill 评估用例
```

## 使用步骤

### 1. 初始化本地任务库（仅需一次）

```bash
cd <skill-path>
python scripts/init_task_db.py
```

读取历史报工 Excel，生成 `assets/task_db.json`。

### 2. 启动 Chrome 调试模式

```powershell
& "D:\Program Files\RunningCheeseChrome\App\chrome.exe" --remote-debugging-port=9222
```

或使用脚本：
```powershell
python scripts/launcher.ps1
```

启动后**手动登录**目标系统。

### 3. 执行填单

**推荐方式（自然语言描述）：**

```bash
python scripts/fill_worklog.py --describe "项目级：中移信息，CP评审，1小时；江苏端木，评审，0.5小时 公共工时：云服务器续期，1小时" --preview
```

先 `--preview` 生成草稿，确认无误后再去掉 `--preview` 正式执行。

**其他模式：**

```bash
# 关键词匹配
python scripts/fill_worklog.py --keywords "中移信息 评审" --preview

# 完全随机生成
python scripts/fill_worklog.py --random full --preview

# 关键词 + 随机补充
python scripts/fill_worklog.py --keywords "评审 测试" --random supplement --preview
```

## 安全策略

- **永不自动提交**：脚本始终停留在"保存草稿"阶段，需用户手动核对后点击提交。
- **已有记录拦截**：目标日期已有填报记录时暂停，不自动覆盖/追加。
- **日期拦截**：不允许填报未来日期。
- **关键操作前自动截图**，失败后保留现场。
- **工时自动取整**到 0.5 倍数，**日志超过 500 字自动截断**。

## 核心逻辑

1. **解析**：`--describe` 参数由 `parse_daily_description()` 解析为结构化记录
2. **匹配**：关键词在 `search_text` 中 OR 匹配
3. **聚合**：按 `(申报类型, 项目名称, 工作项)` 分组，工时取平均并向上取整到 0.5 倍数
4. **对齐**：总和始终严格等于 8.0H，不足时补充，超出时压缩
5. **填单**：连接已启动的 Chrome（CDP），自动完成表单填写

