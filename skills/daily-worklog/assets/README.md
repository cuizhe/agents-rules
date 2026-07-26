# assets 目录

该目录用于存放 skill 运行时数据与配置文件：

- `config.yaml`：运行时配置（Chrome 路径、目标 URL、节假日、时间过滤规则）
- `task_db.json`：历史任务库（每次填单后自动更新）
- `workitem_weights.yaml`：随机补充权重配置
- `active_projects.json`：月度活跃项目清单（由 `scripts/update_active_projects.py` 生成）

运行时脚本通过 `Path(__file__).parent.parent / "assets"` 定位本目录。
