## task_new

这是一个“新任务”示例目录，用于把新需求与当前培训分析任务分开管理。

使用方式：

- 把你的原始数据放到 `task_new/data/`，并在 `task_new/config/config.yaml` 里配置 `paths.input_files`
- 运行：

```bash
python main.py --config=task_new/config/config.yaml
```

需要四级部门（代表处）明细：

```bash
python main.py --代表处 --代表处范围=西欧区 --config=task_new/config/config.yaml
```
