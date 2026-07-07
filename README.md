# 培训数据分析及 PPT 汇报工具

该工具用于处理从培训平台导出的 Excel 原始数据，自动计算各区域/小组的学习进度，并生成汇报 PPT。

## 🚀 快速开始

### 1. 准备环境
确保您的电脑已安装 Python，并安装以下依赖：
```bash
pip install pandas python-pptx openpyxl pyyaml
```

### 2. 准备数据
1. 从培训平台导出原始数据 Excel 文件。
2. 将该文件放入项目的 `data/` 目录中。
3. 将文件重命名为 `raw_learning_data.xlsx`（或者在 `config/config.yaml` 中修改路径）。

### 3. 准备 PPT 模板
1. 将您的 PPT 模板放入 `reports/templates/` 目录。
2. 命名为 `template.pptx`。

### 4. 运行分析
在项目根目录运行：
```bash
python main.py
```

只分析指定大区/小组：
```bash
python main.py --小组=西欧区
```

不发送邮件（只生成 Excel/PPT）：
```bash
python main.py --no-email
```

### 新建一个独立任务（推荐方式）
把新需求与现有任务分开：新建一个 `task_xxx/` 目录，并提供独立配置文件，然后用 `--config` 运行。

仓库已提供示例任务目录：`task_new/`，你可以直接复制改名为你的任务名。

```bash
python main.py --config=task_new/config/config.yaml
```

## 📊 数据分析维度
工具会自动根据以下维度进行汇总统计：
- **学员数量**：各小组的去重人数。
- **人均学习时长**：各小组学员的平均观看时间。
- **学习完成率**：状态为“已完成”的学员占比。
- **考试通过率**：状态为“通过”的学员占比。

## 📂 项目结构
- `main.py`: 程序总入口。
- `config/config.yaml`: 配置文件。
- `modules/data_analyzer.py`: 核心分析逻辑。
- `reports/ppt_generator.py`: PPT 生成逻辑。
- `data/`: 存放原始数据。
- `outputs/`: 存放生成的分析 Excel 和 PPT 汇报。
- `task_customer/`: 客户培训覆盖分析任务（独立配置与输出）。
- `examples/`: 脱敏示例数据（用于理解字段结构）。

## 🔒 数据与仓库策略
- 仓库默认忽略 `outputs/`、`task_*/outputs/`、以及 `task_customer/data/` 下的业务数据文件，避免把客户数据/证书数据/邮件 .msg 推送到 GitHub。
- `task_customer/data/` 仅保留 PPT 模板与说明文件；真实运行时请把数据文件放到本地并在 `task_customer/config/config.yaml` 配置 paths 指向实际路径。

## 🛠️ 常见问题
- **列名不匹配**：如果导出的 Excel 列名与工具预设的不一致，可以在 `modules/data_analyzer.py` 的 `column_mapping` 中添加映射。
- **PPT 样式**：工具会尝试在模板中寻找表格并填充，建议模板中包含一个带标题和内容占位符的幻灯片。
