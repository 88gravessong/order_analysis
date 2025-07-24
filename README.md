# 订单分析系统

一个基于Flask的Web应用，用于分析订单数据并生成各项指标报告。

## 功能特点

- 📊 **多文件上传**：支持同时上传多个Excel或CSV订单文件
- 📅 **日期筛选**：可按日期范围筛选订单数据
- 📈 **指标计算**：自动计算签收率、完成率、退款率等关键指标
- 🗺️ **省份分析**：统计各省份的SKU签收率和订单占比
- 📄 **结果导出**：生成Excel格式的分析报告
- 🎨 **友好界面**：简洁美观的Web界面

## 快速开始

### 方法一：使用启动脚本（推荐）

#### Windows用户
双击运行 `start.bat` 文件，或在命令行中执行：
```cmd
start.bat
```

#### Linux/Mac用户
在终端中执行：
```bash
./start.sh
```

启动脚本会自动：
1. 检查并激活虚拟环境
2. 安装所需依赖（如果缺失）
3. 启动Flask服务

### 方法二：手动启动

1. **创建虚拟环境**
```bash
python -m venv .venv
```

2. **激活虚拟环境**
```bash
# Windows
.venv\Scripts\activate

# Linux/Mac
source .venv/bin/activate
```

3. **安装依赖**
```bash
pip install -r requirements.txt
```

4. **启动应用**
```bash
python app.py
```

## 使用方法

1. 访问 `http://localhost:5000`
2. 上传订单Excel或CSV文件（支持多文件）
3. 选择日期范围（可选）
4. 点击"分析订单"或"省份分析"按钮
5. 查看分析结果并下载Excel报告

## 支持的文件格式

可上传 `.xlsx` 或 `.csv` 文件，系统会自动识别以下列名（不区分大小写）：
- `Order Substatus` - 订单子状态
- `Cancelation/Return Type` 或 `Cancellation/Return Type` - 取消/退货类型
- `Seller SKU` - 商品SKU
- `Shipped Time` - 发货时间
- `Created Time` - 创建时间

## 计算指标

- **订单数**：该SKU的总订单数
- **签收率**：已完成率 + 已送达率 + 退款率
- **已完成率**：状态为"已完成"(`Completed`)且无取消类型的订单比例
- **已送达率**：状态为"已送达"(`Delivered`)的订单比例
- **退款率**：状态包含"Return"或"Refund"的订单比例
- **发货前取消率**：状态为"已取消"(`Canceled`)且发货时间为空的订单比例
- **发货后取消率**：状态为"已取消"(`Canceled`)且有发货时间的订单比例
- **仍在途率**：状态为"运输中"(`In transit`)的订单比例

## 项目结构

```
order_analysis/
├── app.py                    # Flask主应用
├── compute_logic.py          # 核心计算逻辑
├── compute_province_metrics.py # 按省份统计SKU指标
├── requirements.txt          # Python依赖
├── start.bat                 # Windows启动脚本
├── start.sh                  # Linux/Mac启动脚本
├── .gitignore               # Git忽略文件
└── templates/               # HTML模板
    ├── index.html           # 首页
    └── results.html         # 结果页面
```

## 技术栈

- **后端**：Flask 3.0.0
- **Excel处理**：openpyxl 3.1.2
- **前端**：HTML + CSS + JavaScript（原生）
- **样式**：Bootstrap 5

## 注意事项

- 确保上传的文件包含必需的列名
- 大文件处理可能需要较长时间
- 建议在处理大量数据时关闭其他应用以节省内存

## 许可证

MIT License 