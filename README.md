# 质量检验报告生成工具

一个用于从 Excel 数据生成质量检验报告的 Python 工具。

## 系统要求

- Python 3.6+
- pandas
- openpyxl

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

1. 准备输入文件：
   - `1.xls`：源数据文件，包含 HEADER、Dimension 和 Sand 工作表
   - `2.xlsx`：模板文件，包含 WACKER 工作表

2. 运行脚本：
```bash
python main.py
```

3. 检查输出文件 `2_updated.xlsx`

## 许可证

MIT License 