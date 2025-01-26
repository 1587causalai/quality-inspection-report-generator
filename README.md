# Quality Inspection Report Generator

一个用于从 Excel 数据生成质量检验报告的 Python 工具。

## 功能特点

- 从源数据 Excel 文件自动生成质量检验报告
- 支持多个客户的批量处理
- 保持模板格式和样式
- 自动计算日期和有效期
- 完整的化学元素和尺寸数据处理

## 系统要求

- Python 3.6+
- pandas
- openpyxl

## 安装

1. 克隆仓库：
```bash
git clone https://github.com/1587causalai/quality-inspection-report-generator.git
cd quality-inspection-report-generator
```

2. 安装依赖：
```bash
pip install -r requirements.txt
```

## 使用方法

1. 准备输入文件：
   - 源数据文件（参考 `examples/example_source_data.xls`）：
     * HEADER 工作表：包含基本信息
     * Dimension 工作表：包含尺寸数据
     * Sand 工作表：包含化学元素测试数据
   - 模板文件（参考 `examples/example_template.xlsx`）：
     * 包含预设的报告格式

2. 运行脚本：
```bash
python generate_quality_inspection_reports.py
```

3. 检查输出：
   - 生成的报告将保存为 `2_updated.xlsx`
   - 每个客户的数据将保存在单独的工作表中

## 示例文件

在 `examples` 目录中提供了示例文件：

1. `example_source_data.xls`：源数据文件示例
   - 展示了正确的数据格式和结构
   - 包含了所有必需的工作表和字段

2. `example_template.xlsx`：模板文件示例
   - 展示了标准的报告格式
   - 包含了所有必需的字段和公式

您可以参考这些示例文件来准备您自己的输入文件。

## 数据格式要求

### 输入文件结构

1. 源数据文件需要包含：
   - HEADER：基本信息（数量、批准人等）
   - Dimension：尺寸数据和检验日期
   - Sand：化学元素测试数据

2. 模板文件需要包含：
   - WACKER：报告模板格式

### 输出报告格式

- 基本信息（B3-D5）：数量、批号、日期等
- 化学元素数据（D9-D19）：11种元素的测试结果
- 尺寸数据（D20-D28）：外径、高度、壁厚等
- 批准信息（D29）：批准人姓名

## 文档

详细的代码文档请参考 `quality_inspection_documentation.py`。

## 许可证

MIT License

## 作者

[Your Name]

## 贡献

欢迎提交 Issue 和 Pull Request！ 