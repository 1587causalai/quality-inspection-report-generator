import pandas as pd

def extract_sand_data(input_file, output_file):
    try:
        # 读取Excel文件的Sand工作表
        xls = pd.ExcelFile(input_file)
        if 'Sand' not in xls.sheet_names:
            raise ValueError("Excel文件中缺少'Sand'工作表")

        # 读取Sand工作表
        df = pd.read_excel(xls, sheet_name='Sand', header=None)

        # 定义数据起始行和列索引
        start_row = 13  # 数据从第11行开始（0-based索引）
        crucible_id_col = 2  # Crucible ID在C列（0-based索引）
        element_start_col = 4  # 元素数据从第5列开始（0-based索引）
        element_end_col = 14  # 元素数据到第15列结束（0-based索引）

        # 提取数据
        results = []
        for row_idx in range(start_row, len(df)):
            row = df.iloc[row_idx]
            if pd.isna(row[crucible_id_col]):  # 跳过空行
                continue

            # 提取Crucible ID
            crucible_id = str(row[crucible_id_col]).strip()

            # 提取元素数据
            elements = {
                "Crucible ID": crucible_id,
                "Al": row[element_start_col],
                "Ca": row[element_start_col + 1],
                "Cu": row[element_start_col + 2],
                "Fe": row[element_start_col + 3],
                "K": row[element_start_col + 4],
                "Li": row[element_start_col + 5],
                "Mg": row[element_start_col + 6],
                "Mn": row[element_start_col + 7],
                "Na": row[element_start_col + 8],
                "Ti": row[element_start_col + 9],
                "Zr": row[element_start_col + 10],
            }
            results.append(elements)

        # 将结果转换为DataFrame
        result_df = pd.DataFrame(results)

        # 保存到新的Excel文件
        result_df.to_excel(output_file, index=False)
        return f"数据已成功保存到 {output_file}"

    except Exception as e:
        return f"处理错误: {str(e)}"

# 使用示例
if __name__ == "__main__":
    input_file = "1.xls"  # 输入文件路径
    output_file = "output.xlsx"  # 输出文件路径
    result = extract_sand_data(input_file, output_file)
    print(result)