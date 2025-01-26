from openpyxl import load_workbook
import pandas as pd

def print_sheet_content(file_path, sheet_name):
    wb = load_workbook(file_path)
    sheet = wb[sheet_name]
    
    print(f"\nContent of sheet '{sheet_name}':")
    for i, row in enumerate(sheet.rows, 1):
        values = [str(cell.value) if cell.value is not None else '' for cell in row]
        print(f"Row {i}: {values}")

# 打印模板文件的内容
print_sheet_content('2.xlsx', 'WACKER')

# 读取第一个文件的Sand工作表
print("\n原始数据结构：")
file1 = '1.xls'
df = pd.read_excel(file1, sheet_name='Sand')
print("\n列名：")
print(df.columns.tolist())

# 打印前几行数据以查看结构
print("\n前几行数据：")
print(df.head())

# 读取跳过12行的数据
print("\n跳过12行的数据：")
sand_df_skipped = pd.read_excel(file1, sheet_name='Sand', skiprows=12)
print(sand_df_skipped.head())

# 打印所有列名
print("\n跳过12行后的列名：")
print(sand_df_skipped.columns.tolist()) 