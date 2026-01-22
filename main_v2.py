#!/usr/bin/env python3
"""
质量检验报告生成器 v2 - 交互式版本
用法：python3 main_v2.py
根据提示输入两个Excel文件路径即可自动处理
"""

import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta
import os
import sys


def find_sheet_by_keyword(sheet_names, keyword):
    """查找包含关键词的工作表名称（忽略空格）"""
    for name in sheet_names:
        if keyword.lower() in name.strip().lower():
            return name
    return None


def process_files(file1, file2):
    """处理两个Excel文件并生成报告"""
    
    print(f"\n📂 正在读取文件...")
    print(f"   文件1: {file1}")
    print(f"   文件2: {file2}")
    
    # 读取第一个文件的工作表名称
    xls = pd.ExcelFile(file1)
    sheet_names = xls.sheet_names
    print(f"\n📋 文件1包含的工作表: {sheet_names}")
    
    # 自动查找工作表
    header_sheet = find_sheet_by_keyword(sheet_names, 'HEADER')
    dimension_sheet = find_sheet_by_keyword(sheet_names, 'Dimension')
    sand_sheet = find_sheet_by_keyword(sheet_names, 'Sand')
    
    if not all([header_sheet, dimension_sheet, sand_sheet]):
        print("❌ 错误：无法找到所需的工作表 (HEADER, Dimension, Sand)")
        print(f"   找到: HEADER={header_sheet}, Dimension={dimension_sheet}, Sand={sand_sheet}")
        return False
    
    print(f"   ✓ 找到工作表: HEADER='{header_sheet}', Dimension='{dimension_sheet}', Sand='{sand_sheet}'")
    
    # 读取数据
    header_df = pd.read_excel(file1, sheet_name=header_sheet)
    
    # 读取Dimension表，跳过前12行，然后使用第13行作为列名
    dimension_df = pd.read_excel(file1, sheet_name=dimension_sheet, skiprows=12)
    dimension_df.columns = dimension_df.iloc[0]
    dimension_df = dimension_df.iloc[1:].reset_index(drop=True)
    
    # 读取Sand表
    sand_df = pd.read_excel(file1, sheet_name=sand_sheet, header=None)
    
    # 读取第二个文件
    wb = load_workbook(file2)
    print(f"\n📋 文件2包含的工作表: {wb.sheetnames}")
    
    wacker_sheet_name = find_sheet_by_keyword(wb.sheetnames, 'WACKER')
    if not wacker_sheet_name:
        print("❌ 错误：无法找到WACKER模板工作表")
        return False
    
    print(f"   ✓ 找到模板工作表: '{wacker_sheet_name}'")
    wacker_sheet = wb[wacker_sheet_name]
    
    # 获取Sales Order Quantity和Quality Assured By
    sales_order_quantity = header_df.iloc[5, 2]
    quality_assured_by = header_df.iloc[3, 7]
    
    print(f"\n📊 基础信息:")
    print(f"   订单数量: {sales_order_quantity}")
    print(f"   批准人: {quality_assured_by}")
    print(f"   待处理记录数: {len(dimension_df)}")
    
    # 定义元素和行号的对应关系
    element_row_mapping = {
        'Al': 9, 'Ca': 10, 'Cu': 11, 'Fe': 12, 'K': 13,
        'Li': 14, 'Mg': 15, 'Mn': 16, 'Na': 17, 'Ti': 18, 'Zr': 19
    }
    
    # 定义元素在Sand表中的列索引
    element_col_mapping = {
        'Al': 4, 'Ca': 5, 'Cu': 6, 'Fe': 7, 'K': 8,
        'Li': 9, 'Mg': 10, 'Mn': 11, 'Na': 12, 'Ti': 13, 'Zr': 14
    }
    
    print(f"\n⏳ 正在生成检验报告...")
    
    # 遍历Dimension表格中的每个Customer ID
    for index, row in dimension_df.iterrows():
        customer_id = row['Customer ID']
        inspection_date = pd.to_datetime(row['Inspection Date']).strftime('%Y-%m-%d')
        
        # 创建新的工作表
        new_sheet = wb.create_sheet(title=str(customer_id))
        
        # 复制WACKER表格的内容到新工作表
        for row_wacker in wacker_sheet.iter_rows(values_only=True):
            new_sheet.append(row_wacker)
        
        # 填充数据
        new_sheet['B3'] = str(sales_order_quantity) + ' PCS'
        new_sheet['B4'] = customer_id
        new_sheet['D4'] = inspection_date
        new_sheet['B5'] = inspection_date
        new_sheet['D5'] = (datetime.strptime(inspection_date, '%Y-%m-%d') + timedelta(days=730)).strftime('%Y-%m-%d')
        
        # 从sand表中获取当前customer_id的数据
        sand_rows = sand_df[sand_df[2] == customer_id]
        if not sand_rows.empty:
            sand_row = sand_rows.iloc[0]
            for element, target_row in element_row_mapping.items():
                source_col = element_col_mapping[element]
                new_sheet[f'D{target_row}'] = sand_row[source_col]
        
        # 填充尺寸数据
        new_sheet['D20'] = row['OD1']
        new_sheet['D21'] = row['OD2']
        new_sheet['D22'] = row['OD3']
        new_sheet['D23'] = row['Height']
        new_sheet['D24'] = row['Wall11']
        new_sheet['D25'] = row['Wall12']
        new_sheet['D26'] = row['Wall13']
        new_sheet['D27'] = row['Wall2']
        new_sheet['D28'] = row['Wall3']
        
        # 批准人
        new_sheet['D29'] = f"批准人：{quality_assured_by}"
        
        # 显示进度
        if (index + 1) % 20 == 0 or index == len(dimension_df) - 1:
            print(f"   已处理: {index + 1}/{len(dimension_df)}")
    
    # 生成输出文件名
    base_name = os.path.splitext(os.path.basename(file2))[0]
    output_file = f"{base_name}_updated.xlsx"
    
    # 保存文件
    wb.save(output_file)
    
    print(f"\n✅ 处理完成!")
    print(f"   生成工作表数: {len(wb.sheetnames)}")
    print(f"   输出文件: {output_file}")
    
    return True


def main():
    print("=" * 60)
    print("       质量检验报告生成器 v2 - 交互式版本")
    print("=" * 60)
    
    # 列出当前目录的Excel文件
    current_dir = os.getcwd()
    excel_files = [f for f in os.listdir(current_dir) 
                   if f.endswith(('.xls', '.xlsx')) and not f.startswith('~')]
    
    if excel_files:
        print("\n📁 当前目录下的Excel文件:")
        for i, f in enumerate(excel_files, 1):
            print(f"   {i}. {f}")
    
    print("\n" + "-" * 60)
    
    # 获取文件1
    print("\n请输入第一个文件 (包含HEADER/Dimension/Sand数据):")
    file1 = input(">>> ").strip()
    
    if not file1:
        print("❌ 未输入文件名，退出。")
        sys.exit(1)
    
    if not os.path.exists(file1):
        print(f"❌ 文件不存在: {file1}")
        sys.exit(1)
    
    # 获取文件2
    print("\n请输入第二个文件 (包含WACKER模板):")
    file2 = input(">>> ").strip()
    
    if not file2:
        print("❌ 未输入文件名，退出。")
        sys.exit(1)
    
    if not os.path.exists(file2):
        print(f"❌ 文件不存在: {file2}")
        sys.exit(1)
    
    # 处理文件
    success = process_files(file1, file2)
    
    if not success:
        sys.exit(1)


if __name__ == "__main__":
    main()
