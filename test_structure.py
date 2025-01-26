import pandas as pd

def analyze_excel_structure(file_path):
    """分析Excel文件的基本结构"""
    print(f"\n开始分析文件: {file_path}")
    
    # 读取Excel文件
    xls = pd.ExcelFile(file_path)
    
    # 打印所有sheet名称
    print("\n所有Sheet页:")
    for sheet_name in xls.sheet_names:
        print(f"- {sheet_name}")
    
    # 分析每个sheet的结构
    for sheet_name in xls.sheet_names:
        print(f"\n\n分析 Sheet [{sheet_name}]:")
        
        # 尝试不同的skiprows值来找到正确的数据结构
        for skip_rows in [0, 6, 12]:
            print(f"\n跳过 {skip_rows} 行后的结构:")
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=skip_rows)
                print("\n列名:")
                print(df.columns.tolist())
                print("\n前3行数据:")
                print(df.head(3))
                
                # 打印非空行数
                non_empty_rows = len(df.dropna(how='all'))
                print(f"\n非空行数: {non_empty_rows}")
                
            except Exception as e:
                print(f"读取出错: {str(e)}")

if __name__ == "__main__":
    INPUT_FILE = "1.xls"
    analyze_excel_structure(INPUT_FILE) 