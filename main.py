import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta
import os # Added for path manipulation
import tempfile # Import tempfile
import logging # Import logging

# Get logger for this module
logger = logging.getLogger(__name__)

def process_files(file1_path, file2_path, output_filename="generated_report.xlsx"):
    """
    Processes two input Excel files and generates a combined report.
    Saves the output to a temporary directory to avoid potential permission issues.

    Args:
        file1_path (str): Path to the first input Excel file (data source).
        file2_path (str): Path to the second input Excel file (template).
        output_filename (str): Desired *base* name for the output report file (e.g., 'report.xlsx').
                               The actual path will be in a temp directory.

    Returns:
        str: The full path to the generated output Excel file in a temporary directory.
             Returns None if an error occurs during processing.
    """
    logger.info(f"Starting report generation with input files: {file1_path}, {file2_path}")
    try:
        # 读取第一个文件
        logger.info(f"Reading header from {file1_path}")
        header_df = pd.read_excel(file1_path, sheet_name='HEADER')
        logger.info(f"Reading dimension data from {file1_path}")
        dimension_df = pd.read_excel(file1_path, sheet_name='Dimension', skiprows=12)
        dimension_df.columns = dimension_df.iloc[0]
        dimension_df = dimension_df.iloc[1:].reset_index(drop=True)
        logger.info(f"Reading sand data from {file1_path}")
        sand_df = pd.read_excel(file1_path, sheet_name='Sand', header=None)
        logger.info(f"Finished reading data from {file1_path}")

        # 读取第二个文件
        logger.info(f"Loading template workbook from {file2_path}")
        wb = load_workbook(file2_path)
        
        # Check if 'WACKER' sheet exists
        if 'WACKER' not in wb.sheetnames:
            logger.error(f"Template file '{file2_path}' must contain a sheet named 'WACKER'.")
            return None # Indicate error
        wacker_sheet = wb['WACKER']
        logger.info("Template workbook loaded successfully.")

        # 获取Sales Order Quantity和Quality Assured By
        sales_order_quantity = header_df.iloc[5, 2]
        quality_assured_by = header_df.iloc[3, 7]
        logger.info(f"Retrieved Sales Order Qty: {sales_order_quantity}, Quality Assured By: {quality_assured_by}")

        # 定义元素和行号的对应关系 (Copied from original script)
        element_row_mapping = {
            'Al': 9, 'Ca': 10, 'Cu': 11, 'Fe': 12, 'K': 13, 'Li': 14,
            'Mg': 15, 'Mn': 16, 'Na': 17, 'Ti': 18, 'Zr': 19
        }
        element_col_mapping = {
            'Al': 4, 'Ca': 5, 'Cu': 6, 'Fe': 7, 'K': 8, 'Li': 9,
            'Mg': 10, 'Mn': 11, 'Na': 12, 'Ti': 13, 'Zr': 14
        }

        logger.info(f"Processing {len(dimension_df)} entries from dimension data.")
        # 遍历Dimension表格中的每个Customer ID
        for index, row in dimension_df.iterrows():
            customer_id = row['Customer ID']
            # Ensure customer_id is a valid sheet name (Excel has limitations)
            safe_customer_id = str(customer_id).replace('/', '-').replace('\\', '-').replace('?', '').replace('*', '').replace('[', '').replace(']', '')
            safe_customer_id = safe_customer_id[:31] # Max sheet name length

            # Handle potential NaN or empty Customer ID
            if pd.isna(customer_id) or not str(customer_id).strip():
                logger.warning(f"Skipping row {index+14} due to missing or invalid Customer ID.")
                continue

            logger.debug(f"Processing Customer ID: {customer_id} (Index: {index})") # Use debug for per-item processing

            inspection_date_str = ""
            inspection_date = None # Initialize inspection_date
            try:
                 # Check if inspection_date is already datetime or needs conversion
                if isinstance(row['Inspection Date'], datetime):
                    inspection_date = row['Inspection Date']
                else:
                    inspection_date = pd.to_datetime(row['Inspection Date'])
                inspection_date_str = inspection_date.strftime('%Y-%m-%d')
            except Exception as e:
                logger.warning(f"Could not parse Inspection Date for Customer ID {customer_id}: {e}. Skipping date fields.")
                # inspection_date remains None

            new_sheet_title = safe_customer_id
            # Avoid duplicate sheet names if safe_customer_id becomes the same for different original IDs
            sheet_count = 1
            while new_sheet_title in wb.sheetnames:
                 suffix = f"_{sheet_count}"
                 max_len = 31 - len(suffix)
                 new_sheet_title = safe_customer_id[:max_len] + suffix
                 sheet_count += 1

            new_sheet = wb.create_sheet(title=new_sheet_title)

            # 复制WACKER表格的内容到新工作表
            for row_wacker in wacker_sheet.iter_rows(values_only=True):
                new_sheet.append(row_wacker)

            # 填充数据
            new_sheet['B3'] = str(sales_order_quantity) + ' PCS'
            new_sheet['B4'] = customer_id # Use original ID here
            if inspection_date: # Only fill dates if parsing was successful
                 new_sheet['D4'] = inspection_date_str
                 new_sheet['B5'] = inspection_date_str
                 new_sheet['D5'] = (inspection_date + timedelta(days=730)).strftime('%Y-%m-%d')

            # 从sand表中获取当前customer_id的数据
            sand_rows = sand_df[sand_df[2] == customer_id] # 使用第3列（索引2）作为Crucible ID
            if not sand_rows.empty:
                sand_row = sand_rows.iloc[0]
                # 填充元素数据 (with added error handling)
                for element, target_row in element_row_mapping.items():
                    try:
                        source_col = element_col_mapping[element]
                        # Check if value exists and handle potential errors
                        value = sand_row.get(source_col) # Use .get for safety
                        if value is not None and not pd.isna(value):
                             new_sheet[f'D{target_row}'] = value
                        else:
                             logger.warning(f"Missing or invalid sand data for {element}, Customer ID {customer_id}, Col Index {source_col}")
                             # Optionally fill with a default value or leave blank
                             # new_sheet[f'D{target_row}'] = "N/A"
                    except KeyError:
                         logger.warning(f"Column index {source_col} not found in sand_row for {element}, Customer ID {customer_id}")
                    except Exception as e:
                         logger.error(f"Error filling element {element} for Customer ID {customer_id}: {e}")

            # 填充Analysis result/分析结果 (with added error handling)
            dim_mapping = {
                20: 'OD1', 21: 'OD2', 22: 'OD3', 23: 'Height',
                24: 'Wall11', 25: 'Wall12', 26: 'Wall13',
                27: 'Wall2', 28: 'Wall3'
            }
            for target_row, source_col_name in dim_mapping.items():
                 try:
                     # Check if value exists and handle potential errors
                     value = row.get(source_col_name) # Use .get for safety
                     if value is not None and not pd.isna(value):
                          new_sheet[f'D{target_row}'] = value
                     else:
                         logger.warning(f"Missing or invalid dimension data for {source_col_name}, Customer ID {customer_id}")
                         # Optionally fill with a default value or leave blank
                         # new_sheet[f'D{target_row}'] = "N/A"
                 except KeyError:
                     logger.warning(f"Column '{source_col_name}' not found in dimension_df for Customer ID {customer_id}")
                 except Exception as e:
                     logger.error(f"Error filling dimension {source_col_name} for Customer ID {customer_id}: {e}")

            # 保持"批准人："文本，并在其后添加名字
            new_sheet['D29'] = f"批准人：{quality_assured_by}"
            logger.debug(f"Finished processing data for Customer ID: {customer_id}") # Use debug

        # Remove the original template sheet if it exists and wasn't intended to be kept
        if 'WACKER' in wb.sheetnames:
             logger.info("Removing 'WACKER' template sheet from the output workbook.")
             del wb['WACKER'] # Remove template if no longer needed

        # Create a temporary file path for the output
        try:
            temp_dir = tempfile.gettempdir()
            # Ensure the base output filename is used, not a potentially problematic one from input args
            safe_output_filename = os.path.basename(output_filename if output_filename else "generated_report.xlsx")
            # Create a unique temporary file path
            # Using mkstemp gives more control, but let's try simple join first
            # _, temp_output_path = tempfile.mkstemp(suffix=".xlsx", prefix="report_", dir=temp_dir)
            # Let's use a predictable name within the temp dir, might be easier for Gradio/platform
            temp_output_path = os.path.join(temp_dir, safe_output_filename)
            
            logger.info(f"Attempting to save report to temporary path: {temp_output_path}")
            wb.save(temp_output_path)
            logger.info(f"Successfully saved report to: {temp_output_path}")
            return temp_output_path # Return the full path to the temporary file
        except Exception as save_error:
            logger.exception(f"Error saving workbook to temporary path {temp_output_path}") # Log exception
            return None

    except FileNotFoundError:
        logger.error(f"Error: Input file not found. Check paths: {file1_path}, {file2_path}")
        return None
    except KeyError as e:
        logger.exception(f"Error: Missing expected column or sheet name: {e}. Check input file formats.")
        return None
    except Exception as e:
        # Log other unexpected errors
        logger.exception(f"An unexpected error occurred in process_files: {e}")
        return None


# Keep the original script behavior if run directly (optional)
if __name__ == "__main__":
    # Configure basic logging for direct script execution if needed
    # Note: app.py usually handles the main config when run via Gradio
    if not logging.getLogger().hasHandlers(): # Only configure if not already configured by app.py import
        log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        logging.basicConfig(level=logging.INFO, format=log_format)

    # Define default input/output files for direct execution
    default_file1 = '1.xls'
    default_file2 = '2.xlsx'
    default_output = '2_updated.xlsx'

    print(f"Running script directly. Processing {default_file1} and {default_file2}...")
    # Use logger here too if desired, but print might be fine for direct runs
    logger.info(f"Running script directly. Processing {default_file1} and {default_file2}...")
    output_path = process_files(default_file1, default_file2, default_output)

    if output_path:
        print(f"Report generated successfully: {output_path}")
        logger.info(f"Direct run: Report generated successfully: {output_path}")
    else:
        print("Report generation failed.")
        logger.error("Direct run: Report generation failed.")