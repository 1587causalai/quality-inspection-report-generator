# app.py
import gradio as gr
import logging # Import logging
import sys # Import sys to output to stderr
# Removed imports: pandas, openpyxl, datetime, timedelta, os, tempfile

# Import the processing function from main.py
from main import process_files

# --- Configure Logging --- (Configure before functions)
log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
logging.basicConfig(level=logging.INFO, # Use INFO for general info, DEBUG for more detail
                    format=log_format,
                    handlers=[logging.StreamHandler(sys.stderr)]) # Output to stderr

logger = logging.getLogger(__name__) # Get logger for this module

def generate_report(file1_obj, file2_obj):
    """
    Gradio wrapper function to handle file uploads and call the main processing logic.

    Args:
        file1_obj: Gradio File object for the first input file.
        file2_obj: Gradio File object for the second input file.

    Returns:
        str: Path to the generated output Excel file if successful.
    Raises:
        gr.Error: If file processing fails.
    """
    logger.info("generate_report function called.")
    if file1_obj is None or file2_obj is None:
        logger.error("Missing one or both input files.")
        raise gr.Error("Please upload both required files.")

    try:
        # Gradio provides temporary paths for uploaded files
        file1_path = file1_obj.name
        file2_path = file2_obj.name
        logger.info(f"Input file 1 path: {file1_path}")
        logger.info(f"Input file 2 path: {file2_path}")

        # Define the output filename
        output_filename = "generated_report.xlsx"
        logger.info(f"Target output base filename: {output_filename}")

        # Call the core processing logic from main.py
        logger.info(f"Calling process_files...")
        result_path = process_files(file1_path, file2_path, output_filename)
        logger.info(f"process_files returned: {result_path}")

        if result_path:
            logger.info(f"Report generation successful. Returning path: {result_path}")
            # Return the path of the generated file for Gradio to serve
            return result_path
        else:
            # If process_files returned None, it means an error occurred
            logger.error("process_files returned None, indicating failure.")
            raise gr.Error("Failed to generate the report. Check application logs for details.")

    except Exception as e:
        # Catch any other unexpected errors during the wrapper execution
        logger.exception("An unexpected error occurred in the Gradio wrapper (generate_report)") # Logs exception info automatically
        raise gr.Error(f"An unexpected error occurred. Check application logs for details.") # User-friendly message


# Create Gradio Interface
inputs = [
    gr.File(label="上传数据源文件 (类似 1.xls)"),
    gr.File(label="上传模板文件 (类似 2.xlsx)")
]
outputs = gr.File(label="下载生成的报告")

title = "Quality Inspection Report Generator" # Use English title
description = "Upload the data source file and template file to generate the combined quality inspection report." # Use English description

# Ensure interface runs on the default port 7860
# share=False is default and recommended for custom deployments
# server_name="0.0.0.0" makes it accessible within the container/network
demo = gr.Interface(
    fn=generate_report,
    inputs=inputs,
    outputs=outputs,
    title=title,
    description=description,
    allow_flagging='never' # Disable flagging
)

if __name__ == "__main__":
    # Launch the Gradio app
    demo.launch(server_name="0.0.0.0") # Port defaults to 7860 