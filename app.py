# app.py
import gradio as gr
import logging # Import logging
import sys # Import sys
# Removed imports: pandas, openpyxl, datetime, timedelta, os, tempfile

# Import the processing function from main.py
from main import process_files

# Configure logging (consistent with main.py)
# Note: If main.py is imported, its basicConfig might already run.
# Re-configuring here ensures app.py specific logs also follow the format.
# Using getLogger and adding handler might be more robust in complex scenarios,
# but basicConfig is often sufficient for simpler apps.
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
                    stream=sys.stderr)

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
    logging.info("generate_report function called by Gradio.")
    if file1_obj is None or file2_obj is None:
        logging.error("One or both files were not provided.")
        raise gr.Error("Please upload both required files.")

    try:
        # Gradio provides temporary paths for uploaded files
        file1_path = file1_obj.name
        file2_path = file2_obj.name
        logging.info(f"Input file paths received: file1='{file1_path}', file2='{file2_path}'")

        # Define the output filename (can be customized if needed)
        output_filename = "generated_report.xlsx"
        logging.info(f"Target output base filename: {output_filename}")

        # Call the core processing logic from main.py
        logging.info(f"Calling main.process_files...")
        result_path = process_files(file1_path, file2_path, output_filename)
        logging.info(f"main.process_files returned: '{result_path}'")

        if result_path:
            # Return the path of the generated file for Gradio to serve
            logging.info(f"Report generation successful. Returning path: {result_path}")
            return result_path
        else:
            # If process_files returned None, it means an error occurred (already logged in main.py)
            logging.error("main.process_files indicated failure (returned None).")
            raise gr.Error("Failed to generate the report. Check application logs for details.")

    except Exception as e:
        # Catch any other unexpected errors during the wrapper execution
        # Log the exception with traceback before raising Gradio error
        logging.exception(f"An unexpected error occurred in the Gradio wrapper (generate_report): {e}")
        raise gr.Error(f"An unexpected error occurred. Check application logs.")


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
    # Setup logging for direct script execution as well
    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
                        stream=sys.stderr)
    logging.info("Starting Gradio application...")
    # Launch the Gradio app
    demo.launch(server_name="0.0.0.0") # Port defaults to 7860 