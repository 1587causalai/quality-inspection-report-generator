# app.py
import gradio as gr
# Removed imports: pandas, openpyxl, datetime, timedelta, os, tempfile

# Import the processing function from main.py
from main import process_files

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
    if file1_obj is None or file2_obj is None:
        raise gr.Error("Please upload both required files.")

    try:
        # Gradio provides temporary paths for uploaded files
        file1_path = file1_obj.name
        file2_path = file2_obj.name

        # Define the output filename (can be customized if needed)
        output_filename = "generated_report.xlsx"

        # Call the core processing logic from main.py
        print(f"Processing files: {file1_path}, {file2_path}") # Log input paths
        result_path = process_files(file1_path, file2_path, output_filename)
        print(f"process_files returned: {result_path}") # Log result path

        if result_path:
            # Return the path of the generated file for Gradio to serve
            return result_path
        else:
            # If process_files returned None, it means an error occurred
            raise gr.Error("Failed to generate the report. Check logs or input files.")

    except Exception as e:
        # Catch any other unexpected errors during the wrapper execution
        import traceback
        print(f"Error in Gradio wrapper (generate_report): {e}")
        print(traceback.format_exc())
        raise gr.Error(f"An unexpected error occurred: {e}")


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