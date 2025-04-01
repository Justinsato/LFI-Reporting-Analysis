import logging
import pandas as pd
from io import BytesIO
from variance_notes_processor import main as run_variance_pipeline

def main(req):
    try:
        uploaded_file = req.files["file"]
        temp_input_path = "/tmp/input_variance_file.xlsx"
        with open(temp_input_path, "wb") as f:
            f.write(uploaded_file.read())

        # Override paths in the main script
        import variance_notes_processor as vnp
        vnp.INPUT_FILE = temp_input_path
        vnp.OUTPUT_FILE = "/tmp/output_with_notes.xlsx"
        vnp.OUTPUT_SUMMARY_FILE = "/tmp/output_summary.xlsx"
        vnp.SUPPORTING_DATA_DIR = "/tmp/support_data"  # assumes this is mounted or provided

        run_variance_pipeline()

        with open(vnp.OUTPUT_SUMMARY_FILE, "rb") as f:
            result = f.read()

        return {
            "status": 200,
            "headers": {"Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
            "body": result
        }

    except Exception as e:
        logging.exception("Error during variance processing")
        return {
            "status": 500,
            "body": str(e)
        }

