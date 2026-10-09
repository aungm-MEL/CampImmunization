from __future__ import annotations

import contextlib
import importlib.util
import io
import traceback
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory

import streamlit as st


APP_DIR = Path(__file__).resolve().parent
PIPELINE_PATH = APP_DIR / "PRF EPI report.py"
DEFAULT_OUTPUT_NAME = f"Umpium_EPI_Quarterly_report_{date.today():%Y%m%d}.xlsx"


def load_pipeline_module(script_path: Path):
    if not script_path.exists():
        raise FileNotFoundError(f"Pipeline script not found: {script_path}")

    spec = importlib.util.spec_from_file_location("prf_epi_pipeline", str(script_path))
    if spec is None or spec.loader is None:
        raise ImportError(f"Could not import pipeline from: {script_path}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run_pipeline(input_bytes: bytes, source_name: str) -> tuple[bytes, str]:
    module = load_pipeline_module(PIPELINE_PATH)

    with TemporaryDirectory(prefix="prf_streamlit_") as temp_dir:
        temp_dir_path = Path(temp_dir)
        input_path = temp_dir_path / source_name
        output_path = temp_dir_path / DEFAULT_OUTPUT_NAME

        input_path.write_bytes(input_bytes)

        log_buffer = io.StringIO()
        with contextlib.redirect_stdout(log_buffer):
            module.combine_sheets(input_path, output_path)

        if not output_path.exists():
            raise FileNotFoundError("Expected output workbook was not created.")

        return output_path.read_bytes(), log_buffer.getvalue()


def parse_pipeline_messages(logs: str) -> tuple[list[str], list[str]]:
    warnings: list[str] = []
    others: list[str] = []

    for raw_line in logs.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("WARNING:"):
            warnings.append(line)
        else:
            others.append(line)

    return warnings, others


st.set_page_config(page_title="PRF EPI Report", layout="wide")
st.title("PRF EPI Quarterly Report Builder")
st.caption("Upload a PRF source workbook, run the existing pipeline, and download the output report.")

if not PIPELINE_PATH.exists():
    st.error(f"Required file not found: {PIPELINE_PATH.name}")
    st.stop()

upload = st.file_uploader("Upload source workbook", type=["xlsx", "xlsm"])
default_output_name = DEFAULT_OUTPUT_NAME
output_name = st.text_input("Output filename", value=default_output_name).strip() or default_output_name

run_clicked = st.button("Run report", type="primary", use_container_width=True)

if upload is None:
    st.info("Please upload an Excel workbook to continue.")
    st.stop()

if not run_clicked:
    st.success(f"Ready to process: {upload.name}")
    st.stop()

if not output_name.lower().endswith(".xlsx"):
    output_name = f"{output_name}.xlsx"

progress = st.progress(15, text="Loading pipeline...")

try:
    progress.progress(45, text="Processing workbook...")
    output_bytes, logs = run_pipeline(upload.getvalue(), upload.name)
    progress.progress(100, text="Done")

    st.success("Report created successfully.")

    warnings, others = parse_pipeline_messages(logs)

    st.subheader("Warning Signs")
    if warnings:
        for warning_line in warnings:
            st.warning(warning_line)
    else:
        st.info("No warning signs were found during processing.")

    if logs.strip():
        st.subheader("Pipeline logs")
        if others:
            st.caption("Additional log lines")
        st.code(logs, language="text")

    st.download_button(
        label="Download report",
        data=output_bytes,
        file_name=output_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
except Exception as exc:
    progress.empty()
    st.error(f"Pipeline failed: {exc}")
    st.code(traceback.format_exc(), language="text")
