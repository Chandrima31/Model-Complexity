
import streamlit as st
import pandas as pd
from io import BytesIO
from trace_model import trace_excel_formulas

st.set_page_config(page_title="Excel Formula Tracer", layout="wide")

st.title("🔍 Excel Formula Dependency Tracer")
st.write("Upload an Excel file to trace recursive and cross-sheet formula dependencies.")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    st.success("File uploaded successfully!")
    with st.spinner("Tracing formula dependencies..."):
        formula_traces = trace_excel_formulas(uploaded_file)

    st.write("### Formula Tracing Results")
    if formula_traces:
        for cell, trace_lines in formula_traces.items():
            st.markdown(f"**{cell}**")
            st.text("\n".join(trace_lines))
    else:
        st.warning("No formulas found in the specified range (default X9:Z17).")
