
import streamlit as st
import pandas as pd
from trace_model_updated import trace_excel_formulas_full

st.set_page_config(page_title="Excel Formula Dependency Tracer", layout="wide")

st.title("🔍 Excel Formula Dependency Tracer")
st.write("Upload an Excel file to trace recursive and cross-sheet formula dependencies.")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    st.success("File uploaded successfully!")
    with st.spinner("Tracing formula dependencies..."):
        formula_traces, hop_summary = trace_excel_formulas_full(uploaded_file)

    st.write("### 📊 Hop Count Summary")
    hop_df = pd.DataFrame(list(hop_summary.items()), columns=["# of Hops", "No. of Formulae"])
    st.dataframe(hop_df.sort_values(by="# of Hops", ascending=True), use_container_width=True)

    st.write("### 🔗 Formula Tracing Results")
    if formula_traces:
        for cell, trace_lines in formula_traces.items():
            st.markdown(f"**{cell}**")
            st.text("\n".join(trace_lines))
    else:
        st.warning("No formulas found in the workbook.")
