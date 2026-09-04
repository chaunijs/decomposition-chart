"""
Columnar Decomposition & Marimekko Chart Studio
-----------------------------------------------
Loads and visualizes named range datasets from dataset-decomp.xlsx (Powder & Liquid)
as well as custom CSV/Excel uploads.
Matched pixel-for-pixel with reference visualization.
"""

from __future__ import annotations
import json
import os
from datetime import datetime
from typing import Dict, Any, List
# pyrefly: ignore [missing-import]
import streamlit as st
import polars as pl
import pandas as pd
from utils.columnar_processor import extract_excel_named_ranges, compute_columnar_decomposition
from components.columnar_decomposition_chart import render_columnar_decomposition
from utils.data_processor import load_dataset, read_excel, detect_columns
from utils.history_manager import load_history, add_file_to_history, clear_history

st.set_page_config(
    page_title="Columnar Decomposition Chart",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
  .main-title {
    font-size: 2.1rem;
    font-weight: 800;
    color: #0284c7;
    margin-bottom: 0.2rem;
  }
  .sub-title {
    color: #64748b;
    font-size: 0.95rem;
    margin-bottom: 1.2rem;
  }
</style>
""", unsafe_allow_html=True)


@st.cache_data(show_spinner="⚡ Processing Excel workbook and loading all categories into memory...")
def get_cached_named_range_data_from_bytes(file_bytes: bytes, filename: str) -> Dict[str, Dict[str, Any]]:
    return extract_excel_named_ranges(file_bytes)


@st.cache_data(show_spinner="⚡ Loading Excel workbook named ranges...")
def get_cached_named_range_data_from_path(file_path: str, mtime: float) -> Dict[str, Dict[str, Any]]:
    return extract_excel_named_ranges(file_path)


@st.cache_data(show_spinner="⚡ Loading dataset...")
def get_cached_custom_dataframe_from_path(file_path: str, mtime: float) -> pl.DataFrame:
    return load_dataset(file_path)


@st.cache_data(show_spinner="⚡ Loading dataset...")
def get_cached_custom_dataframe(file_bytes: bytes, filename: str) -> pl.DataFrame:
    import io
    return load_dataset(io.BytesIO(file_bytes))


def main():
    workspace_dir = os.path.dirname(os.path.abspath(__file__))
    default_dataset = os.path.join(workspace_dir, "dataset-decomp.xlsx")

    # Sidebar: Data Source & History
    st.sidebar.header("📂 Data Source")

    # 1. File Uploader
    uploaded_file = st.sidebar.file_uploader(
        "Upload Excel or CSV",
        type=["xlsx", "xls", "csv"],
        key="file_uploader"
    )

    # 2. History of Recent Workbooks
    history = load_history()
    active_path = None
    active_filename = None

    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        entry = add_file_to_history(uploaded_file.name, file_bytes)
        active_path = entry["path"]
        active_filename = entry["filename"]
        history = load_history()
    elif history:
        history_options = ["-- Select from recent files --"] + [f"📄 {h['display_name']} · {h['size_str']}" for h in history]
        selected_choice = st.sidebar.selectbox(
            "🕒 Recent Workbooks",
            range(len(history_options)),
            format_func=lambda i: history_options[i],
            index=0,
            help="Select any previously uploaded workbook"
        )
        if selected_choice > 0:
            hist_item = history[selected_choice - 1]
            active_path = hist_item["path"]
            active_filename = hist_item["filename"]

        if st.sidebar.button("🗑️ Clear History", use_container_width=True):
            clear_history()
            st.rerun()

    # 3. Auto-load default dataset-decomp.xlsx if no file explicitly chosen
    if not active_path and os.path.exists(default_dataset):
        active_path = default_dataset
        active_filename = "dataset-decomp.xlsx"

    columns_data = []
    display_df = pl.DataFrame()
    selected_prod_label = ""
    all_categories_dict = {}

    if active_path and os.path.exists(active_path):
        mtime = os.path.getmtime(active_path)
        is_excel = active_filename.lower().endswith((".xlsx", ".xls"))

        if is_excel:
            named_range_data = get_cached_named_range_data_from_path(active_path, mtime)
            if named_range_data:
                available_products = list(named_range_data.keys())
                # Default to LIQUID to immediately match the reference screenshot
                default_idx = available_products.index("LIQUID") if "LIQUID" in available_products else 0
                selected_prod = st.sidebar.selectbox(
                    "Select Product Category",
                    available_products,
                    index=default_idx
                )
                selected_prod_label = selected_prod
                columns_data = named_range_data[selected_prod]["columns_data"]
                display_df = named_range_data[selected_prod]["df"]
                all_categories_dict = {k: v["columns_data"] for k, v in named_range_data.items()}
                st.sidebar.success(f"Loaded: `{selected_prod}` from `{active_filename}`")
            else:
                df = get_cached_custom_dataframe_from_path(active_path, mtime)
                display_df = df
                columns_data = compute_columnar_decomposition(df=df, dimensions=[], metric_col="")
                all_categories_dict = {"Custom_Data": columns_data}
                st.sidebar.success(f"Loaded: `{active_filename}`")
        else:
            df = get_cached_custom_dataframe_from_path(active_path, mtime)
            display_df = df
            columns_data = compute_columnar_decomposition(df=df, dimensions=[], metric_col="")
            all_categories_dict = {"Custom_Data": columns_data}
            st.sidebar.success(f"Loaded: `{active_filename}`")
    else:
        st.sidebar.info("Please upload an Excel/CSV file or select a workbook from history.")

    st.sidebar.markdown("---")
    st.sidebar.markdown("## 🎨 Visual Colors")
    col_pos = st.sidebar.color_picker("Positive Color (Growth / BPS >= 0)", "#15803d")
    col_neg = st.sidebar.color_picker("Negative Color (Growth / BPS < 0)", "#d81e3a")

    # Main UI Header
    st.markdown('<div class="main-title">📊 Columnar Decomposition Chart</div>', unsafe_allow_html=True)
    if selected_prod_label:
        st.markdown(f'<div class="sub-title">Hierarchical contribution & variance analysis for <strong>{selected_prod_label}</strong></div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="sub-title">Multi-level hierarchical breakdown with proportional shares, growth % badges, and bps impacts.</div>', unsafe_allow_html=True)

    tab_columnar, tab_data = st.tabs([
        "📊 Columnar Decomposition Chart",
        "📋 Data & Range Summary"
    ])

    with tab_columnar:
        if columns_data:
            render_columnar_decomposition(
                columns_data=columns_data,
                all_categories_data=all_categories_dict or {selected_prod_label or "Decomposition": columns_data},
                height=800,
                positive_color=col_pos,
                negative_color=col_neg,
                chart_title=selected_prod_label or "Decomposition_Chart"
            )
        else:
            st.info("👋 **Welcome!** Please upload an Excel file (`.xlsx`) or CSV file in the sidebar to generate the Columnar Decomposition Chart.")

    with tab_data:
        st.subheader("Underlying Extracted Data")
        has_data = False
        if isinstance(display_df, pl.DataFrame):
            has_data = not display_df.is_empty()
        elif isinstance(display_df, pd.DataFrame):
            has_data = not display_df.empty

        if has_data:
            st.dataframe(display_df, use_container_width=True)
            if hasattr(display_df, "write_csv"):
                csv_bytes = display_df.write_csv().encode("utf-8")
            else:
                csv_bytes = display_df.to_csv(index=False).encode("utf-8")
            now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_prod = (selected_prod_label or "data").replace(" ", "_")
            st.download_button(
                "📥 Download Table as CSV",
                data=csv_bytes,
                file_name=f"decomp_{safe_prod}_{now_str}.csv",
                mime="text/csv"
            )
        else:
            st.info("No table data available.")


if __name__ == "__main__":
    main()
