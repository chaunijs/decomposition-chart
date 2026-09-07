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

    # Do not auto-load on start: only process when user uploads a file or selects from history
    columns_data = []
    display_df = pl.DataFrame()
    selected_prod_label = ""
    all_categories_dict = {}

    if active_path and os.path.exists(active_path):
        mtime = os.path.getmtime(active_path)
        is_excel = active_filename.lower().endswith((".xlsx", ".xls"))

        if is_excel:
            named_range_data = get_cached_named_range_data_from_path(active_path, mtime)
            if not named_range_data:
                # Bypass cache in case an earlier empty result was stored
                named_range_data = extract_excel_named_ranges(active_path)
            if named_range_data:
                available_products = list(named_range_data.keys())
                
                # Dynamic category selection strictly based on workbook named ranges
                saved_category = st.session_state.get("selected_product_category")
                default_idx = available_products.index(saved_category) if saved_category in available_products else 0

                selected_prod = st.sidebar.selectbox(
                    "Select Product Category",
                    available_products,
                    index=default_idx,
                    key="product_category_selector"
                )
                st.session_state["selected_product_category"] = selected_prod

                selected_prod_label = selected_prod
                columns_data = named_range_data[selected_prod]["columns_data"]
                display_df = named_range_data[selected_prod]["df"]
                all_categories_dict = {k: v["columns_data"] for k, v in named_range_data.items()}
                st.sidebar.success(f"Loaded: **{selected_prod}** ({len(available_products)} categories in `{active_filename}`)")
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

        st.sidebar.markdown("---")
        st.sidebar.markdown("## 🎨 Visual Colors")
        col_pos = st.sidebar.color_picker("Positive Color (Growth / BPS >= 0)", "#15803d")
        col_neg = st.sidebar.color_picker("Negative Color (Growth / BPS < 0)", "#d81e3a")
    else:
        col_pos = "#15803d"
        col_neg = "#d81e3a"

    # Main UI Header with Download Template on the top right
    col_hdr_left, col_hdr_right = st.columns([0.72, 0.28], vertical_alignment="center")
    with col_hdr_left:
        st.markdown('<div class="main-title">📊 Columnar Decomposition Chart</div>', unsafe_allow_html=True)
        if selected_prod_label:
            st.markdown(f'<div class="sub-title">Hierarchical contribution & variance analysis for <strong>{selected_prod_label}</strong></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="sub-title">Multi-level hierarchical breakdown with proportional shares, growth % badges, and bps impacts.</div>', unsafe_allow_html=True)

    with col_hdr_right:
        template_path = os.path.join(workspace_dir, "sample_data", "Decomposition_Chart_Template.xlsx")
        if not os.path.exists(template_path):
            template_path = default_dataset

        if os.path.exists(template_path):
            with open(template_path, "rb") as f_tpl:
                tpl_bytes = f_tpl.read()
            st.download_button(
                label="📥 Download Excel Template",
                data=tpl_bytes,
                file_name="Decomposition_Chart_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Download a ready-to-use Excel template configured with all named tables (contribution, bps, growth, labels).",
                use_container_width=True
            )

    # Render Columnar Decomposition Chart directly
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
        st.markdown("""
        <div style="background-color: #1e293b; border-radius: 12px; padding: 2.2rem; margin-top: 1.5rem; border: 1px solid #334155; text-align: center;">
            <div style="font-size: 2.5rem; margin-bottom: 0.8rem;">📊</div>
            <h3 style="color: #f8fafc; margin-bottom: 0.6rem; font-weight: 700;">No File Loaded Yet</h3>
            <p style="color: #94a3b8; font-size: 1rem; max-width: 580px; margin: 0 auto 1.5rem auto; line-height: 1.6;">
                To generate the <strong>Columnar Decomposition & Marimekko Chart</strong>, please <strong>upload an Excel workbook</strong> (.xlsx) or select one from <strong>Recent Workbooks</strong> in the sidebar.
            </p>
            <div style="display: inline-flex; gap: 1rem; color: #cbd5e1; font-size: 0.9rem; background: #0f172a; padding: 0.8rem 1.4rem; border-radius: 8px; border: 1px solid #1e293b;">
                <span>📁 <strong>Step 1:</strong> Select or upload file</span>
                <span>⚡ <strong>Step 2:</strong> Automatic processing</span>
                <span>📈 <strong>Step 3:</strong> Interactive chart visualizes</span>
            </div>
        </div>
        """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
