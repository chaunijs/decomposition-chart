# 📊 Decomposition Automated

> **Automated Columnar Metric Decomposition & Marimekko Hierarchy Visualization Studio**

An interactive analytics web application built with **Streamlit**, **D3.js / HTML5 Canvas**, and **Pandas** to analyze and visualize multi-level business breakdowns, variance drivers, market shares, growth percentages, and basis point (BPS) contributions.

---

## 🌟 Key Features

- **📊 Columnar Multi-Level Decomposition (Marimekko Style)**:
  - Parallel vertical stacked columns representing dimension hierarchies (e.g. `Total TH` $\rightarrow$ `Channel` $\rightarrow$ `Price` $\rightarrow$ `Region` $\rightarrow$ `Packsize`).
  - Block heights strictly proportional to percentage volume/value share (`%`).
- **🎯 Growth & BPS Contribution Badges**:
  - **Top-Left Badge**: Growth rate percentage (e.g. `+8%`, `-15%`, `-32%`).
  - **Top-Right Badge**: Basis points variance impact (e.g. `177 bps`, `-159 bps`, `36 bps`).
  - **Color-Coded Performance**: Green for positive growth/contribution ($\ge 0$), Red for negative growth/contraction ($< 0$).
- **🔍 Smart Adaptive Readability Engine**:
  - Dynamically calculates pixel heights and adapts layout for small/compact blocks (e.g. `SH Online 1%`, `Tier2 4%`, `XXL 4%`) using inline single-row alignment, completely preventing overlapping text.
- **📦 High-Resolution & Batch Image Export**:
  - **`📸 Export Current PNG`**: High-DPI (3x scale) crystal-clear export of the active category view.
  - **`📦 Export All Categories (ZIP)`**: Batch renders all product categories (`POWDER`, `LIQUID`, etc.) at 3x resolution and downloads them bundled into a `.zip` archive.
- **📁 Native Excel Named Ranges & File Upload**:
  - Automatically discovers and parses named ranges in `dataset-decomp.xlsx` (`contribution_powder`, `cont_liquid`, `bps_*`, `growth_pct_*`).
  - Supports custom CSV, Excel (`.xlsx`, `.xls`), and JSON uploads.

---

## 📂 Project Structure

```text
decomposition-automated/
├── app.py                             # Main Streamlit dashboard application
├── dataset-decomp.xlsx                # Reference Excel dataset with defined Named Ranges
├── requirements.txt                   # Project dependencies (streamlit, pandas, openpyxl, plotly)
├── components/
│   └── columnar_decomposition_chart.py # Columnar Marimekko D3/Canvas renderer & export engine
├── utils/
│   ├── columnar_processor.py          # Excel named range extractor & column aggregation engine
│   └── data_processor.py              # Generic dataset parser & hierarchical tree builder
└── sample_data/
    └── sample_retail_sales.csv        # Fallback sample dataset
```

---

## 🚀 Quickstart Guide

### Prerequisites
- Python 3.10+

### 1. Setup Virtual Environment (`.venv`)
```powershell
# Navigate to the repository
cd decomposition-chart

# Create virtual environment
python -m venv .venv

# Install dependencies into .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### 2. Run the Application
You can start the application using the PowerShell script:
```powershell
powershell -ExecutionPolicy Bypass -File .\run_dashboard.ps1
```
Or directly with Python:
```powershell
.\.venv\Scripts\python.exe -m streamlit run app.py
```

Open your web browser at: **`http://localhost:8501`**

---

## 📑 Excel Named Range Specification (`dataset-decomp.xlsx`)

The system automatically extracts data using the following standard Excel defined names:

| Range Name | Dimensions Covered | Description |
| :--- | :--- | :--- |
| `contribution_powder` / `cont_powder` | `Total TH`, `Channel`, `Price`, `Region`, `Packsize` | Share proportions (decimal `0.0-1.0` or percentage `0-100%`) for Powder |
| `bps_powder` | `Total`, `Channel`, `Price`, `Region`, `Packsize` | Basis point impacts (e.g. `172 bps`, `-164 bps`) for Powder |
| `growth_pct_powder` | `Total`, `Channel`, `Price`, `Region`, `Packsize` | Growth percentages (e.g. `0.14` or `14%`) for Powder |
| `cont_liquid` | `Total TH`, `Channel`, `Price`, `Region`, `Packsize` | Share proportions for Liquid category |
| `bps_liquid` | `Total`, `Channel`, `Price`, `Region`, `Packsize` | Basis point impacts for Liquid category |
| `growth_pct_liquid` | `Total`, `Channel`, `Price`, `Region`, `Packsize` | Growth percentages for Liquid category |
| `cont_fab_en` | `Total TH`, `Channel`, `Price`, `Region`, `Packsize` | Share proportions for Fabric Enhancer category |
| `cont_dish` | `Total TH`, `Channel`, `Price`, `Region`, `Packsize` | Share proportions for Dishwashing category |

### 🏷️ Dynamic Label Table
All block labels (e.g., `SH CHAIN`, `SH Online`, `Tier1`, `GBKK`, `S`, etc.) are read **dynamically from the label table** in `dataset-decomp.xlsx` (located in Columns `A:E` adjacent to the data tables, or via named ranges like `label_<product>`). No labels are hardcoded, so changing any name in Excel will automatically update the charts.

---

## 🎨 Visual Customization

From the sidebar, you can configure:
- **Product Category Selector**: Instant switching between `POWDER`, `LIQUID`, or custom products.
- **Positive Impact Color**: Default `#1b8129` (Green).
- **Negative Impact Color**: Default `#d81e3a` (Red).
- **Underlying Data Viewer**: Tabular grid preview and CSV download of parsed records.

---

## 📄 License
Internal Development / Proprietary
