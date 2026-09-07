# 📊 Columnar Decomposition Chart Studio

> **Automated Columnar Metric Decomposition & Marimekko Hierarchy Visualization Studio**

An interactive analytics web application built with **Streamlit**, **D3.js / HTML5 Canvas**, and **Polars/Pandas** to analyze and visualize multi-level business breakdowns, variance drivers, market shares, growth percentages, and basis point (BPS) contributions.

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
  - **`📦 Export All Categories (ZIP)`**: Batch renders all product categories (`POWDER`, `LIQUID`, `FAB ENH`, `DW`) at 3x resolution and downloads them bundled into a `.zip` archive.
- **📁 Dual Excel Support (Tables & Named Ranges)**:
  - Supports both **modern Excel Tables** (`Ctrl+T` / ListObjects) and classical **Defined Names** in Name Manager.
  - Automatically matches and extracts `contribution_*`, `bps_*`, `growth_*`, and `label_*` across categories.
- **📥 Pre-Configured Excel Template**:
  - Built-in **"Download Excel Template"** button in the top-right header providing a pre-configured workbook with all 16 tables ready to use.
- **✨ Clean Landing Experience**:
  - Starts in a clean waiting state — visuals and controls only trigger once a file is uploaded or chosen from Recent Workbooks.

---

## 📂 Project Structure

```text
decomposition-chart/
├── app.py                             # Main Streamlit dashboard application
├── app.yaml                           # Databricks Apps runtime configuration (port 8000)
├── databricks.yml                     # Databricks Asset Bundle (DABs) deployment definition
├── dataset-decomp.xlsx                # Reference Excel dataset with defined Named Ranges
├── requirements.txt                   # Project dependencies (streamlit, polars, fastexcel, openpyxl, plotly)
├── run_dashboard.ps1                  # PowerShell launcher for Windows
├── components/
│   └── columnar_decomposition_chart.py # Columnar Marimekko D3/Canvas renderer & export engine
├── utils/
│   ├── columnar_processor.py          # Dual table/range extractor & column aggregation engine
│   ├── data_processor.py              # Generic dataset parser
│   └── history_manager.py             # Recent workbooks tracking & caching
└── sample_data/
    └── Decomposition_Chart_Template.xlsx # Ready-to-fill 16-table Excel template
```

---

## 🚀 Local Quickstart Guide

### Prerequisites
- Python 3.10+

### 1. Setup Virtual Environment (`.venv`)
```powershell
# Navigate to the repository
cd decomposition-chart

# Create virtual environment
python -m venv .venv

# Install dependencies
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### 2. Run Locally
Start the application using the PowerShell script:
```powershell
powershell -ExecutionPolicy Bypass -File .\run_dashboard.ps1
```
Or directly with Python:
```powershell
.\.venv\Scripts\python.exe -m streamlit run app.py
```

Open your browser at: **`http://localhost:8501`**

---

## ☁️ Deployment Option 1: Azure Databricks Apps (Serverless)

This app is pre-configured for **Databricks Apps** with single sign-on (SSO) and serverless hosting.

### 1. Prerequisites
- Azure Databricks workspace (e.g. `https://<workspace-id>.azuredatabricks.net`)
- Databricks CLI v0.250.0+ (`winget install Databricks.DatabricksCLI`)

### 2. Configure `databricks.yml`
Ensure your workspace URL is set in [`databricks.yml`](databricks.yml):
```yaml
targets:
  dev:
    mode: development
    workspace:
      host: https://<your-workspace-id>.azuredatabricks.net
```

### 3. Deploy & Run
```powershell
# Authenticate (one-time)
databricks auth login --host https://<your-workspace-id>.azuredatabricks.net

# Validate bundle configuration
databricks bundle validate -t dev

# Deploy code bundle
databricks bundle deploy -t dev

# Start/run the app
databricks bundle run decomposition_chart -t dev
```

### 4. Cost Management on Databricks
- **Compute Size**: `MEDIUM` (2 vCPUs, 6 GB RAM) = **0.5 DBU / hour** (~$0.20–$0.27/hr).
- **Idle Behavior**: Databricks Apps currently run 24/7 once started unless manually stopped.
- **Stop / Start Commands**:
  ```powershell
  # Stop when not in use ($0 / 0 DBU)
  databricks apps stop decomposition-chart

  # Start when needed
  databricks apps start decomposition-chart
  ```
- **Scheduled Auto On/Off**: You can set up a Databricks Workflow Job calling `w.apps.start("decomposition-chart")` at 08:30 AM and `w.apps.stop("decomposition-chart")` at 06:30 PM (Mon–Fri) to save ~75% on compute costs.

---

## 🌐 Deployment Option 2: Azure App Service (For Non-Databricks Users)

If users across your company need access without having a Databricks account:

1. **Create Web App**: Azure Portal → Create **App Service** (Linux, Python 3.11, Plan: Basic B1 ~$13/mo).
2. **Configure Settings**:
   - **WebSockets**: Turn **ON** (Required by Streamlit).
   - **Startup Command**:
     ```bash
     python -m streamlit run app.py --server.port 8000 --server.address 0.0.0.0
     ```
   - **Application Settings**:
     - `WEBSITES_PORT`: `8000`
     - `SCM_DO_BUILD_DURING_DEPLOYMENT`: `true`
3. **Deploy**: Connect GitHub repository under **Deployment Center** for auto-deploy on push.
4. **Authentication**: Under **Authentication**, enable **Microsoft Entra ID** to restrict access to company email (`@unilever.com`) without requiring Databricks licenses.

---

## 📑 Excel Structure & Naming Conventions

The engine automatically extracts data matching these standard naming prefixes:

| Prefix | Example | Description |
| :--- | :--- | :--- |
| `contribution_*` / `cont_*` | `contribution_powder`, `cont_dw` | Share proportions (decimal `0.0-1.0` or percentage `0-100%`) |
| `bps_*` | `bps_powder`, `bps_dw` | Basis point impacts (e.g. `172 bps`, `-164 bps`) |
| `growth_*` / `growth_pct_*` | `growth_powder`, `growth_dw` | Growth percentages (e.g. `0.14` or `14%`) |
| `label_*` | `label_powder`, `label_dish` | Dimension labels for each column |

Both **Excel Tables** (created via `Ctrl+T` and named in *Table Design*) and **Named Ranges** (defined in *Name Manager*) are supported.

---

## 📄 License
Internal Development / Proprietary
