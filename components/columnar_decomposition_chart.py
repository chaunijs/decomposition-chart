"""
Columnar Multi-level Decomposition / Marimekko Chart Component
Includes High-Res PNG Export, Batch Export for All Product Categories (ZIP/PNG),
Smart Minimum-Height Guarantee, and Guaranteed High-Legibility Labels across all columns.
Renders contiguous, adjacent columns with no gaps.
"""

from __future__ import annotations
import json
# pyrefly: ignore [missing-import]
import streamlit.components.v1 as components
from typing import List, Dict, Any, Optional


def render_columnar_decomposition(
    columns_data: List[Dict[str, Any]],
    all_categories_data: Optional[Dict[str, List[Dict[str, Any]]]] = None,
    height: int = 800,
    positive_color: str = "#15803d",
    negative_color: str = "#d81e3a",
    chart_title: str = "Decomposition Chart",
    **kwargs: Any
):
    """
    Renders the multi-column decomposition chart with high-res PNG export,
    batch export of all categories, and contiguous adjacent columns.
    """
    if not columns_data:
        return

    data_json = json.dumps(columns_data)
    all_cats_json = json.dumps(all_categories_data or {chart_title: columns_data})

    html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Columnar Decomposition Chart</title>
  <script src="https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
  <style>
    :root {{
      --pos-color: {positive_color};
      --neg-color: {negative_color};
    }}
    * {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    }}
    body {{
      background: transparent;
      color: #0f172a;
      padding: 6px 10px;
      display: flex;
      flex-direction: column;
      align-items: center;
      min-height: 100vh;
      user-select: none;
    }}
    /* Toolbar Action Panel */
    .toolbar-panel {{
      width: 100%;
      max-width: 1260px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      background: #ffffff;
      padding: 10px 18px;
      border-radius: 10px;
      border: 1.5px solid #e2e8f0;
      box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
      margin-bottom: 12px;
    }}
    .toolbar-left {{
      display: flex;
      align-items: center;
      gap: 12px;
      font-size: 14px;
      font-weight: 600;
      color: #334155;
    }}
    .cat-badge {{
      background: #e0f2fe;
      color: #0369a1;
      padding: 4px 14px;
      border-radius: 6px;
      font-size: 14px;
      font-weight: 800;
      letter-spacing: 0.02em;
    }}
    .cat-dropdown {{
      background: #e0f2fe;
      color: #0369a1;
      border: 1.5px solid #7dd3fc;
      padding: 5px 12px;
      border-radius: 6px;
      font-size: 14px;
      font-weight: 800;
      letter-spacing: 0.02em;
      cursor: pointer;
      outline: none;
      transition: all 0.15s ease;
    }}
    .cat-dropdown:hover, .cat-dropdown:focus {{
      border-color: #0284c7;
      background: #bae6fd;
    }}
    .toolbar-right {{
      display: flex;
      align-items: center;
      gap: 10px;
    }}
    .export-btn {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      background: #0284c7;
      color: #ffffff;
      border: none;
      padding: 8px 14px;
      border-radius: 6px;
      font-size: 13px;
      font-weight: 700;
      cursor: pointer;
      box-shadow: 0 2px 4px rgba(2, 132, 199, 0.25);
      transition: all 0.15s ease;
    }}
    .export-btn.btn-all {{
      background: #059669;
      box-shadow: 0 2px 5px rgba(5, 150, 105, 0.3);
    }}
    .export-btn:hover {{
      opacity: 0.92;
      transform: translateY(-1px);
    }}
    .export-btn:disabled {{
      background: #94a3b8;
      cursor: not-allowed;
      transform: none;
    }}
    
    /* Main Chart Card */
    .chart-container {{
      width: 100%;
      max-width: 1260px;
      display: flex;
      flex-direction: column;
      background: #ffffff;
      padding: 16px 16px 20px 16px;
      border-radius: 16px;
      box-shadow: 0 4px 20px rgba(0, 0, 0, 0.06);
      border: 1.5px solid #e2e8f0;
    }}
    /* Contiguous Adjacent Columns (Flush with 0 gap) */
    .columns-wrapper {{
      display: flex;
      width: 100%;
      height: 620px;
      gap: 0px;
      background: #ffffff;
      border: 1px solid #ffffff;
      overflow: hidden;
      border-radius: 4px;
    }}
    .col-stack {{
      flex: 1 1 0px;
      width: 0;
      min-width: 0;
      display: flex;
      flex-direction: column;
      height: 100%;
      gap: 0px;
      position: relative;
    }}
    
    /* Block Styles (Adjacent with Uniform 1px White Dividing Lines) */
    .block {{
      position: relative;
      display: flex;
      flex-direction: column;
      justify-content: center;
      align-items: center;
      color: #ffffff;
      transition: filter 0.15s ease;
      cursor: pointer;
      overflow: hidden;
      box-sizing: border-box;
      border: 1px solid #ffffff;
      padding: 2px 4px;
      min-height: 25px;
      min-width: 0;
      width: 100%;
    }}
    .block:hover {{
      filter: brightness(1.12);
      z-index: 10;
    }}
    .block.positive {{
      background-color: var(--pos-color);
    }}
    .block.negative {{
      background-color: var(--neg-color);
    }}

    /* Global Typography */
    .block-share {{
      font-weight: 800;
      color: #ffffff;
      letter-spacing: -0.01em;
      line-height: 1.1;
      text-shadow: 0 1px 2px rgba(0, 0, 0, 0.45);
    }}
    .block-label {{
      font-weight: 700;
      color: #ffffff;
      text-transform: uppercase;
      letter-spacing: 0.02em;
      line-height: 1.15;
      text-shadow: 0 1px 2px rgba(0, 0, 0, 0.45);
    }}

    /* Top Left Badge (Growth %) - White Pill with Dark Text */
    .badge-growth {{
      position: absolute;
      top: 4px;
      left: 4px;
      background: #ffffff;
      color: #0f172a;
      font-size: 11px;
      font-weight: 800;
      padding: 2px 5.5px;
      border-radius: 4px;
      line-height: 1.15;
      box-shadow: 0 1px 2px rgba(0, 0, 0, 0.25);
      z-index: 2;
    }}

    /* Top Right Badge (bps Impact) - Dark Bold Text */
    .badge-bps {{
      position: absolute;
      top: 5px;
      right: 5px;
      font-size: 11.5px;
      font-weight: 800;
      color: #0f172a;
      line-height: 1.15;
      z-index: 2;
    }}

    /* 1. Standard / Tall Block Layout (Height >= 50px) */
    .block.layout-standard .block-content {{
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      text-align: center;
      pointer-events: none;
    }}
    .block.layout-standard .block-share {{
      font-size: 22px;
    }}
    .block.layout-standard .block-label {{
      font-size: 13.5px;
      margin-top: 2px;
    }}

    /* 2. Compact / Thin Block Layout (Height < 50px) */
    .block.layout-compact {{
      min-height: 25px;
      padding: 1px 3px;
    }}
    .block.layout-compact .badge-growth {{
      top: 2px;
      left: 3px;
      font-size: 9.5px;
      padding: 1px 3.5px;
      border-radius: 3px;
    }}
    .block.layout-compact .badge-bps {{
      top: 2.5px;
      right: 4px;
      font-size: 9.5px;
    }}
    .block.layout-compact .block-content {{
      display: flex;
      flex-direction: row;
      gap: 4px;
      align-items: center;
      justify-content: center;
      text-align: center;
      pointer-events: none;
      width: 100%;
      min-width: 0;
      padding: 0 2px;
      overflow: visible;
    }}
    .block.layout-compact .block-share {{
      font-size: 13.5px;
      font-weight: 800;
      flex-shrink: 0;
    }}
    .block.layout-compact .block-label {{
      font-size: 11.5px;
      font-weight: 700;
      white-space: nowrap;
      overflow: visible;
      text-overflow: clip;
      letter-spacing: 0.01em;
    }}

    /* Column Headers at Bottom (Strict 1:1 Adjacent Alignment) */
    .headers-wrapper {{
      display: flex;
      width: 100%;
      margin-top: 14px;
      gap: 0px;
    }}
    .col-header {{
      flex: 1 1 0px;
      width: 0;
      min-width: 0;
      text-align: center;
      color: #0284c7;
      font-size: 18px;
      font-weight: 800;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      padding: 4px 0;
      letter-spacing: -0.01em;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}

    /* Tooltip */
    #tooltip {{
      position: absolute;
      display: none;
      background: #0f172a;
      color: #ffffff;
      padding: 10px 14px;
      border-radius: 7px;
      font-size: 12.5px;
      line-height: 1.4;
      z-index: 1000;
      pointer-events: none;
      box-shadow: 0 6px 18px rgba(0,0,0,0.35);
      border: 1.5px solid #38bdf8;
    }}
    
    #hidden-render-container {{
      position: fixed;
      left: -9999px;
      top: -9999px;
      width: 1260px;
      visibility: hidden;
    }}
  </style>
</head>
<body>
  <!-- Action Header Bar -->
  <div class="toolbar-panel">
    <div class="toolbar-left">
      <span>Category:</span>
      <select id="cat-selector" class="cat-dropdown" title="Select Product Category"></select>
    </div>
    <div class="toolbar-right">
      <button class="export-btn btn-all" id="btn-export-all-png" title="Export PNG images for all categories (POWDER, LIQUID, etc.) into a ZIP file">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
          <path d="M21 8v13H3V8"></path>
          <path d="M1 3h22v5H1z"></path>
          <path d="M10 12h4"></path>
        </svg>
        📦 Export All (ZIP)
      </button>
      <button class="export-btn" id="btn-export-png" title="Download High Resolution Image of {chart_title}">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
          <polyline points="7 10 12 15 17 10"></polyline>
          <line x1="12" y1="15" x2="12" y2="3"></line>
        </svg>
        📸 Export Current PNG
      </button>
    </div>
  </div>

  <div class="chart-container" id="capture-area">
    <div class="columns-wrapper" id="columns-container"></div>
    <div class="headers-wrapper" id="headers-container"></div>
  </div>

  <div id="hidden-render-container"></div>
  <div id="tooltip"></div>

  <script>
    const data = {data_json};
    const allCategories = {all_cats_json};
    const colContainer = document.getElementById("columns-container");
    const headerContainer = document.getElementById("headers-container");
    const tooltip = document.getElementById("tooltip");

    const totalCanvasHeight = 620;

    function buildChartDOM(cols, targetColContainer, targetHeaderContainer) {{
      targetColContainer.innerHTML = "";
      targetHeaderContainer.innerHTML = "";

      cols.forEach((col) => {{
        const validBlocks = (col.blocks || []).filter(b => b && (b.share > 0));
        if (validBlocks.length === 0) return;

        const colDiv = document.createElement("div");
        colDiv.className = "col-stack";

        // Total share in column
        const totalColShare = validBlocks.reduce((acc, b) => acc + (b.share || 0), 0) || 100;
        const blockCount = validBlocks.length;

        validBlocks.forEach((b) => {{
          const blockDiv = document.createElement("div");
          const isPos = b.is_positive !== undefined ? b.is_positive : (b.bps >= 0);
          blockDiv.className = `block ${{isPos ? "positive" : "negative"}}`;

          // Proportional flex distribution with guaranteed minimum block height
          const shareWeight = Math.max(b.share || 0.1, (100 / blockCount) * 0.20);
          blockDiv.style.flex = `${{shareWeight}} 1 0%`;

          const approxHeightPx = (totalCanvasHeight * (b.share / totalColShare));

          if (approxHeightPx >= 50) {{
            blockDiv.classList.add("layout-standard");
          }} else {{
            blockDiv.classList.add("layout-compact");
          }}

          // Badge Growth
          const growthSign = b.growth > 0 ? "+" : "";
          const growthBadge = document.createElement("div");
          growthBadge.className = "badge-growth";
          growthBadge.innerText = `${{growthSign}}${{b.growth}}%`;
          blockDiv.appendChild(growthBadge);

          // Center Content
          const contentDiv = document.createElement("div");
          contentDiv.className = "block-content";

          const shareSpan = document.createElement("span");
          shareSpan.className = "block-share";
          shareSpan.innerText = `${{b.share}}%`;

          const activeLabel = b.actual_label || b.label;

          const labelSpan = document.createElement("span");
          labelSpan.className = "block-label";
          labelSpan.innerText = activeLabel;

          contentDiv.appendChild(shareSpan);
          contentDiv.appendChild(labelSpan);
          blockDiv.appendChild(contentDiv);

          // Badge BPS
          const bpsBadge = document.createElement("div");
          bpsBadge.className = "badge-bps";
          bpsBadge.innerText = `${{b.bps}} bps`;
          blockDiv.appendChild(bpsBadge);

          // Tooltip Interaction
          blockDiv.addEventListener("mousemove", (e) => {{
            tooltip.style.display = "block";
            tooltip.style.left = (e.pageX + 14) + "px";
            tooltip.style.top = (e.pageY - 24) + "px";
            tooltip.innerHTML = `
              <div style="font-weight:700; color:#38bdf8; font-size:13.5px; margin-bottom:4px;">
                ${{col.column_name}} &rsaquo; ${{activeLabel}}
              </div>
              <table style="width:100%; border-collapse:collapse; font-size:12px;">
                <tr><td style="color:#94a3b8; padding-right:12px;">Share:</td><td><strong>${{b.share}}%</strong></td></tr>
                <tr><td style="color:#94a3b8; padding-right:12px;">Growth:</td><td><span style="color:${{b.growth >= 0 ? '#4ade80' : '#f87171'}}"><strong>${{growthSign}}${{b.growth}}%</strong></span></td></tr>
                <tr><td style="color:#94a3b8; padding-right:12px;">BPS Impact:</td><td><span style="color:${{b.bps >= 0 ? '#4ade80' : '#f87171'}}"><strong>${{b.bps}} bps</strong></span></td></tr>
              </table>
            `;
          }});

          blockDiv.addEventListener("mouseleave", () => {{
            tooltip.style.display = "none";
          }});

          colDiv.appendChild(blockDiv);
        }});

        targetColContainer.appendChild(colDiv);

        // Header label
        const headerDiv = document.createElement("div");
        headerDiv.className = "col-header";
        headerDiv.innerText = col.column_name;
        targetHeaderContainer.appendChild(headerDiv);
      }});
    }}

    let currentCategory = "{chart_title}";

    // Initialize interactive category dropdown in toolbar
    const catSelector = document.getElementById("cat-selector");
    if (catSelector) {{
      catSelector.innerHTML = "";
      const catList = Object.keys(allCategories);
      catList.forEach((cat) => {{
        const opt = document.createElement("option");
        opt.value = cat;
        opt.innerText = cat;
        if (cat.toUpperCase() === currentCategory.toUpperCase()) {{
          opt.selected = true;
        }}
        catSelector.appendChild(opt);
      }});

      catSelector.addEventListener("change", (e) => {{
        const chosen = e.target.value;
        if (allCategories[chosen]) {{
          currentCategory = chosen;
          buildChartDOM(allCategories[chosen], colContainer, headerContainer);
        }}
      }});
    }}

    // Initial render
    buildChartDOM(allCategories[currentCategory] || data, colContainer, headerContainer);

    function getFileTimestamp() {{
      const d = new Date();
      const pad = function(n) {{ return (n < 10 ? "0" : "") + n; }};
      return "" + d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + "_" + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds());
    }}

    // 1. Export Current High-Res PNG (3x Scale)
    document.getElementById("btn-export-png").addEventListener("click", () => {{
      const exportBtn = document.getElementById("btn-export-png");
      exportBtn.innerText = "⏳ Generating...";
      exportBtn.disabled = true;

      const captureTarget = document.getElementById("capture-area");

      html2canvas(captureTarget, {{
        scale: 3,
        useCORS: true,
        backgroundColor: "#ffffff",
        logging: false
      }}).then((canvas) => {{
        const link = document.createElement("a");
        const ts = getFileTimestamp();
        link.download = currentCategory.replace(/[^a-zA-Z0-9_-]/g, '_') + "_Decomposition_" + ts + ".png";
        link.href = canvas.toDataURL("image/png");
        link.click();

        exportBtn.innerHTML = `
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
            <polyline points="7 10 12 15 17 10"></polyline>
            <line x1="12" y1="15" x2="12" y2="3"></line>
          </svg>
          📸 Export Current PNG
        `;
        exportBtn.disabled = false;
      }}).catch((err) => {{
        console.error("Export error:", err);
        alert("Could not export image: " + err);
        exportBtn.innerText = "📸 Export Current PNG";
        exportBtn.disabled = false;
      }});
    }});

    // 2. Export All Categories PNGs in ZIP
    document.getElementById("btn-export-all-png").addEventListener("click", async () => {{
      const allBtn = document.getElementById("btn-export-all-png");
      allBtn.innerText = "⏳ Exporting All...";
      allBtn.disabled = true;

      try {{
        const zip = new JSZip();
        const hiddenContainer = document.getElementById("hidden-render-container");
        hiddenContainer.style.visibility = "visible";

        const catKeys = Object.keys(allCategories);
        const batchTs = getFileTimestamp();

        for (let i = 0; i < catKeys.length; i++) {{
          const catName = catKeys[i];
          allBtn.innerText = `⏳ Exporting (${{i+1}}/${{catKeys.length}}): ${{catName}}...`;

          const catCols = allCategories[catName];

          const tempChart = document.createElement("div");
          tempChart.className = "chart-container";
          tempChart.style.width = "1260px";

          const tempCols = document.createElement("div");
          tempCols.className = "columns-wrapper";
          const tempHeaders = document.createElement("div");
          tempHeaders.className = "headers-wrapper";

          tempChart.appendChild(tempCols);
          tempChart.appendChild(tempHeaders);
          hiddenContainer.innerHTML = "";
          hiddenContainer.appendChild(tempChart);

          buildChartDOM(catCols, tempCols, tempHeaders);

          const canvas = await html2canvas(tempChart, {{
            scale: 3,
            useCORS: true,
            backgroundColor: "#ffffff",
            logging: false
          }});

          const imgData = canvas.toDataURL("image/png").replace(/^data:image\\/png;base64,/, "");
          zip.file(catName.replace(/[^a-zA-Z0-9_-]/g, '_') + "_Decomposition_" + batchTs + ".png", imgData, {{ base64: true }});
        }}

        hiddenContainer.innerHTML = "";
        hiddenContainer.style.visibility = "hidden";

        const content = await zip.generateAsync({{ type: "blob" }});
        const zipLink = document.createElement("a");
        zipLink.href = URL.createObjectURL(content);
        zipLink.download = "All_Categories_Decomposition_" + batchTs + ".zip";
        zipLink.click();

        allBtn.innerHTML = `
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
            <path d="M21 8v13H3V8"></path>
            <path d="M1 3h22v5H1z"></path>
            <path d="M10 12h4"></path>
          </svg>
          📦 Export All (ZIP)
        `;
        allBtn.disabled = false;
      }} catch (err) {{
        console.error("Batch export error:", err);
        alert("Batch export failed: " + err);
        allBtn.innerText = "📦 Export All (ZIP)";
        allBtn.disabled = false;
      }}
    }});
  </script>
</body>
</html>
"""
    components.html(html_content, height=height, scrolling=False)
