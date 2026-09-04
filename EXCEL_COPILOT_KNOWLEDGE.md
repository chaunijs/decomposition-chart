# Excel Copilot Knowledge: Columnar Decomposition Chart (Pure Named Ranges)

This document specifies the exact ground-truth rules and Python in Excel (`=PY`) scripts for generating Columnar Marimekko Decomposition charts from Excel Defined Named Ranges.

---

## 🏷️ Defined Named Range Conventions

Each product category in the Excel workbook is mapped to 4 defined named ranges:

| Target Component | Defined Name in Excel | Refers To (Example) |
| :--- | :--- | :--- |
| **Share Proportions** | `cont_<category>` (e.g. `cont_powder`) | `Sheet1!$G$2:$K$7` (with headers) |
| **Variance / BPS** | `bps_<category>` (e.g. `bps_powder`) | `Sheet1!$M$2:$Q$7` (with headers) |
| **Growth Rates** | `growth_pct_<category>` (e.g. `growth_pct_powder`) | `Sheet1!$S$2:$W$7` (with headers) |
| **Dimension Labels** | `label_<category>` (e.g. `label_powder`) | `Sheet1!$A$3:$E$7` (with headers) |

Known product categories in Name Manager: `powder`, `liquid`, `fab_en`, `dish`.

---

## 📐 Core Chart Generation Rules

1. **RULE 1 (No Renormalization)**:
   - `cont_*` values are absolute proportions of the total market (e.g., Powder SH CHAIN = 11.5%).
   - NEVER divide raw shares by column sums (`height = raw_share / total_share * 100` is WRONG and inflates shares).
   - Scale from the largest column total: `max_total = max(sum of shares in each dimension)`.
   - `ax.set_ylim(0, max_total)`.
   
2. **RULE 2 (Parse BPS Strictly)**:
   - Extract integer from strings (e.g. `"-164 bps"` -> `-164`).
   - Blank or missing BPS must NOT be coerced to `0` (which would falsely turn green). Blank BPS is `None`, colored `#9ca3af` (neutral grey), displaying `"n/a"`.

3. **RULE 3 (Zero / Blank Shares)**:
   - Zero or missing shares produce no block (`if not share or share <= 0: continue`).

4. **RULE 4 (Flush Columns & Sizing)**:
   - `COL_W = 1.0` so columns sit flush without gaps.
   - Text threshold: labels and badges render on segments where `height / max_total >= 0.045`.

---

## 🐍 Corrected Python in Excel Script (`=PY`)

```python
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import re

category = "powder"          # powder | liquid | fab_en | dish

shares_df = xl("cont_" + category, headers=True)
bps_df    = xl("bps_" + category, headers=True)
growth_df = xl("growth_pct_" + category, headers=True)
labels_df = xl("label_" + category, headers=True)

columns = list(shares_df.columns)
POS, NEG, NEUTRAL = "#1b8129", "#d81e3a", "#9ca3af"

def to_bps(v):
    s = "" if v is None else str(v).strip()
    if s == "" or s.lower() == "nan":
        return None
    c = re.sub(r"[^\d\-+.]", "", s)
    return int(round(float(c))) if c not in ("", "-", "+", ".") else None

def to_num(v):
    try:
        f = float(v)
        return None if pd.isna(f) else f
    except (TypeError, ValueError):
        return None

# RULE 1: scale from the largest column total. NO renormalisation.
col_totals = [sum(to_num(v) or 0.0 for v in shares_df[c]) for c in columns]
max_total = max(col_totals) or 1.0

fig, ax = plt.subplots(figsize=(12, 7), dpi=150)
COL_W = 1.0

for ci, cname in enumerate(columns):
    y = 0.0
    for ri in range(len(shares_df)):
        share = to_num(shares_df[cname].iloc[ri])
        if not share or share <= 0:
            continue

        bps    = to_bps(bps_df[cname].iloc[ri])    if cname in bps_df    else None
        growth = to_num(growth_df[cname].iloc[ri]) if cname in growth_df else None
        label  = labels_df[cname].iloc[ri]         if cname in labels_df else ""
        label  = "" if pd.isna(label) else str(label).upper()

        color = NEUTRAL if bps is None else (POS if bps >= 0 else NEG)
        h = share

        ax.bar(ci, h, bottom=y, width=COL_W,
               color=color, edgecolor="white", linewidth=1.5)

        frac = h / max_total
        if frac >= 0.045:
            lbl_text = str(round(share, 1)) + "%\n" + label
            ax.text(ci, y + h/2, lbl_text, ha="center", va="center",
                    color="white", weight="bold", fontsize=10 if frac < 0.10 else 13)
            g_prefix = "+" if (growth and growth > 0) else ""
            g_txt = "n/a" if growth is None else (g_prefix + str(round(growth * 100, 1)) + "%")
            ax.text(ci - COL_W/2 + 0.02, y + h - max_total*0.008, g_txt,
                    ha="left", va="top", color="#1a1a1a", weight="bold", fontsize=7.5,
                    bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="none"))
            b_txt = "n/a" if bps is None else (str(bps) + " bps")
            ax.text(ci + COL_W/2 - 0.02, y + h - max_total*0.008, b_txt,
                    ha="right", va="top", color="white", weight="bold", fontsize=7.5)
        y += h

ax.set_xticks(range(len(columns)))
ax.set_xticklabels(columns, fontweight="bold", fontsize=13, color="#0284c7")
ax.set_xlim(-0.5, len(columns) - 0.5)
ax.set_ylim(0, max_total)
ax.set_yticks([])
for sp in ax.spines.values():
    sp.set_visible(False)
ax.tick_params(axis="x", length=0)
plt.tight_layout()
fig
```
