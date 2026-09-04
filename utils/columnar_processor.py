"""
Processor for Multi-Column Hierarchical Decomposition / Marimekko Chart
with support for Excel Named Ranges (dataset-decomp.xlsx).
"""

from __future__ import annotations
import openpyxl
import re
import polars as pl
import pandas as pd
import numpy as np
from typing import List, Dict, Any, Optional, Tuple, Union

DEFAULT_DIMENSION_LABELS = {
    "Total TH": ["TOTAL"],
    "Channel": ["SH CHAIN", "SH Online", "OT", "LMT", "CVS"],
    "Price": ["Tier1", "Tier2", "Tier3", "Tier4", "Tier5"],
    "Region": ["GBKK", "CENTRAL", "NORTH", "NORTHEAST", "SOUTH"],
    "Packsize": ["S", "M", "L", "XL", "XXL"]
}


def parse_named_range_matrix(wb: openpyxl.Workbook, range_name: str) -> Tuple[List[str], List[List[Any]]]:
    """Extracts header and rows from an Excel defined name."""
    if range_name not in wb.defined_names:
        return [], []
    d = wb.defined_names[range_name]
    for title, coord in d.destinations:
        sheet = wb[title]
        rows = list(sheet[coord]) if ":" in coord else [[sheet[coord]]]
        header = [str(c.value).strip() if c.value is not None else "" for c in rows[0]]
        matrix = []
        for r in rows[1:]:
            matrix.append([c.value for c in r])
        return header, matrix
    return [], []


def extract_dimension_labels(
    wb: openpyxl.Workbook,
    prod: str,
    cont_name: str,
    cont_headers: List[str]
) -> Dict[str, List[str]]:
    """
    Dynamically extracts dimension labels from dataset-decomp.xlsx without hardcoding.
    
    Discovery strategy:
      1. Explicit defined name (e.g. label_{prod}, labels_{prod}, {prod}_labels, label_table, dim_labels)
      2. Label table located to the left / adjacent to cont_* range in the same worksheet
      3. A dedicated 'Labels' or 'Label' worksheet
      4. Fallback to DEFAULT_DIMENSION_LABELS if not defined in workbook
    """
    dynamic_labels: Dict[str, List[str]] = {}
    from openpyxl.utils.cell import range_boundaries

    # Strategy 1: Explicit named range for product labels (e.g. label_powder, label_liquid)
    for cand in [f"label_{prod}", f"labels_{prod}", f"{prod}_label", f"{prod}_labels", "labels", "label_table", "dim_labels", "dimension_labels"]:
        if cand in wb.defined_names:
            d = wb.defined_names[cand]
            for title, coord in d.destinations:
                sheet = wb[title]
                min_col, min_row, max_col, max_row = range_boundaries(coord)
                
                # Check whether min_row contains header names or is already the first data row (e.g. $A$3:$E$7)
                first_row_vals = [str(sheet.cell(min_row, c).value or "").strip() for c in range(min_col, max_col + 1)]
                has_headers = any(h.lower() in [ch.lower() for ch in cont_headers] for h in first_row_vals if h)
                
                if has_headers:
                    data_start = min_row + 1
                    col_names = first_row_vals
                else:
                    data_start = min_row
                    col_names = cont_headers

                for idx, c_name in enumerate(col_names):
                    if idx < len(cont_headers):
                        target_key = cont_headers[idx] if not has_headers else c_name
                        if target_key:
                            dynamic_labels[target_key] = [
                                str(sheet.cell(r, min_col + idx).value or "").strip()
                                for r in range(data_start, max_row + 1)
                            ]
            if dynamic_labels:
                return dynamic_labels

    # Strategy 2: Label table directly adjacent / to the left of cont_* range
    if cont_name in wb.defined_names:
        d = wb.defined_names[cont_name]
        for title, coord in d.destinations:
            sheet = wb[title]
            min_col, min_row, max_col, max_row = range_boundaries(coord)

            header_row = min_row
            label_cols: Dict[str, int] = {}
            for c in range(1, min_col):
                val = sheet.cell(header_row, c).value
                if val is not None and str(val).strip():
                    label_cols[str(val).strip()] = c

            if not label_cols and min_row > 1:
                header_row = min_row - 1
                for c in range(1, min_col):
                    val = sheet.cell(header_row, c).value
                    if val is not None and str(val).strip():
                        label_cols[str(val).strip()] = c

            for h_name, c_idx in label_cols.items():
                items = []
                for r in range(min_row + 1, max_row + 1):
                    val = sheet.cell(r, c_idx).value
                    items.append(str(val).strip() if val is not None else "")
                dynamic_labels[h_name] = items

            if dynamic_labels:
                return dynamic_labels

    # Strategy 3: Dedicated worksheet named Labels or Label
    for s_name in wb.sheetnames:
        if s_name.lower() in ["labels", "label", "dimension labels", "dim_labels"]:
            sheet = wb[s_name]
            headers = [str(sheet.cell(1, c).value or "").strip() for c in range(1, sheet.max_column + 1)]
            for idx, h in enumerate(headers):
                if h:
                    dynamic_labels[h] = [
                        str(sheet.cell(r, idx + 1).value or "").strip()
                        for r in range(2, sheet.max_row + 1)
                    ]
            if dynamic_labels:
                return dynamic_labels

    return dynamic_labels

import zipfile
import os
import io
import contextlib
import xml.etree.ElementTree as ET
try:
    import fastexcel
    HAS_FASTEXCEL = True
except ImportError:
    HAS_FASTEXCEL = False


@contextlib.contextmanager
def silence_rust_stderr():
    """Suppresses low-level C/Rust stderr messages like Calamine type inference warnings."""
    try:
        devnull = os.open(os.devnull, os.O_WRONLY)
        old_stderr = os.dup(2)
        os.dup2(devnull, 2)
        os.close(devnull)
        yield
    except Exception:
        yield
    finally:
        try:
            os.dup2(old_stderr, 2)
            os.close(old_stderr)
        except Exception:
            pass


def fast_extract_excel_named_ranges(file_source: Union[str, bytes, io.BytesIO]) -> Optional[Dict[str, Dict[str, Any]]]:
    """
    Ultra-fast named range extractor using raw zipfile parsing (milliseconds) and fastexcel.
    Avoids openpyxl's heavy overhead on large multi-sheet workbooks.
    """
    if not HAS_FASTEXCEL:
        return None

    try:
        from openpyxl.utils.cell import range_boundaries

        if isinstance(file_source, bytes):
            zf = zipfile.ZipFile(io.BytesIO(file_source))
            f_input = io.BytesIO(file_source)
        elif hasattr(file_source, "read"):
            b = file_source.read()
            zf = zipfile.ZipFile(io.BytesIO(b))
            f_input = io.BytesIO(b)
        else:
            zf = zipfile.ZipFile(file_source)
            f_input = file_source

        with zf:
            wb_xml = zf.read("xl/workbook.xml")
            root = ET.fromstring(wb_xml)
            ns = {"ns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            defined_names = {}
            for dn in root.findall(".//ns:definedName", ns):
                name = dn.attrib.get("name")
                text = dn.text
                if name and text:
                    defined_names[name] = text.strip()

        # Identify unique product names
        products = set()
        for n in defined_names.keys():
            for prefix in ["bps_", "cont_", "contribution_", "growth_pct_", "growth_"]:
                if n.startswith(prefix):
                    prod = n.replace(prefix, "")
                    products.add(prod)
                elif n.endswith("_" + prefix.rstrip("_")):
                    prod = n.split("_")[0]
                    products.add(prod)

        if not products:
            return None

        with silence_rust_stderr():
            excel = fastexcel.read_excel(f_input)
        loaded_sheets = {}

        def get_sheet_df(sheet_name: str) -> pl.DataFrame:
            clean_name = sheet_name.strip("'")
            if clean_name not in loaded_sheets:
                with silence_rust_stderr():
                    sheet = excel.load_sheet_by_name(clean_name, header_row=None, dtype_coercion="coerce")
                    loaded_sheets[clean_name] = sheet.to_polars()
            return loaded_sheets[clean_name]

        def parse_range(ref_str: str):
            if "!" not in ref_str:
                return None, None, None, None, None
            sheet_part, coord_part = ref_str.split("!", 1)
            clean_sheet = sheet_part.strip("'")
            clean_coord = coord_part.replace("$", "")
            min_col, min_row, max_col, max_row = range_boundaries(clean_coord)
            return clean_sheet, min_col, min_row, max_col, max_row

        results = {}

        for prod in sorted(products):
            cont_ref = next((defined_names[c] for c in [f"contribution_{prod}", f"cont_{prod}", f"{prod}_cont"] if c in defined_names), None)
            bps_ref = next((defined_names[c] for c in [f"bps_{prod}", f"{prod}_bps"] if c in defined_names), None)
            gw_ref = next((defined_names[c] for c in [f"growth_pct_{prod}", f"growth_{prod}", f"{prod}_growth"] if c in defined_names), None)
            lbl_ref = next((defined_names[c] for c in [f"label_{prod}", f"labels_{prod}", f"{prod}_label", f"{prod}_labels"] if c in defined_names), None)

            if not cont_ref:
                continue

            sheet_name, c_min_c, c_min_r, c_max_c, c_max_r = parse_range(cont_ref)
            if not sheet_name:
                continue

            sheet_df = get_sheet_df(sheet_name)
            
            # Header values
            header_row_vals = sheet_df.slice(c_min_r - 1, 1).select(sheet_df.columns[c_min_c - 1 : c_max_c]).to_dicts()[0]
            cont_headers = [str(v or "").strip() for v in header_row_vals.values()]

            # Matrix values
            data_rows = sheet_df.slice(c_min_r, c_max_r - c_min_r).select(sheet_df.columns[c_min_c - 1 : c_max_c]).to_dicts()
            cont_m = [[r[col] for col in sheet_df.columns[c_min_c - 1 : c_max_c]] for r in data_rows]

            # BPS
            bps_m = []
            if bps_ref:
                b_s, b_min_c, b_min_r, b_max_c, b_max_r = parse_range(bps_ref)
                b_df = get_sheet_df(b_s)
                b_rows = b_df.slice(b_min_r, b_max_r - b_min_r).select(b_df.columns[b_min_c - 1 : b_max_c]).to_dicts()
                bps_m = [[r[col] for col in b_df.columns[b_min_c - 1 : b_max_c]] for r in b_rows]

            # Growth
            gw_m = []
            if gw_ref:
                g_s, g_min_c, g_min_r, g_max_c, g_max_r = parse_range(gw_ref)
                g_df = get_sheet_df(g_s)
                g_rows = g_df.slice(g_min_r, g_max_r - g_min_r).select(g_df.columns[g_min_c - 1 : g_max_c]).to_dicts()
                gw_m = [[r[col] for col in g_df.columns[g_min_c - 1 : g_max_c]] for r in g_rows]

            # Dynamic Labels
            labels_by_col = {}
            if lbl_ref:
                l_s, l_min_c, l_min_r, l_max_c, l_max_r = parse_range(lbl_ref)
                l_df = get_sheet_df(l_s)
                l_rows = l_df.slice(l_min_r - 1, l_max_r - l_min_r + 1).select(l_df.columns[l_min_c - 1 : l_max_c]).to_dicts()
                first_l = [str(v or "").strip() for v in l_rows[0].values()]
                has_h = any(h.lower() in [ch.lower() for ch in cont_headers] for h in first_l if h)
                val_rows = l_rows[1:] if has_h else l_rows
                
                for idx, ch in enumerate(cont_headers):
                    c_id = l_df.columns[l_min_c - 1 + idx] if idx < (l_max_c - l_min_c + 1) else None
                    if c_id:
                        labels_by_col[ch] = [str(r.get(c_id) or "").strip() for r in val_rows]
            else:
                # Check columns to the left of cont_*
                if c_min_c > 1:
                    left_cols = sheet_df.columns[:c_min_c - 1]
                    header_r = sheet_df.slice(c_min_r - 1, 1).select(left_cols).to_dicts()[0]
                    left_headers = {str(v).strip(): col for col, v in header_r.items() if v}
                    val_rows = sheet_df.slice(c_min_r, c_max_r - c_min_r).select(left_cols).to_dicts()
                    for ch, c_id in left_headers.items():
                        labels_by_col[ch] = [str(r.get(c_id) or "").strip() for r in val_rows]

            category_display_name = prod.upper()
            total_key = next((k for k in labels_by_col.keys() if "total" in k.lower()), None)
            if total_key and labels_by_col[total_key] and labels_by_col[total_key][0]:
                category_display_name = labels_by_col[total_key][0]

            columns_data = []
            flat_records = []

            for col_idx, col_name in enumerate(cont_headers):
                if not col_name:
                    continue

                dim_labels = labels_by_col.get(col_name) or []
                if not dim_labels:
                    for k, v in labels_by_col.items():
                        if k.strip().lower() == col_name.strip().lower():
                            dim_labels = v
                            break

                blocks = []
                for row_idx in range(len(cont_m)):
                    share_val = cont_m[row_idx][col_idx] if row_idx < len(cont_m) and col_idx < len(cont_m[row_idx]) else None
                    if share_val is None or str(share_val).strip() == "":
                        continue

                    try:
                        share_num = float(share_val)
                    except (ValueError, TypeError):
                        continue

                    share_pct = round(share_num * 100.0 if share_num <= 1.0 and share_num > 0 else share_num, 1)

                    bps_num = 0
                    if bps_m and row_idx < len(bps_m) and col_idx < len(bps_m[row_idx]):
                        b_raw = bps_m[row_idx][col_idx]
                        if b_raw is not None and str(b_raw).strip() != "":
                            b_clean = re.sub(r"[^\d\-+.]", "", str(b_raw))
                            try:
                                bps_num = int(round(float(b_clean)))
                            except ValueError:
                                bps_num = 0

                    gw_pct = 0.0
                    if gw_m and row_idx < len(gw_m) and col_idx < len(gw_m[row_idx]):
                        g_raw = gw_m[row_idx][col_idx]
                        if g_raw is not None and str(g_raw).strip() != "":
                            try:
                                g_num = float(g_raw)
                                gw_pct = round(g_num * 100.0, 1) if abs(g_num) <= 5.0 and g_num != 0 else round(g_num, 1)
                            except ValueError:
                                gw_pct = 0.0

                    if "total" in col_name.lower():
                        actual_label = dim_labels[row_idx] if dim_labels and row_idx < len(dim_labels) and dim_labels[row_idx] else category_display_name
                        generic_label = category_display_name
                    elif dim_labels and row_idx < len(dim_labels) and dim_labels[row_idx]:
                        actual_label = dim_labels[row_idx]
                        generic_label = f"ITEM {row_idx+1}"
                    else:
                        actual_label = f"ITEM {row_idx+1}"
                        generic_label = f"ITEM {row_idx+1}"

                    is_pos = (bps_num >= 0 if bps_ref else gw_pct >= 0)
                    blocks.append({
                        "label": actual_label,
                        "actual_label": actual_label,
                        "generic_label": generic_label,
                        "share": share_pct,
                        "growth": gw_pct,
                        "bps": bps_num,
                        "is_positive": is_pos
                    })
                    flat_records.append({
                        "Product": category_display_name,
                        "Column": col_name,
                        "Label": actual_label,
                        "Share %": share_pct,
                        "Growth %": gw_pct,
                        "BPS Impact": bps_num
                    })

                columns_data.append({
                    "column_name": col_name,
                    "blocks": blocks,
                    "total_share": sum(b["share"] for b in blocks)
                })

            results[category_display_name] = {
                "columns_data": columns_data,
                "df": pl.DataFrame(flat_records)
            }

        return results
    except Exception:
        return None


def extract_excel_named_ranges(
    excel_path: Union[str, bytes, io.BytesIO],
    **kwargs: Any
) -> Dict[str, Dict[str, Any]]:
    """
    Parses workbook named ranges using openpyxl for 100% cell coordinate accuracy.
    Returns a dictionary of product_name -> {columns_data, df}.
    """

    # 2. Graceful fallback to openpyxl
    if isinstance(excel_path, bytes):
        wb_input = io.BytesIO(excel_path)
    else:
        wb_input = excel_path

    wb = openpyxl.load_workbook(wb_input, data_only=True)
    all_names = list(wb.defined_names.keys()) if hasattr(wb.defined_names, "keys") else [d.name for d in wb.defined_names.definedName]

    # Find unique product prefixes/suffixes
    products = set()
    for n in all_names:
        for prefix in ["bps_", "cont_", "contribution_", "growth_pct_", "growth_"]:
            if n.startswith(prefix):
                prod = n.replace(prefix, "")
                products.add(prod)
            elif n.endswith("_" + prefix.rstrip("_")):
                prod = n.split("_")[0]
                products.add(prod)

    results = {}

    for prod in sorted(products):
        # Locate corresponding cont, bps, growth names
        cont_name = None
        for cand in [f"contribution_{prod}", f"cont_{prod}", f"{prod}_cont", f"{prod}_contribution"]:
            if cand in wb.defined_names:
                cont_name = cand
                break

        bps_name = None
        for cand in [f"bps_{prod}", f"{prod}_bps"]:
            if cand in wb.defined_names:
                bps_name = cand
                break

        growth_name = None
        for cand in [f"growth_pct_{prod}", f"growth_{prod}", f"{prod}_growth"]:
            if cand in wb.defined_names:
                growth_name = cand
                break

        if not cont_name:
            continue

        cont_h, cont_m = parse_named_range_matrix(wb, cont_name)
        bps_h, bps_m = parse_named_range_matrix(wb, bps_name) if bps_name else ([], [])
        gw_h, gw_m = parse_named_range_matrix(wb, growth_name) if growth_name else ([], [])

        # Extract dynamic dimension labels from the Excel workbook table
        dynamic_labels = extract_dimension_labels(wb, prod, cont_name, cont_h)

        # Check if Total column has an explicit category name in the label table (e.g. 'Powder', 'DW')
        total_col_key = next((k for k in dynamic_labels.keys() if "total" in k.lower()), None)
        raw_cat_name = dynamic_labels[total_col_key][0] if (total_col_key and dynamic_labels[total_col_key] and dynamic_labels[total_col_key][0]) else prod
        category_display_name = str(raw_cat_name).strip().upper()

        columns_data = []
        flat_records = []

        for col_idx, col_name in enumerate(cont_h):
            if not col_name:
                continue

            # Resolve labels for this column from dynamic label table
            dim_labels = dynamic_labels.get(col_name)
            if not dim_labels:
                # Case-insensitive / normalized lookup
                for k, v in dynamic_labels.items():
                    if k.strip().lower() == col_name.strip().lower():
                        dim_labels = v
                        break
            if not dim_labels:
                # Fallback to default if no label table found in workbook
                dim_labels = DEFAULT_DIMENSION_LABELS.get(col_name, [])

            blocks = []

            for row_idx in range(len(cont_m)):
                share_val = cont_m[row_idx][col_idx] if row_idx < len(cont_m) and col_idx < len(cont_m[row_idx]) else None
                if share_val is None or str(share_val).strip() == "":
                    continue

                try:
                    share_num = float(share_val)
                except (ValueError, TypeError):
                    continue

                # Scale share from 0-1 decimal to percentage if needed
                if share_num <= 1.0 and share_num > 0:
                    share_pct = round(share_num * 100.0, 1)
                else:
                    share_pct = round(share_num, 1)

                # BPS
                bps_num = 0
                if bps_m and row_idx < len(bps_m) and col_idx < len(bps_m[row_idx]):
                    bps_raw = bps_m[row_idx][col_idx]
                    if bps_raw is not None and str(bps_raw).strip() != "":
                        bps_clean = re.sub(r"[^\d\-+.]", "", str(bps_raw))
                        try:
                            bps_num = int(round(float(bps_clean)))
                        except ValueError:
                            bps_num = 0

                # Growth %
                gw_pct = 0.0
                if gw_m and row_idx < len(gw_m) and col_idx < len(gw_m[row_idx]):
                    gw_raw = gw_m[row_idx][col_idx]
                    if gw_raw is not None and str(gw_raw).strip() != "":
                        try:
                            gw_num = float(gw_raw)
                            if abs(gw_num) <= 5.0 and gw_num != 0:
                                gw_pct = round(gw_num * 100.0, 1)
                            else:
                                gw_pct = round(gw_num, 1)
                        except ValueError:
                            gw_pct = 0.0

                # Actual label and generic label
                is_total_col = "total" in col_name.lower()
                if is_total_col:
                    if dim_labels and row_idx < len(dim_labels) and dim_labels[row_idx]:
                        actual_label = dim_labels[row_idx]
                    else:
                        actual_label = category_display_name
                    generic_label = category_display_name
                elif dim_labels and row_idx < len(dim_labels) and dim_labels[row_idx]:
                    actual_label = dim_labels[row_idx]
                    generic_label = f"ITEM {row_idx + 1}"
                else:
                    actual_label = f"ITEM {row_idx + 1}"
                    generic_label = f"ITEM {row_idx + 1}"

                # Positive / negative condition: strictly BPS >= 0 (or Growth >= 0 if BPS unavailable)
                is_pos = (bps_num >= 0 if bps_name else gw_pct >= 0)

                block_item = {
                    "label": actual_label,
                    "actual_label": actual_label,
                    "generic_label": generic_label,
                    "share": share_pct,
                    "growth": gw_pct,
                    "bps": bps_num,
                    "is_positive": is_pos
                }
                blocks.append(block_item)
                flat_records.append({
                    "Product": category_display_name,
                    "Column": col_name,
                    "Label": actual_label,
                    "Share %": share_pct,
                    "Growth %": gw_pct,
                    "BPS Impact": bps_num
                })

            columns_data.append({
                "column_name": col_name,
                "blocks": blocks,
                "total_share": sum(b["share"] for b in blocks)
            })

        results[category_display_name] = {
            "columns_data": columns_data,
            "df": pl.DataFrame(flat_records)
        }

    return results


def compute_columnar_decomposition(
    df: Union[pl.DataFrame, pd.DataFrame],
    dimensions: List[str],
    metric_col: str,
    growth_col: Optional[str] = None,
    bps_col: Optional[str] = None,
    benchmark_total: Optional[float] = None
) -> List[Dict[str, Any]]:
    """
    Computes columnar decomposition for any custom generic dataset using high-performance Polars.
    """
    if isinstance(df, pd.DataFrame):
        df = pl.from_pandas(df)

    lower_cols = [c.lower() for c in df.columns]
    if "column" in lower_cols and "label" in lower_cols and "share" in lower_cols:
        col_map = {c.lower(): c for c in df.columns}
        result = []
        ordered_columns = df[col_map["column"]].unique().to_list()
        for col_name in ordered_columns:
            sub = df.filter(pl.col(col_map["column"]) == col_name)
            blocks = []
            for b_idx, row in enumerate(sub.to_dicts()):
                growth_val = row.get(col_map.get("growth", "growth"), 0.0) or 0.0
                bps_val = row.get(col_map.get("bps", "bps"), 0) or 0
                actual_lbl = str(row[col_map["label"]])
                generic_lbl = f"ITEM {b_idx + 1}"
                blocks.append({
                    "label": actual_lbl,
                    "actual_label": actual_lbl,
                    "generic_label": generic_lbl,
                    "share": float(row[col_map["share"]]),
                    "growth": float(growth_val),
                    "bps": int(bps_val),
                    "is_positive": (bps_val >= 0 if col_map.get("bps") else growth_val >= 0)
                })
            result.append({
                "column_name": col_name,
                "blocks": blocks,
                "total_share": sum(b["share"] for b in blocks)
            })
        return result

    result = []
    total_val = benchmark_total or float(df[metric_col].sum() or 0.0)

    for dim in dimensions:
        if dim not in df.columns:
            continue

        aggs = [pl.col(metric_col).sum().alias(metric_col)]
        if growth_col and growth_col in df.columns:
            aggs.append(pl.col(growth_col).mean().alias("growth"))
        if bps_col and bps_col in df.columns:
            aggs.append(pl.col(bps_col).sum().alias("bps"))

        grouped = df.group_by(dim).agg(aggs)
        grouped = grouped.with_columns(
            (pl.col(metric_col) / total_val * 100.0).round(1).alias("share")
        )

        if "growth" not in grouped.columns:
            grouped = grouped.with_columns(pl.lit(0.0).alias("growth"))
        else:
            grouped = grouped.with_columns(pl.col("growth").round(1))

        if "bps" not in grouped.columns:
            grouped = grouped.with_columns((pl.col("growth") * 10).cast(pl.Int64).alias("bps"))
        else:
            grouped = grouped.with_columns(pl.col("bps").round(0).cast(pl.Int64))

        grouped = grouped.sort(metric_col, descending=True)

        blocks = []
        for b_idx, row in enumerate(grouped.to_dicts()):
            growth_val = float(row["growth"] or 0.0)
            bps_val = int(row["bps"] or 0)
            is_pos = (bps_val >= 0 if bps_col else growth_val >= 0)
            actual_lbl = str(row[dim])
            generic_lbl = f"ITEM {b_idx + 1}"
            blocks.append({
                "label": actual_lbl,
                "actual_label": actual_lbl,
                "generic_label": generic_lbl,
                "share": float(row["share"]),
                "growth": growth_val,
                "bps": bps_val,
                "is_positive": is_pos,
                "raw_value": float(row[metric_col] or 0.0)
            })

        result.append({
            "column_name": dim,
            "blocks": blocks,
            "total_share": sum(b["share"] for b in blocks)
        })

    return result
