"""
Data Processing and Hierarchy Builder for Decomposition Chart using high-performance Polars.
"""

from __future__ import annotations
import os
import contextlib
import polars as pl
import json
from typing import List, Dict, Any, Optional, Tuple


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


def read_excel(
    file_or_path: Any,
    strict: bool = False,
    engine: str = "calamine",
    **kwargs: Any
) -> pl.DataFrame:
    """
    Reads an Excel spreadsheet into a Polars DataFrame.
    Supports strict=False parameter to handle mixed dtypes, empty rows, and schema mismatches gracefully.
    """
    with silence_rust_stderr():
        if not strict:
            # 1. Fast Calamine engine with coercion and full schema inference
            try:
                read_opts = kwargs.pop("read_options", {})
                read_opts.setdefault("dtype_coercion", "coerce")
                return pl.read_excel(
                    file_or_path,
                    engine="calamine",
                    read_options=read_opts,
                    infer_schema_length=None,
                    **kwargs
                )
            except Exception:
                pass

            # 2. Openpyxl fallback (handles complex formulas & unusual cell layouts)
            try:
                return pl.read_excel(file_or_path, engine="openpyxl", **kwargs)
            except Exception:
                pass

            # 3. Pandas fallback with Polars conversion (maximum compatibility)
            try:
                import pandas as pd
                pdf = pd.read_excel(file_or_path, **kwargs)
                return pl.from_pandas(pdf)
            except Exception:
                return pl.DataFrame()
        else:
            return pl.read_excel(file_or_path, engine=engine, **kwargs)


def load_dataset(file_or_path: Any, strict: bool = False) -> pl.DataFrame:
    """
    Loads a DataFrame from CSV, Excel, or JSON using high-performance Polars (Rust engine).
    Uses non-strict type coercion (strict=False) to seamlessly handle mixed or empty columns.
    """
    if hasattr(file_or_path, "name"):
        filename = file_or_path.name.lower()
        if filename.endswith(".csv"):
            return pl.read_csv(file_or_path, strict=strict)
        elif filename.endswith((".xlsx", ".xls")):
            return read_excel(file_or_path, strict=strict)
        elif filename.endswith(".json"):
            return pl.read_json(file_or_path)
        else:
            try:
                return pl.read_csv(file_or_path, strict=strict)
            except Exception:
                return read_excel(file_or_path, strict=strict)
    elif isinstance(file_or_path, str):
        if file_or_path.endswith(".csv"):
            return pl.read_csv(file_or_path, strict=strict)
        elif file_or_path.endswith((".xlsx", ".xls")):
            return read_excel(file_or_path, strict=strict)
        elif file_or_path.endswith(".json"):
            return pl.read_json(file_or_path)
    elif hasattr(file_or_path, "read") or isinstance(file_or_path, bytes):
        try:
            return read_excel(file_or_path, strict=strict)
        except Exception:
            try:
                if hasattr(file_or_path, "seek"):
                    file_or_path.seek(0)
                return pl.read_csv(file_or_path, strict=strict)
            except Exception:
                return pl.DataFrame()
    return pl.DataFrame()


def detect_columns(df: pl.DataFrame) -> Tuple[List[str], List[str]]:
    """Separates numerical metric columns from categorical dimension columns using Polars schema."""
    numerical_cols = []
    categorical_cols = []

    for col, dtype in df.schema.items():
        if dtype.is_numeric():
            numerical_cols.append(col)
        else:
            categorical_cols.append(col)

    # Fallback: if no categorical, include all columns
    if not categorical_cols and len(df.columns) > 1:
        categorical_cols = [c for c in df.columns if c != numerical_cols[0]]

    return numerical_cols, categorical_cols


def build_decomposition_tree(
    df: pl.DataFrame,
    metric_col: str,
    dimension_hierarchy: List[str],
    agg_func: str = "sum",
    top_n_per_level: int = 8,
    sort_descending: bool = True
) -> Dict[str, Any]:
    """
    Constructs a nested hierarchical tree dictionary using Polars multi-threaded aggregations.
    """
    if df.is_empty() or not metric_col or not dimension_hierarchy:
        return {}

    # Calculate root value
    if agg_func.lower() == "sum":
        root_val = float(df[metric_col].sum() or 0.0)
    elif agg_func.lower() in ("mean", "avg"):
        root_val = float(df[metric_col].mean() or 0.0)
    elif agg_func.lower() == "count":
        root_val = float(len(df))
    elif agg_func.lower() == "max":
        root_val = float(df[metric_col].max() or 0.0)
    elif agg_func.lower() == "min":
        root_val = float(df[metric_col].min() or 0.0)
    else:
        root_val = float(df[metric_col].sum() or 0.0)

    total_records = len(df)

    def recurse_tree(
        sub_df: pl.DataFrame,
        current_dim_idx: int,
        parent_id: str,
        parent_val: float
    ) -> List[Dict[str, Any]]:
        if current_dim_idx >= len(dimension_hierarchy) or sub_df.is_empty():
            return []

        dim = dimension_hierarchy[current_dim_idx]

        # Group by current dimension using Polars
        if agg_func.lower() == "sum":
            grouped = sub_df.group_by(dim).agg(pl.col(metric_col).sum().alias("value"))
        elif agg_func.lower() in ("mean", "avg"):
            grouped = sub_df.group_by(dim).agg(pl.col(metric_col).mean().alias("value"))
        elif agg_func.lower() == "count":
            grouped = sub_df.group_by(dim).agg(pl.len().alias("value"))
        elif agg_func.lower() == "max":
            grouped = sub_df.group_by(dim).agg(pl.col(metric_col).max().alias("value"))
        elif agg_func.lower() == "min":
            grouped = sub_df.group_by(dim).agg(pl.col(metric_col).min().alias("value"))
        else:
            grouped = sub_df.group_by(dim).agg(pl.col(metric_col).sum().alias("value"))

        grouped = grouped.sort("value", descending=sort_descending)
        grouped = grouped.with_columns(
            pl.col(dim).cast(pl.Utf8).fill_null("(Blank)").alias("label")
        )

        rows = grouped.to_dicts()

        # Handle top-n & 'Others' rollup if needed
        if len(rows) > top_n_per_level:
            top_rows = rows[:top_n_per_level]
            other_vals = [r["value"] for r in rows[top_n_per_level:] if r["value"] is not None]
            if agg_func.lower() in ("sum", "count"):
                other_val = sum(other_vals)
            elif agg_func.lower() in ("mean", "avg"):
                other_val = sum(other_vals) / len(other_vals) if other_vals else 0.0
            else:
                other_val = max(other_vals) if other_vals else 0.0

            top_rows.append({
                dim: "Others",
                "value": other_val,
                "label": f"Others ({len(rows) - top_n_per_level} items)"
            })
            display_rows = top_rows
        else:
            display_rows = rows

        children = []
        for i, row in enumerate(display_rows):
            child_label = row["label"]
            child_val = float(row["value"] or 0.0)
            child_id = f"{parent_id}_d{current_dim_idx}_{i}"

            # Calculate percentages
            pct_of_parent = (child_val / parent_val * 100.0) if parent_val != 0 else 0.0
            pct_of_root = (child_val / root_val * 100.0) if root_val != 0 else 0.0

            # Filter sub-df for next level
            if child_label.startswith("Others"):
                next_children = []
            else:
                raw_filter_val = row[dim]
                if raw_filter_val is None or raw_filter_val == "(Blank)":
                    next_sub_df = sub_df.filter(pl.col(dim).is_null())
                else:
                    next_sub_df = sub_df.filter(pl.col(dim) == raw_filter_val)

                next_children = recurse_tree(
                    next_sub_df,
                    current_dim_idx + 1,
                    child_id,
                    child_val
                )

            children.append({
                "id": child_id,
                "label": child_label,
                "dimension": dim,
                "value": child_val,
                "pct_of_parent": round(pct_of_parent, 2),
                "pct_of_root": round(pct_of_root, 2),
                "count": len(sub_df.filter(pl.col(dim) == row[dim])) if dim in sub_df.columns else 0,
                "children": next_children
            })

        return children

    root_node = {
        "id": "root",
        "label": f"Total {metric_col.replace('_', ' ').title()}",
        "dimension": "Total",
        "value": root_val,
        "pct_of_parent": 100.0,
        "pct_of_root": 100.0,
        "count": total_records,
        "children": recurse_tree(df, 0, "root", root_val)
    }

    return {
        "metric": metric_col,
        "agg_func": agg_func,
        "dimensions": dimension_hierarchy,
        "total_records": total_records,
        "root": root_node
    }
