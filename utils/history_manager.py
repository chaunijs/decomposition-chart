"""
History Manager for tracking and caching recently used Excel and CSV files.
Saves metadata and file cache so users can easily switch between previously loaded workbooks.
"""

from __future__ import annotations
import os
import json
from datetime import datetime
from typing import List, Dict, Any, Optional

HISTORY_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), ".recent_files.json")
CACHE_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), ".history_cache")


def ensure_dirs():
    os.makedirs(CACHE_DIR, exist_ok=True)


def load_history() -> List[Dict[str, Any]]:
    """Loads list of recent files from JSON file."""
    ensure_dirs()
    workspace_dir = os.path.dirname(os.path.dirname(__file__))
    default_decomp = os.path.join(workspace_dir, "dataset-decomp.xlsx")

    history: List[Dict[str, Any]] = []
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                history = json.load(f)
        except Exception:
            history = []

    # Filter out entries where file no longer exists
    valid_history = []
    for item in history:
        p = item.get("path")
        if p and os.path.exists(p):
            valid_history.append(item)

    return valid_history


def save_history(history: List[Dict[str, Any]]) -> None:
    """Persists history list to JSON."""
    ensure_dirs()
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


def add_file_to_history(filename: str, file_bytes: bytes) -> Dict[str, Any]:
    """
    Saves an uploaded file to history cache and records it in recent files.
    Returns the created history entry.
    """
    ensure_dirs()
    cached_path = os.path.join(CACHE_DIR, filename)

    need_write = True
    if os.path.exists(cached_path) and os.path.getsize(cached_path) == len(file_bytes):
        need_write = False

    if need_write:
        try:
            with open(cached_path, "wb") as f:
                f.write(file_bytes)
        except Exception:
            pass

    size_kb = len(file_bytes) / 1024.0
    size_str = f"{size_kb / 1024.0:.1f} MB" if size_kb >= 1024 else f"{size_kb:.1f} KB"
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    history = load_history()
    # If the topmost entry is already this file, reuse it
    if history and history[0].get("path") == cached_path:
        return history[0]

    # Remove existing entry for the same filename if present
    history = [h for h in history if h.get("filename") != filename and h.get("path") != cached_path]

    entry = {
        "filename": filename,
        "display_name": filename,
        "path": cached_path,
        "size_str": size_str,
        "last_used": now_str
    }
    history.insert(0, entry)
    save_history(history[:10])
    return entry


def clear_history() -> None:
    """Clears uploaded file cache and resets history."""
    ensure_dirs()
    try:
        for f in os.listdir(CACHE_DIR):
            fp = os.path.join(CACHE_DIR, f)
            if os.path.isfile(fp):
                os.remove(fp)
    except Exception:
        pass

    save_history([])
