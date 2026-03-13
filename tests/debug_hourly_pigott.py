"""
Debug harness for Hourly Room Utilization exports.

Focuses on a single building (Pigott) to show:

1) How rows look in the long-format DataFrame returned by
   transform_hourly_utilization().
2) How the same rows look in the wide-format pivot, including
   hour columns and the Average position.
3) How DataFrame columns map to Excel columns (A, B, C, ...).

Usage (from project root):

    python -m tests.debug_hourly_pigott path/to/EMS_hourly_export.xls
"""

import sys
import re
from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter

# Adjust import paths as needed for your project layout
from cleaner.hourly.transform import transform_hourly_utilization


# --- Helper functions copied from writer logic -----------------------------

def split_room_fields(building_name: str, room_full: str):
    """
    Split 'Pigott 103 - Caseroom' into:
      Building, Room Number, Classroom Type

    This mirrors the logic in hourly/writer.py so the test
    matches the real export behavior.
    """
    room_full = str(room_full).strip()
    room_number = room_full
    room_type = ""

    # Case 1: a dash separates number and type
    if " - " in room_full:
        left, right = room_full.split(" - ", 1)
        left = left.strip()
        right = right.strip()

        # Try to find a number-like token in the left part
        m = re.search(r"\b(\d+[A-Za-z0-9/-]*)\b", left)
        if m:
            room_number = m.group(1)
        else:
            # fallback: use everything after building name
            room_number = left.replace(building_name, "", 1).strip()

        room_type = right
        return building_name, room_number, room_type

    # Case 2: no dash → try to pull a number-ish token from remainder
    cleaned = room_full.replace(building_name, "", 1).strip()
    m = re.search(r"\b(\d+[A-Za-z0-9/-]*)\b", cleaned)
    if m:
        room_number = m.group(1)
        room_type = cleaned.replace(room_number, "").strip()
    else:
        # No obvious number: treat the whole cleaned string as room_number
        # and leave room_type blank
        room_number = cleaned
        room_type = ""

    return building_name, room_number.strip(), room_type.strip()


def format_hour_label(label: str) -> str:
    """Convert 'xa' -> 'x AM', 'xp' -> 'x PM'."""
    m = re.match(r"^(\d+)([ap])$", str(label).strip().lower())
    if not m:
        return label
    h = int(m.group(1))
    ap = m.group(2)
    return f"{h} {'AM' if ap == 'a' else 'PM'}"


def build_wide_from_long(df_long: pd.DataFrame) -> pd.DataFrame:
    """
    Rebuild the wide-format DataFrame the same way writer.py does,
    but in-memory, so we can inspect columns and rows.
    """
    df = df_long.copy()

    # Derive Room Number and Classroom Type from Room text
    buildings = []
    room_numbers = []
    room_types = []

    for _, row in df.iterrows():
        b, rn, rt = split_room_fields(row["Building"], row["Room"])
        buildings.append(b)
        room_numbers.append(rn)
        room_types.append(rt)

    df["Building"] = buildings
    df["Room Number"] = room_numbers
    df["Classroom Type"] = room_types

    df = df.drop(columns=["Room"])

    # Base index: all unique room rows that should appear in the wide table
    base_index = df[["Building", "Room Number", "Classroom Type"]].drop_duplicates()

    if df.empty:
        wide = pd.DataFrame(columns=["Building", "Room Number", "Classroom Type"])
        return wide

    # Pivot: one row per room, columns per Hour
    wide_pivot = df.pivot_table(
        index=["Building", "Room Number", "Classroom Type"],
        columns="Hour",
        values="Value",
        aggfunc=lambda x: x.iloc[0],
    ).reset_index()

    # Flatten MultiIndex columns, if any
    wide_pivot.columns = [
        col if isinstance(col, str) else col[1] for col in wide_pivot.columns
    ]

    # Merge base list of rooms with pivot to ensure rooms with no data are retained
    wide = base_index.merge(
        wide_pivot,
        on=["Building", "Room Number", "Classroom Type"],
        how="left",
    )

    # Drop EMS 'Average' if it sneaks in
    if "Average" in wide.columns:
        wide = wide.drop(columns=["Average"])

    # Define expected hour columns (superset 6a–9p)
    expected_hours = [
        "6a", "7a", "8a", "9a", "10a", "11a",
        "12p", "1p", "2p", "3p", "4p", "5p",
        "6p", "7p", "8p", "9p",
    ]

    # Ensure all expected hour columns exist, even if empty
    for h in expected_hours:
        if h not in wide.columns:
            wide[h] = pd.NA

    # Sort hour columns logically
    def hour_key(label):
        if label in ("Building", "Room Number", "Classroom Type"):
            return -1
        m = re.match(r"^(\d+)([ap])$", str(label).strip().lower())
        if not m:
            return 999
        hh = int(m.group(1))
        ap = m.group(2)
        if ap == "a":
            h24 = hh if hh != 12 else 0
        else:
            h24 = hh + 12 if hh != 12 else 12
        return h24

    base_cols = ["Building", "Room Number", "Classroom Type"]
    hour_cols = [c for c in wide.columns if c not in base_cols]
    hour_cols_sorted = sorted(hour_cols, key=hour_key)

    # Convert hour columns to numeric, fill empty with 0
    for h in hour_cols_sorted:
        wide[h] = pd.to_numeric(wide[h], errors="coerce").fillna(0)

    # Add Average column at the end
    wide["Average"] = 0.0

    # Reorder: base, hours, Average
    wide = wide[base_cols + hour_cols_sorted + ["Average"]]

    # Rename hour labels for display
    rename_map = {h: format_hour_label(h) for h in hour_cols_sorted}
    wide = wide.rename(columns=rename_map)

    return wide


# --- Main debug routine ----------------------------------------------------

def debug_pigott(input_path: str):
    """
    Run transform_hourly_utilization on the given EMS export,
    then show long-format and wide-format views for Pigott Building.
    """
    print(f"\n=== Running transform_hourly_utilization on {input_path} ===")
    df_long = transform_hourly_utilization(input_path)
    print(f"Long-format total rows: {len(df_long)}")

    # Filter for Pigott-related rows in long-format
    mask_pigott = df_long["Building"].astype(str).str.contains("Pigott", na=False) | \
                  df_long["Room"].astype(str).str.contains("Pigott", na=False)

    df_pigott_long = df_long[mask_pigott].copy()

    print("\n=== Long-format: Pigott rows (sample) ===")
    print(df_pigott_long.head(20))

    print("\nUnique Pigott rooms (long):")
    print(df_pigott_long["Room"].drop_duplicates().tolist())

    print("\nUnique Pigott hours (long):")
    print(sorted(df_pigott_long["Hour"].dropna().unique().tolist()))

    # Build wide-format frame as the writer would
    wide = build_wide_from_long(df_pigott_long)

    print("\n=== Wide-format: Pigott rows (head) ===")
    print(wide.head(20))

    # Show column index → column name → Excel column letter mapping
    print("\n=== Wide-format columns and Excel mapping ===")
    for idx, col_name in enumerate(wide.columns):
        excel_col = get_column_letter(idx + 1)  # 1-based for Excel
        print(f"{idx:2d}  {excel_col:>2}  {col_name}")

    # Show specifically Pigott 103 - Caseroom if present
    mask_103 = wide["Room Number"].astype(str).eq("103")
    df_103 = wide[mask_103]

    print("\n=== Wide-format row(s) for Room Number '103' (Pigott 103 - Caseroom) ===")
    if df_103.empty:
        print("No rows found for Room Number '103'.")
    else:
        print(df_103)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(
            "Usage:\n"
            "  python -m tests.debug_hourly_pigott path/to/EMS_hourly_export.xls"
        )
        sys.exit(1)

    input_file = sys.argv[1]
    p = Path(input_file)

    if not p.is_file():
        print(f"Error: '{p}' is not a file or does not exist.")
        sys.exit(1)

    debug_pigott(str(p))