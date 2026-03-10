# cleaner/hourly/transform.py

import os
import re
from typing import List

import pandas as pd


def _read_hourly_raw(input_path: str) -> pd.DataFrame:
    """
    Read raw Hourly Room Utilization export into a pandas DataFrame
    with no header row, dropping fully empty rows.
    """
    ext = os.path.splitext(input_path)[1].lower()

    if ext == ".csv":
        df = pd.read_csv(input_path, header=None)
    elif ext == ".xlsx":
        df = pd.read_excel(input_path, header=None, engine="openpyxl")
    elif ext == ".xls":
        df = pd.read_excel(input_path, header=None, engine="xlrd")
    else:
        raise ValueError(f"Unsupported file type: {ext}. Use .xls, .xlsx, or .csv.")

    df = df.dropna(how="all").copy()
    df.reset_index(drop=True, inplace=True)
    return df


def transform_hourly_utilization(input_path: str) -> pd.DataFrame:
    """
    Transform an EMS Hourly Room Utilization export into long-format DataFrame.

    IMPORTANT CHANGE:
        - Hour column values are imported EXACTLY as they appear in the EMS report.
        - No numeric conversion, no percent normalization.

    Output columns:
        Building
        Room
        Hour
        Value
        Reporting Period
    """
    df = _read_hourly_raw(input_path)

    # 1. Extract Reporting Period suffix once
    reporting_period_suffix = None
    for _, row in df.iterrows():
        for val in row:
            if isinstance(val, str) and val.startswith("Reporting Period:"):
                reporting_period_suffix = val.split("Reporting Period:")[1].strip()
                break
        if reporting_period_suffix is not None:
            break

    # 2. Find "Seattle University" row indices (block starts)
    su_indices: List[int] = []
    for idx, row in df.iterrows():
        if any(isinstance(v, str) and v.strip() == "Seattle University" for v in row):
            su_indices.append(idx)

    # 3. Find all "Page x of y" footer rows (block ends)
    page_rows: List[int] = []
    for idx, row in df.iterrows():
        for v in row:
            if isinstance(v, str) and re.search(r"Page\s+\d+\s+of\s+\d+", v):
                page_rows.append(idx)
                break

    # helper to find the page-row for a given SU block
    def find_page_for_block(su_idx: int):
        for pr in page_rows:
            if pr > su_idx:
                return pr
        return None

    records = []

    # 4. For each mini-table
    for su_idx in su_indices:
        page_idx = find_page_for_block(su_idx)
        if page_idx is None:
            continue

        # Layout offsets
        building_row = su_idx + 3
        location_row = su_idx + 4
        first_room_row = su_idx + 6

        if location_row >= len(df):
            continue

        # Building name
        building_val = df.loc[building_row, 0] if building_row < len(df) else None
        building_name = (
            building_val.strip()
            if isinstance(building_val, str) and building_val.strip()
            else (str(building_val).strip() if building_val is not None else "")
        )

        # ===== FIXED HEADER PARSING BLOCK =====
        header_row = df.loc[location_row]
        time_cols = {}

        for col_idx, val in header_row.items():
            if not isinstance(val, str):
                continue

            # ---- ALWAYS skip Column S ----
            if col_idx == 18:  # Column S
                continue

            vs = val.strip()
            vsl = vs.lower()

            # Skip Location or blank
            if vsl in ("location", "", None):
                continue

            # Reject EMS summary columns containing "avg" or extra "average"
            if ("average" in vsl or "avg" in vsl) and vsl != "average":
                continue

            # Keep EXACT "Average"
            if vsl == "average":
                time_cols[col_idx] = "Average"
                continue

            # Keep ONLY true hour labels: 6a–8p
            if re.fullmatch(r"\d{1,2}[ap]", vsl):
                time_cols[col_idx] = vs
                continue

            # Everything else ignored
        # ===== END FIXED HEADER PARSING BLOCK =====

        # 5. Process room rows
        r = first_room_row
        while r < page_idx:
            room_val = df.loc[r, 0]

            if not isinstance(room_val, str) or not room_val.strip():
                r += 1
                continue

            room_full = room_val.strip()

            if room_full == "Total":
                break

            # Skip noise rows
            if any(
                marker in room_full
                for marker in (
                    "Seattle University",
                    "Reporting Period",
                    "All figures",
                    "Page ",
                    "Grand Total",
                )
            ):
                r += 1
                continue

            # Extract values AS-IS
            for col_idx, hour_label in time_cols.items():
                raw_val = df.loc[r, col_idx]

                rec = {
                    "Building": building_name,
                    "Room": room_full,
                    "Hour": hour_label,
                    "Value": raw_val,
                }
                if reporting_period_suffix:
                    rec["Reporting Period"] = reporting_period_suffix

                records.append(rec)

            r += 2  # skip spacer row

    out = pd.DataFrame(records)

    # Order columns
    if not out.empty:
        cols = ["Building", "Room", "Hour", "Value"]
        if "Reporting Period" in out.columns:
            cols.append("Reporting Period")
        out = out[cols]

    return out