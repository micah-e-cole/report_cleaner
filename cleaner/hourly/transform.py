# cleaner/hourly/transform.py
# ------------- ABOUT -------------
# Author: Micah Braun
# Purpose: provides the logic for cleaning hourly-style exported files from EMS
#          to pandas dataframes. Plugin-compatible version.

import re
from typing import List
import pandas as pd


def transform(df_raw: pd.DataFrame, **options) -> pd.DataFrame:
    """
    Transform raw EMS Hourly Room Utilization data (loaded earlier by
    common.read_raw_table) into long-format DataFrame.

    IMPORTANT:
      - Hour labels are kept EXACTLY as they appear in EMS.
      - No numeric normalization.
      - No reading from disk; pure transform only.

    Output columns:
        Building
        Room
        Hour
        Value
        Reporting Period (optional)
    """

    # Work on a copy so caller's DataFrame isn't modified
    df = df_raw.copy()

    # ---------------------------------------------------------
    # 1. Extract Reporting Period suffix once
    # ---------------------------------------------------------
    reporting_period_suffix = None
    for _, row in df.iterrows():
        for val in row:
            if isinstance(val, str) and val.startswith("Reporting Period:"):
                reporting_period_suffix = val.split("Reporting Period:")[1].strip()
                break
        if reporting_period_suffix:
            break

    # ---------------------------------------------------------
    # 2. Find "Seattle University" block starting rows
    # ---------------------------------------------------------
    su_indices: List[int] = []
    for idx, row in df.iterrows():
        if any(isinstance(v, str) and v.strip() == "Seattle University" for v in row):
            su_indices.append(idx)

    # ---------------------------------------------------------
    # 3. Find page footer rows ("Page x of y")
    # ---------------------------------------------------------
    page_rows: List[int] = []
    for idx, row in df.iterrows():
        for v in row:
            if isinstance(v, str) and re.search(r"Page\s+\d+\s+of\s+\d+", v):
                page_rows.append(idx)
                break

    def find_page_for_block(su_idx: int):
        for pr in page_rows:
            if pr > su_idx:
                return pr
        return None

        val0 = df.loc[i, 0]
        text0 = val0.strip() if isinstance(val0, str) else None

    # ---------------------------------------------------------
    # 4. Process each mini-table
    # ---------------------------------------------------------
    for su_idx in su_indices:
        page_idx = find_page_for_block(su_idx)
        if page_idx is None:
            continue

        # Layout offsets (fixed EMS format)
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

        # ---------------------------------------------------------
        # HEADER PARSING BLOCK (extract hour columns)
        # ---------------------------------------------------------
        header_row = df.loc[location_row]
        time_cols = {}  # col_idx → hour/label ("6a", "Average", etc.)

        for col_idx, val in header_row.items():
            if not isinstance(val, str):
                continue

            # Always skip EMS Column S (index 18)
            if col_idx == 18:
                continue

            vs = val.strip()
            vsl = vs.lower()

            # Skip blank/location columns
            if vsl in ("location", "", None):
                continue

            # Reject EMS summary columns except the *exact* word “Average”
            if ("average" in vsl or "avg" in vsl) and vsl != "average":
                continue

            if vsl == "average":
                time_cols[col_idx] = "Average"
                continue

            # Keep only real hour labels
            if re.fullmatch(r"\d{1,2}[ap]", vsl):
                time_cols[col_idx] = vs
                continue

        # ---------------------------------------------------------
        # 5. Process room rows
        # ---------------------------------------------------------
        r = first_room_row
        while r < page_idx:
            room_val = df.loc[r, 0]

            if not isinstance(room_val, str) or not room_val.strip():
                r += 1
                continue

        # Skip blank/non-string rows
        if not isinstance(val0, str) or not text0:
            i += 1
            continue

            # End of block
            if room_full == "Total":
                break

        for col_idx, hour_label in time_cols.items():
            if hour_label == "Average":
                continue

            # Extract values exactly as EMS lists them
            for col_idx, hour_label in time_cols.items():
                raw_val = df.loc[r, col_idx]

            rec = {
                "Building": current_building,
                "Room": room_full,
                "Hour": hour_label,
                "Value": raw_val,
            }
            if reporting_period_suffix:
                rec["Reporting Period"] = reporting_period_suffix

            records.append(rec)

            r += 2  # EMS spacer row

    # ---------------------------------------------------------
    # 6. Final DataFrame assembly
    # ---------------------------------------------------------
    out = pd.DataFrame(records)

    if not out.empty:
        cols = ["Building", "Room", "Hour", "Value"]
        if "Reporting Period" in out.columns:
            cols.append("Reporting Period")
        out = out[cols]

    return out