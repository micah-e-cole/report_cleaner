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


def _is_page_header(text: str) -> bool:
    """True if col0 text is page/sub-table boilerplate."""
    if not text:
        return False

    if re.search(r"\d{1,2}/\d{1,2}/\d{4}", text):
        return True  # date line

    if text == "Seattle University":
        return True

    if text.startswith("Hourly Room Utilization"):
        return True

    if text.startswith("Reporting Period:"):
        return True

    if text.startswith("All figures"):
        return True

    if re.search(r"Page\s+\d+\s+of\s+\d+", text):
        return True

    return False


def transform_hourly_utilization(input_path: str) -> pd.DataFrame:
    """
    Transform an EMS Hourly Room Utilization export into long-format DataFrame.

    Rules:
      - Row i:   <Building Name>
      - Row i+1: 'Location' + hour labels → header
      - Rows after header, up to 'Total' → room rows (possibly across pages)
      - Buildings may span multiple sub-tables/pages; page headers are ignored.
      - Hour labels kept as-is (e.g. '7a', '12p', '8p', '9p').
      - Rooms with no hour data at all are still included.
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

    records: list[dict] = []

    n = len(df)
    i = 0

    current_building: str | None = None
    time_cols: dict[int, str] | None = None
    in_building = False

    while i < n:
        val0 = df.loc[i, 0]
        text0 = val0.strip() if isinstance(val0, str) else None

        # ------------------------------------------------------------------
        # A. Detect building header: row i has building name, and row i+1 has 'Location'
        # ------------------------------------------------------------------
        if isinstance(text0, str) and not _is_page_header(text0):
            # Check if next row is 'Location'
            if i + 1 < n:
                next_val0 = df.loc[i + 1, 0]
                next_text0 = next_val0.strip() if isinstance(next_val0, str) else None

                if next_text0 == "Location":
                    # Start a new building block
                    current_building = text0
                    in_building = True

                    # Parse hour columns from header row (row i+1)
                    header_idx = i + 1
                    header_row = df.loc[header_idx]
                    time_cols = {}

                    for col_idx, v in header_row.items():
                        if not isinstance(v, str):
                            continue

                        vs = v.strip()
                        vsl = vs.lower()

                        if vsl in ("location", "", None):
                            continue

                        if ("average" in vsl or "avg" in vsl) and vsl != "average":
                            continue

                        if vsl == "average":
                            time_cols[col_idx] = "Average"
                            continue

                        vs_clean = re.sub(r"\s+", "", vsl)
                        if re.fullmatch(r"\d{1,2}[ap]", vs_clean):
                            time_cols[col_idx] = vs_clean
                            continue

                    # Move i to first potential room row (after header)
                    i = header_idx + 1
                    continue  # go to next iteration using new i

        # ------------------------------------------------------------------
        # If not in a building yet, just advance
        # ------------------------------------------------------------------
        if not in_building or time_cols is None or current_building is None:
            i += 1
            continue

        # Now we are inside a building block (between its header and 'Total')

        val0 = df.loc[i, 0]
        text0 = val0.strip() if isinstance(val0, str) else None

        # ------------------------------------------------------------------
        # B. End of building: 'Total'
        # ------------------------------------------------------------------
        if isinstance(text0, str) and text0 == "Total":
            in_building = False
            time_cols = None
            current_building = None
            i += 1
            continue

        # ------------------------------------------------------------------
        # C. Page headers inside building: ignore
        # ------------------------------------------------------------------
        if isinstance(text0, str) and _is_page_header(text0):
            i += 1
            continue

        # ------------------------------------------------------------------
        # D. Skip summary 'Average' rows inside building
        # ------------------------------------------------------------------
        if isinstance(text0, str) and text0 == "Average":
            i += 1
            continue

        # Skip blank/non-string rows
        if not isinstance(val0, str) or not text0:
            i += 1
            continue

        # ------------------------------------------------------------------
        # E. Treat as room row
        # ------------------------------------------------------------------
        room_full = text0

        for col_idx, hour_label in time_cols.items():
            if hour_label == "Average":
                continue

            raw_val = df.loc[i, col_idx]

            rec = {
                "Building": current_building,
                "Room": room_full,
                "Hour": hour_label,
                "Value": raw_val,
            }
            if reporting_period_suffix:
                rec["Reporting Period"] = reporting_period_suffix

            records.append(rec)

        i += 1  # next row

    out = pd.DataFrame(records)

    if not out.empty:
        cols = ["Building", "Room", "Hour", "Value"]
        if "Reporting Period" in out.columns:
            cols.append("Reporting Period")
        out = out[cols]

    return out