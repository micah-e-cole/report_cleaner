from datetime import datetime
import re

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


def split_room_fields(building_name: str, room_full: str):
    """
    Split 'Bannan 222 - Classroom' into:
      Building, Room Number, Classroom Type
    """
    room_number = room_full.strip()
    room_type = ""

    if " - " in room_full:
        left, right = room_full.split(" - ", 1)
        room_type = right.strip()

        tokens = left.split()
        if tokens:
            last = tokens[-1]
            if any(ch.isdigit() for ch in last):
                room_number = last
            else:
                room_number = left.strip()
    return building_name, room_number, room_type


def format_hour_label(label: str) -> str:
    """Convert '6a' -> '6 AM', '12p' -> '12 PM'."""
    m = re.match(r"^(\d+)([ap])$", str(label).strip().lower())
    if not m:
        return label
    h = int(m.group(1))
    ap = m.group(2)
    return f"{h} {'AM' if ap == 'a' else 'PM'}"


def write_hourly_excel(df_long: pd.DataFrame, output_path: str) -> None:
    """
    Take the long-format hourly DataFrame from transform_hourly_utilization
    and write it to an Excel sheet with:

      Row 1: Title 'Seattle University, Reporting Period: ...'
      Row 2: Note 'All figures are percentages (Columns A:U).'
      Row 3: Headers
      Row 4+: Data

    Columns:
      A: Building
      B: Room Number
      C: Classroom Type
      D:R: Hour columns (6 AM:8 PM)
      S: Average Utilization by Room (row-wise average across D:R)
    """
    # 1. Extract reporting period once; remove from data table
    reporting_period = None
    if "Reporting Period" in df_long.columns and not df_long["Reporting Period"].isna().all():
        reporting_period = df_long["Reporting Period"].dropna().iloc[0]
        df_long = df_long.drop(columns=["Reporting Period"])

    # 2. Derive Building, Room Number, Classroom Type
    df = df_long.copy()
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

    # 3. Pivot to wide format: one row per room, columns per Hour
    if df.empty:
        wide = pd.DataFrame(columns=["Building", "Room Number", "Classroom Type"])
    else:
        # Keep raw Value entries, do NOT scale for percentages
        wide = df.pivot_table(
            index=["Building", "Room Number", "Classroom Type"],
            columns="Hour",
            values="Value",
            aggfunc=lambda x: x.iloc[0],  # keep first occurrence as-is
        )
        wide = wide.reset_index()

        # Flatten possible MultiIndex columns from pivot
        wide.columns = [col if isinstance(col, str) else col[1] for col in wide.columns]

        # 🔴 FIX: drop any 'Average' column coming from EMS (pivot),
        #        we'll compute our own Average later.
        if "Average" in wide.columns:
            wide = wide.drop(columns=["Average"])

        # Force expected hour columns to exist even if empty
        expected_hours = [
            "6a", "7a", "8a", "9a", "10a", "11a",
            "12p", "1p", "2p", "3p", "4p", "5p",
            "6p", "7p", "8p",
        ]
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
            h = int(m.group(1))
            ap = m.group(2)
            if ap == "a":
                h24 = h if h != 12 else 0
            else:
                h24 = h + 12 if h != 12 else 12
            return h24

        base_cols = ["Building", "Room Number", "Classroom Type"]
        hour_cols = [c for c in wide.columns if c not in base_cols]
        hour_cols_sorted = sorted(hour_cols, key=hour_key)

        # Ensure hour columns are numeric; empty → 0
        for h in hour_cols_sorted:
            wide[h] = pd.to_numeric(wide[h], errors="coerce").fillna(0)

        # Add Average column (will be the ONLY 'Average' column now)
        wide["Average"] = 0.0

        # Reorder columns: A–C base, D–R hours, S average
        wide = wide[base_cols + hour_cols_sorted + ["Average"]]

        # Rename hour labels for display (6a → 6 AM, etc.)
        rename_map = {h: format_hour_label(h) for h in hour_cols_sorted}
        wide = wide.rename(columns=rename_map)
        # Average column header stays "Average"

    sheet_name = "Hourly Utilization"

    # 4. Write wide DataFrame starting at row 3 (startrow=2 => header at row3)
    wide.to_excel(output_path, index=False, sheet_name=sheet_name, startrow=2)

    # 5. Apply Excel formatting and formulas
    wb = load_workbook(output_path)
    ws = wb[sheet_name]

    max_col = ws.max_column
    end_col = max_col

    # Title row (row 1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    ws["A1"] = (
        f"Seattle University, Reporting Period: {reporting_period}"
        if reporting_period
        else "Seattle University, Hourly Room Utilization"
    )
    ws["A1"].font = Font(bold=True)

    # Note row (row 2)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_col)
    ws["A2"] = "All figures are percentages (Columns A:U)."

    # Header row (row 3) bold
    for cell in ws[3]:
        cell.font = Font(bold=True)

    # Freeze panes below header
    ws.freeze_panes = "A4"

    # Auto-filter
    ws.auto_filter.ref = ws.dimensions

    # Auto-ish column widths based on content in rows 3+
    for col_idx, col_cells in enumerate(
        ws.iter_cols(min_row=3, max_row=ws.max_row, min_col=1, max_col=ws.max_column),
        start=1,
    ):
        max_len = 0
        for cell in col_cells:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 2

    # Set up Average in last column (row-wise, across hour columns)
    first_hour_col_idx = 4   # D
    last_hour_col_idx = max_col - 1  # last hour col (before Average)
    avg_col_idx = max_col          # Average column

    first_hour_letter = get_column_letter(first_hour_col_idx)
    last_hour_letter = get_column_letter(last_hour_col_idx)
    avg_letter = get_column_letter(avg_col_idx)

    # Data rows start at row 4 (since row 3 is the header)
    for row_idx in range(4, ws.max_row + 1):
        ws[f"{avg_letter}{row_idx}"] = (
            f"=AVERAGE({first_hour_letter}{row_idx}:{last_hour_letter}{row_idx})"
        )

    # Format hour and Average columns as standard numeric (no percentage scaling)
    for col_idx in range(first_hour_col_idx, avg_col_idx + 1):
        for cell in ws.iter_rows(
            min_row=4, max_row=ws.max_row, min_col=col_idx, max_col=col_idx
        ):
            cell[0].number_format = "0.0"

    wb.save(output_path)