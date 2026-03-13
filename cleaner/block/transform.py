# excel_cleaner/cleaner/block/transform.py
# ------------- ABOUT -------------
# Author: Micah Braun
# Purpose: provides the logic for cleaning
# block-style exported files from EMS

import os
import re
import pandas as pd
import xlrd
from openpyxl import load_workbook


def extract_room_number(room: str):
    """
    Extracts the room number from any room string by finding the first digit
    and returning the substring starting from that digit.

    Examples:
        'Administration 202' -> '202'
        'Bannan 244 (E)' -> '244 (E)'
        'Xavier 030' -> '030'
        'Pigott 100/110 Combo' -> '100/110 Combo'
        'Lemieux Library 122 - Boeing Room (Classroom)' -> '122'
    """
    if pd.isna(room):
        return room

    s = str(room).strip()

    m = re.search(r'\d', s)
    if not m:
        return s  # no digit found; return original

    return s[m.start():].strip()


def split_room_and_type(value):
    """
    Splits raw room string into (room, room_type).

    Rules:
    1. If there is a '-', split on the first '-' and treat:
       left  -> room
       right -> room_type
    2. If there is no '-', find the last digit in the string.
       Everything up to and including that digit -> room
       Everything after that digit                -> room_type
    3. If no digit exists, return the whole string as room, room_type=None
    """
    if pd.isna(value):
        return None, None

    s = str(value).strip()

    # Case 1: "Administration 202 - Classroom"
    if "-" in s:
        left, right = s.split("-", 1)
        return left.strip(), right.strip()

    # Case 2: no '-', use last numeric char as the boundary
    # r'\d(?!.*\d)' = "match the last digit in the string"
    m = re.search(r"\d(?!.*\d)", s)
    if m:
        idx = m.end()  # position just after the last digit
        room = s[:idx].strip()
        room_type = s[idx:].strip()
        if room_type == "":
            room_type = None
        return room, room_type

    # Case 3: no '-' and no digit → we can't infer room type
    return s, None


def extract_and_remove_footer_datestamp(df: pd.DataFrame):
    """
    Locate the footer that looks like:

        [col 0]  date/time/initials (sometimes)
        [col 10-12] "Page x of y"  (merged across 2 rows visually)

    Strategy:
      1. Scan columns 10, 11, 12 (K,L,M) for 'Page x of x'.
      2. Take the LAST such row index -> footer_anchor_row.
      3. Look at footer_anchor_row - 1, footer_anchor_row, footer_anchor_row + 1
         in column 0 to find a value that matches a datetime-like pattern.
      4. Use that as the datestamp, and drop BOTH:
           - the anchor row with 'Page x of x'
           - the row where the datestamp was actually found (if different)

    Returns:
        df_cleaned, datestamp_text
    """
    page_cols = [10, 11, 12]  # Excel K, L, M (0-based indices)
    footer_anchor_row = None

    # 1. Find the last row that contains "Page x of x" in K/L/M
    for col in page_cols:
        if col in df.columns:
            col_str = df[col].astype(str)
            matches = col_str[col_str.str.contains(r"Page\s+\d+\s+of\s+\d+", na=False)]
            if not matches.empty:
                # Keep the last 'Page x of x' we see
                footer_anchor_row = matches.index[-1]

    if footer_anchor_row is None:
        # No footer detected
        return df, None

    # 2. Search around that row in column 0 (A) for the datestamp
    date_pattern = re.compile(
        r"\d{1,2}/\d{1,2}/\d{2,4}\s+\d{1,2}:\d{2}\s*[AP]M",
        re.IGNORECASE,
    )

    datestamp_row = None
    datestamp_value = None

    # Check [anchor - 1, anchor, anchor + 1] to be safe with the 2-row merge
    neighbor_rows = [
        r for r in [footer_anchor_row - 1, footer_anchor_row, footer_anchor_row + 1]
        if r in df.index
    ]

    for r in neighbor_rows:
        val = df.at[r, 0]
        if pd.notna(val):
            s = str(val)
            if date_pattern.search(s):
                datestamp_row = r
                datestamp_value = s.strip()
                break

    # 3. Build set of rows to drop
    rows_to_drop = {footer_anchor_row}
    if datestamp_row is not None:
        rows_to_drop.add(datestamp_row)

    df_cleaned = df.drop(index=list(rows_to_drop))

    return df_cleaned, datestamp_value


def transform_classroom_utilization(input_path: str) -> pd.DataFrame:

    ext = os.path.splitext(input_path)[1].lower()

    # 1. Choose correct reader
    if ext == ".csv":
        df = pd.read_csv(input_path, header=None)
    elif ext == ".xlsx":
        df = pd.read_excel(input_path, header=None, engine="openpyxl")
    elif ext == ".xls":
        df = pd.read_excel(input_path, header=None, engine="xlrd")
    else:
        raise ValueError(f"Unsupported file type: {ext}. Use .xls, .xlsx, or .csv.")

    # Drop completely blank rows
    df2 = df.dropna(how="all").copy()

    # --- Capture and REMOVE the footer row(s) (Page x of x + datestamp) ---
    df2, footer_datestamp = extract_and_remove_footer_datestamp(df2)

    # String builder for filtering noise
    def row_to_str(row):
        return " ".join([str(x) for x in row if pd.notna(x)])

    row_str = df2.apply(row_to_str, axis=1)

    noise_keywords = [
        "Seattle University",
        "Reporting Period",
        "Classroom Utilization",
        "Class Meetings",
        "Avg. Est.  Enroll",
        "Avg. Est. Enroll",
        "Avg. Act.  Enroll",
        "Avg. Act. Enroll",
        "Seat Fill",
        "Page ",  # other header/footer page text
    ]

    base_noise = row_str.str.contains("|".join(noise_keywords))
    mask_date = row_str.str.contains(r"\d{1,2}/\d{1,2}/\d{4} \d{1,2}:\d{2} [AP]M")

    # Remove headers / noise rows (footer datestamp row already removed explicitly)
    clean = df2[~(base_noise | mask_date)].copy()

    # Forward-fill buildings
    clean["Building"] = None
    current_building = None
    for idx, row in clean.iterrows():
        bname = row[0]
        room = row[1]
        if (
            pd.notna(bname)
            and pd.isna(room)
            and not str(bname).startswith(("Total for", "Average for"))
        ):
            current_building = bname
        clean.at[idx, "Building"] = current_building

    # Filter room rows
    room_rows = clean[clean[1].notna() & clean[4].notna() & clean[6].notna()].copy()

    # Parse Room and Room Type from Column B (raw col 1)
    room_type_df = room_rows[1].apply(
        lambda x: pd.Series(split_room_and_type(x), index=["Room", "Room Type"])
    )
    room_rows["Room"] = room_type_df["Room"]
    room_rows["Room Type"] = room_type_df["Room Type"]
    # Strip building name prefix from Room
    room_rows['Room'] = room_rows['Room'].apply(extract_room_number)
   
    # Build output table
    out = pd.DataFrame(
        {
            "Building": room_rows["Building"].values,
            "Room": room_rows["Room"].values,
            "Room Type": room_rows["Room Type"].values,
            "Class Meetings": room_rows[4].values,
            "Class Hours": room_rows[6].values,
            "Utilization %": room_rows[7].values,
            "Avg Est Enroll": room_rows[8].values,
            "Avg Act Enroll": room_rows[9].values,
            "Max Capacity": room_rows[11].values,
            "Seat Fill %": room_rows[12].values,
        }
    )

    # Numeric fixes
    for col in [
        "Class Meetings",
        "Class Hours",
        "Avg Est Enroll",
        "Avg Act Enroll",
        "Max Capacity",
    ]:
        out[col] = pd.to_numeric(out[col], errors="coerce")

    # Percent fixes
    for col in ["Utilization %", "Seat Fill %"]:
        col_str = out[col].astype(str)
        mask_pct = col_str.str.contains("%", na=False)

        numeric = pd.Series(index=out.index, dtype="float64")
        numeric[mask_pct] = (
            col_str[mask_pct].str.replace("%", "", regex=False).astype(float) / 100.0
        )
        numeric[~mask_pct] = pd.to_numeric(out.loc[~mask_pct, col], errors="coerce")

        # If something looks like 50 (rather than 0.5), assume it's 50% and divide by 100
        numeric = numeric.where((numeric <= 1) | numeric.isna(), numeric / 100.0)

        out[col] = numeric

    # --- Append footer datestamp at the end, top row, without touching Seat Fill % ---
    if footer_datestamp is not None:
        datestamp_col = "Report Generated"
        # Append as a new final column
        out[datestamp_col] = None
        # Place the datestamp in the first row of that column
        out.at[0, datestamp_col] = footer_datestamp

    return out