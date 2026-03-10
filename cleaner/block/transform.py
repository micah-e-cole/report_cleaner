import os
import pandas as pd
import xlrd
from openpyxl import load_workbook

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

    # Drop blank rows
    df2 = df.dropna(how='all').copy()

    # String builder
    def row_to_str(row):
        return ' '.join([str(x) for x in row if pd.notna(x)])

    row_str = df2.apply(row_to_str, axis=1)

    noise_keywords = [
        'Seattle University',
        'Reporting Period',
        'Classroom Utilization',
        'Class Meetings',
        'Avg. Est.  Enroll',
        'Avg. Est. Enroll',
        'Avg. Act.  Enroll',
        'Avg. Act. Enroll',
        'Seat Fill',
        'Page ',
    ]

    base_noise = row_str.str.contains('|'.join(noise_keywords))
    mask_date = row_str.str.contains(r'\d{1,2}/\d{1,2}/\d{4} \d{1,2}:\d{2} [AP]M')

    clean = df2[~(base_noise | mask_date)].copy()

    # Forward-fill buildings
    clean['Building'] = None
    current_building = None
    for idx, row in clean.iterrows():
        bname = row[0]
        room = row[1]
        if pd.notna(bname) and pd.isna(room) and not str(bname).startswith(('Total for', 'Average for')):
            current_building = bname
        clean.at[idx, 'Building'] = current_building

    # Filter room rows
    room_rows = clean[clean[1].notna() & clean[4].notna() & clean[6].notna()].copy()

    # Build output table
    out = pd.DataFrame({
        'Building':       room_rows['Building'].values,
        'Room':           room_rows[1].values,
        'Class Meetings': room_rows[4].values,
        'Class Hours':    room_rows[6].values,
        'Utilization %':  room_rows[7].values,
        'Avg Est Enroll': room_rows[8].values,
        'Avg Act Enroll': room_rows[9].values,
        'Max Capacity':   room_rows[11].values,
        'Seat Fill %':    room_rows[12].values,
    })

    # Numeric fixes
    for col in ['Class Meetings', 'Class Hours', 'Avg Est Enroll', 'Avg Act Enroll', 'Max Capacity']:
        out[col] = pd.to_numeric(out[col], errors='coerce')

    # Percent fixes
    for col in ['Utilization %', 'Seat Fill %']:
        col_str = out[col].astype(str)
        mask_pct = col_str.str.contains('%', na=False)

        numeric = pd.Series(index=out.index, dtype='float64')
        numeric[mask_pct] = (
            col_str[mask_pct]
            .str.replace('%', '', regex=False)
            .astype(float) / 100.0
        )
        numeric[~mask_pct] = pd.to_numeric(out.loc[~mask_pct, col], errors='coerce')
        numeric = numeric.where((numeric <= 1) | numeric.isna(), numeric / 100.0)

        out[col] = numeric

    return out
