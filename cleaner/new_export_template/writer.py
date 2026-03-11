# cleaner/new_export_template/writer.py

import pandas as pd
from openpyxl import Workbook


def write(df_processed: pd.DataFrame, output_path: str, **options) -> None:
    """
    Write the processed data to an Excel file.
    This is intentionally minimal — customize as needed.
    """

    wb = Workbook()
    ws = wb.active
    ws.title = "New Export"

    # Header row
    ws.append(list(df_processed.columns))

    # Data rows
    for _, row in df_processed.iterrows():
        ws.append(list(row.values))

    wb.save(output_path)