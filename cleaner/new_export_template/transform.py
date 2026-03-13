# cleaner/new_export_template/transform.py

import pandas as pd


def transform(df_raw: pd.DataFrame, **options) -> pd.DataFrame:
    """
    Transform raw input into the cleaned structure your report type requires.

    This file is intentionally minimal — copy/paste from a real transform
    (block or hourly) when creating a real export type.

    Parameters:
        df_raw: DataFrame loaded by read_new_type_raw()
        options: extra parameters from GUI or CLI

    Returns:
        pd.DataFrame: cleaned/processed data ready for writing.
    """

    # TODO: replace with real transformation
    df = df_raw.copy()

    # Example cleanup
    df.columns = [str(col).strip() for col in df.columns]

    return df