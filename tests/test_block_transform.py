# tests/test_block_transform.py

import sys
import os
from pathlib import Path

import pandas as pd

# Ensure project root is on the path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from cleaner.block.transform import transform_classroom_utilization


def test_block_transform_from_fixture():
    # Use the real fixture name here:
    fixture_path = Path(__file__).parent / "fixtures" / "Block_Export_Test.xls"
    # If your file is actually Block_Export_Test.xls, change the line above.

    result = transform_classroom_utilization(str(fixture_path))

    # --- Basic shape checks ---
    assert isinstance(result, pd.DataFrame)
    assert not result.empty

    expected_columns = {
        "Building",
        "Room",
        "Class Meetings",
        "Class Hours",
        "Utilization %",
        "Avg Est Enroll",
        "Avg Act Enroll",
        "Max Capacity",
        "Seat Fill %",
    }

    # Check that all expected columns are present
    assert expected_columns.issubset(result.columns), f"Got columns: {result.columns}"

    # --- Content sanity checks ---

    # Building and Room should not be entirely null
    assert result["Building"].notna().any(), "All Building values are NaN"
    assert result["Room"].notna().any(), "All Room values are NaN"

    # Numeric columns should have at least some numeric values
    numeric_cols = [
        "Class Meetings",
        "Class Hours",
        "Avg Est Enroll",
        "Avg Act Enroll",
        "Max Capacity",
    ]
    for col in numeric_cols:
        # At least one non-null value
        assert result[col].notna().any(), f"All values in {col} are NaN"

    # Utilization % and Seat Fill % should be floats between 0 and 1 (when not null)
    for col in ["Utilization %", "Seat Fill %"]:
        not_null = result[col].notna()
        if not_null.any():
            assert result.loc[not_null, col].between(0.0, 1.0).all(), (
                f"{col} values out of [0, 1] range"
            )