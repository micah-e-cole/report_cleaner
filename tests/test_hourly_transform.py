# tests/test_hourly_transform.py

import sys
import os
from pathlib import Path
import pandas as pd

# Ensure project root is importable
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from cleaner.hourly.transform import transform_hourly_utilization


def test_hourly_transform_from_fixture():
    fixture_path = Path(__file__).parent / "fixtures" / "Hourly_Export_Test.xls"

    result = transform_hourly_utilization(str(fixture_path))

    # Basic DataFrame checks
    assert isinstance(result, pd.DataFrame)
    assert not result.empty

    # Expected structure from transform_hourly_utilization
    expected_cols = {"Building", "Room", "Hour", "Value"}
    assert expected_cols.issubset(result.columns)

    # Reporting Period may or may not be present depending on file
    if "Reporting Period" in result.columns:
        assert result["Reporting Period"].notna().any()

    # Building / Room must have non-null values
    assert result["Building"].notna().any()
    assert result["Room"].notna().any()

    # Hour labels must match EMS hour format (6a, 12p, etc.)
    assert result["Hour"].astype(str).str.match(r"^\d{1,2}[ap]$").any()

    # Value column is raw (no normalization), so allow numeric or string
    assert result["Value"].notna().any()


def test_hourly_transform_known_structure_row():
    """
    Optional: checks that at least one row contains all expected elements.
    Not tied to specific room names, so it works across years/terms.
    """
    fixture_path = Path(__file__).parent / "fixtures" / "Hourly_Export_Test.xls"
    result = transform_hourly_utilization(str(fixture_path))

    # Pick the first row as a sanity reference
    row = result.iloc[0]

    assert isinstance(row["Building"], str)
    assert isinstance(row["Room"], str)
    assert isinstance(row["Hour"], str)
    assert row["Hour"].endswith(("a", "p"))
    assert row["Value"] is not None