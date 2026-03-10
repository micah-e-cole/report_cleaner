import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import pytest
from cleaner.dispatcher import run_cleaner


def test_dispatcher_rejects_invalid_report_type(tmp_path):
    dummy_in = tmp_path / "input.xlsx"
    dummy_out = tmp_path / "output.xlsx"
    dummy_in.write_text("dummy")

    with pytest.raises(ValueError):
        run_cleaner(str(dummy_in), str(dummy_out), report_type="NotARealType")