# tests/test_gui.py

import os
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from excel_cleaner.cleaner.gui import create_app


@pytest.fixture
def app(monkeypatch, tmp_path):
    """
    Build the GUI in memory and patch out messageboxes so tests don't hang.
    Returns (root, ctx).
    """
    # Patch messagebox functions to no-op / controlled responses
    from excel_cleaner.cleaner import gui as gui_mod

    monkeypatch.setattr(gui_mod.messagebox, "showinfo", lambda *a, **k: None)
    monkeypatch.setattr(gui_mod.messagebox, "showerror", lambda *a, **k: None)
    # When asked "Process another?", always say "No" to avoid loops
    monkeypatch.setattr(gui_mod.messagebox, "askyesno", lambda *a, **k: False)

    root, ctx = create_app()

    # Make sure we clean up the Tk root after each test
    yield root, ctx
    try:
        root.destroy()
    except Exception:
        pass


def test_single_mode_calls_run_cleaner(monkeypatch, tmp_path, app):
    root, ctx = app
    mode_var = ctx["mode_var"]
    input_var = ctx["input_var"]
    output_var = ctx["output_var"]
    run_process = ctx["run_process"]

    # Create a real temporary input file, because run_process checks os.path.isfile
    input_file = tmp_path / "sample.csv"
    input_file.write_text("dummy data")

    output_file = tmp_path / "out.xlsx"

    mode_var.set("single")
    input_var.set(str(input_file))
    output_var.set(str(output_file))

    # Patch run_cleaner to a MagicMock to capture calls
    from excel_cleaner.cleaner import gui as gui_mod

    mock_run_cleaner = MagicMock()
    monkeypatch.setattr(gui_mod, "run_cleaner", mock_run_cleaner)

    # Act
    run_process()

    # Assert run_cleaner called once with the paths we set
    mock_run_cleaner.assert_called_once_with(str(input_file), str(output_file))


def test_batch_mode_calls_run_batch_cleaner_with_folder(monkeypatch, tmp_path, app):
    root, ctx = app
    mode_var = ctx["mode_var"]
    input_var = ctx["input_var"]
    output_var = ctx["output_var"]
    run_process = ctx["run_process"]

    # Simulate batch mode using a directory as input
    input_dir = tmp_path / "input_dir"
    input_dir.mkdir()
    # Create dummy files - run_batch_cleaner is mocked so content doesn't matter
    (input_dir / "a.csv").write_text("x")
    (input_dir / "b.xlsx").write_text("x")

    output_dir = tmp_path / "output_dir"

    mode_var.set("batch")
    input_var.set(str(input_dir))
    output_var.set(str(output_dir))

    from excel_cleaner.cleaner import gui as gui_mod

    mock_run_batch = MagicMock()
    monkeypatch.setattr(gui_mod, "run_batch_cleaner", mock_run_batch)

    # Act
    run_process()

    # Assert that batch cleaner was called with [input_dir] and output_dir
    mock_run_batch.assert_called_once()
    args, kwargs = mock_run_batch.call_args
    inputs_arg, out_dir_arg = args
    assert inputs_arg == [str(input_dir)]
    assert out_dir_arg == str(output_dir)


def test_run_process_shows_error_when_missing_input(monkeypatch, app):
    root, ctx = app
    mode_var = ctx["mode_var"]
    input_var = ctx["input_var"]
    output_var = ctx["output_var"]
    run_process = ctx["run_process"]

    mode_var.set("single")
    input_var.set("")  # Missing input
    output_var.set("C:/some/output.xlsx")

    from excel_cleaner.cleaner import gui as gui_mod

    mock_showerror = MagicMock()
    monkeypatch.setattr(gui_mod.messagebox, "showerror", mock_showerror)

    run_process()

    mock_showerror.assert_called_once()
    # First arg should be the title "Error"
    args, kwargs = mock_showerror.call_args
    assert "Error" in args[0]