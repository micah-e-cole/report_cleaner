# cleaner/common.py

from pathlib import Path
import pandas as pd
import os
import re


# ---------------------------------------------------------
# File Validation + Collection
# ---------------------------------------------------------

ALLOWED_EXTENSIONS = {'.csv', '.xls', '.xlsx'}


def is_valid_input_file(path: Path) -> bool:
    """
    Return True if the path exists, is a file, and has a valid extension.
    """
    return path.is_file() and path.suffix.lower() in ALLOWED_EXTENSIONS


def collect_input_files(paths: list[str]) -> list[Path]:
    """
    Given a mix of file paths and directory paths, return a list of valid
    input files for supported report types.
    """
    collected: list[Path] = []

    for p_str in paths:
        p = Path(p_str)

        if not p.exists():
            print(f"⚠️  Skipping '{p}': path does not exist.")
            continue

        if p.is_dir():
            for file in p.iterdir():
                if is_valid_input_file(file):
                    collected.append(file)
                elif file.is_file():
                    print(
                        f"⚠️  Skipping '{file.name}': "
                        f"'.{file.suffix.lstrip('.')}' is not an accepted file type. "
                        f"Accepted: {', '.join(ALLOWED_EXTENSIONS)}"
                    )
        else:
            if is_valid_input_file(p):
                collected.append(p)
            else:
                print(
                    f"⚠️  Skipping '{p.name}': "
                    f"'.{p.suffix.lstrip('.')}' is not an accepted file type. "
                    f"Accepted: {', '.join(ALLOWED_EXTENSIONS)}"
                )

    return collected


# ---------------------------------------------------------
# Generic Raw Loader
# ---------------------------------------------------------

def read_raw_table(input_path: str) -> pd.DataFrame:
    """
    Generic raw-table loader for EMS reports.
    Reads CSV, XLS, XLSX with no header.
    Drops completely blank rows.
    Used by all export types unless overridden.
    """
    ext = Path(input_path).suffix.lower()

    if ext == ".csv":
        df = pd.read_csv(input_path, header=None)
    elif ext == ".xlsx":
        df = pd.read_excel(input_path, header=None, engine="openpyxl")
    elif ext == ".xls":
        df = pd.read_excel(input_path, header=None, engine="xlrd")
    else:
        raise ValueError(f"Unsupported file type '{ext}'. "
                         f"Allowed: {', '.join(ALLOWED_EXTENSIONS)}")

    return df.dropna(how="all").reset_index(drop=True)


# ---------------------------------------------------------
# Shared Parsing Helpers Used by Block + Hourly + Future Types
# ---------------------------------------------------------

def split_room_fields(building_name: str, room_full: str):
    """
    Split 'Bannan 222 - Classroom' into:
      Building Name, Room Number, Classroom Type
    """
    room_number = room_full.strip()
    room_type = ""

    if " - " in room_full:
        left, right = room_full.split(" - ", 1)
        room_type = right.strip()

        tokens = left.split()
        last = tokens[-1] if tokens else ""

        if any(ch.isdigit() for ch in last):
            room_number = last
        else:
            room_number = left.strip()

    return building_name, room_number, room_type


def format_hour_label(label: str) -> str:
    """
    Convert '6a' → '6 AM', '12p' → '12 PM'.
    Used by hourly transforms.
    """
    m = re.match(r"^(\d+)([ap])$", str(label).strip().lower())
    if not m:
        return label

    h = int(m.group(1))
    ap = m.group(2)
    return f"{h} {'AM' if ap == 'a' else 'PM'}"


# ---------------------------------------------------------
# Template for future report-type loaders
# ---------------------------------------------------------

def read_new_type_raw(input_path: str) -> pd.DataFrame:
    """
    Placeholder for a future export type's loader.
    By default, just uses the generic loader but can be replaced.
    """
    return read_raw_table(input_path)