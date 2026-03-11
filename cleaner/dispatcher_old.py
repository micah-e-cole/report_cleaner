# cleaner/dispatcher.py
# ---------------------------------------------------------
# Central routing logic for selecting the correct transform
# and write functions for each report type.
# ---------------------------------------------------------

# Block report imports
from .block.transform import transform_classroom_utilization
from .block.writer import write_formatted_excel

# Hourly report imports
from .hourly.transform import transform_hourly_utilization
from .hourly.writer import write_hourly_excel


# ---------------------------------------------------------
# REGISTER REPORT TYPES HERE
# ---------------------------------------------------------
REPORT_TYPES = {
    "Block Room Utilization": {
        "transform": transform_classroom_utilization,
        "write": write_formatted_excel,
    },
    "Hourly Room Utilization": {
        "transform": transform_hourly_utilization,
        "write": write_hourly_excel,
    },
}


# ---------------------------------------------------------
# Public API — used by GUI and CLI
# ---------------------------------------------------------
def run_cleaner(input_path: str, output_path: str, report_type: str):
    """
    Select the appropriate cleaning workflow based on report type.
    """

    if report_type not in REPORT_TYPES:
        raise ValueError(
            f"Unknown report type '{report_type}'. "
            f"Available: {', '.join(REPORT_TYPES.keys())}"
        )

    funcs = REPORT_TYPES[report_type]
    transform_fn = funcs["transform"]
    write_fn = funcs["write"]

    df = transform_fn(input_path)
    write_fn(df, output_path)


def run_batch_cleaner(inputs: list[str], output_dir: str, report_type: str, collect_files_fn):
    """
    Batch cleaning across multiple input paths.
    collect_files_fn comes from common.py.
    """

    from pathlib import Path

    files_to_process = collect_files_fn(inputs)

    if not files_to_process:
        print("No valid input files found. Exiting.")
        return

    if output_dir is not None:
        out_base = Path(output_dir)
        out_base.mkdir(parents=True, exist_ok=True)
    else:
        out_base = None

    print(f"Found {len(files_to_process)} file(s) to process.")

    for f in files_to_process:
        try:
            target_dir = f.parent if out_base is None else out_base
            output_path = target_dir / f"{f.stem}_cleaned.xlsx"

            print(f"🔄 Processing: {f}")
            run_cleaner(str(f), str(output_path), report_type)
            print(f"✅ Finished: {f.name} → {output_path}")

        except Exception as e:
            print(f"❌ Error processing '{f}': {e}")

    print("🎉 Batch processing complete.")