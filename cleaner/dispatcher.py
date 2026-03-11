# cleaner/dispatcher.py

# Import plugin registry
from .templates import get_export_template, list_export_templates

# Import plugin packages so their decorators execute
from . import hourly   # noqa: F401
from . import block    # noqa: F401
from . import new_type # noqa: F401   <-- optional / future

from pathlib import Path


def run_cleaner(input_path: str, output_path: str, report_type: str, **options):
    """
    Run a plugin-based export by report type name.
    """

    TemplateClass = get_export_template(report_type)
    template = TemplateClass()
    template.run(input_path, output_path, **options)


def run_batch_cleaner(inputs: list[str], output_dir: str, report_type: str, collect_files_fn, **options):
    """
    Batch cleaning across multiple input paths.
    """
    files_to_process = collect_files_fn(inputs)

    if not files_to_process:
        print("No valid input files found. Exiting.")
        return

    if output_dir:
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
            run_cleaner(str(f), str(output_path), report_type, **options)
            print(f"✅ Finished: {f.name} → {output_path}")

        except Exception as e:
            print(f"❌ Error processing '{f}': {e}")

    print("🎉 Batch processing complete.")
