# cleaner/new_type/__init__.py

from ..templates import ExportTemplate, register_export_template
from ..common import read_new_type_raw   # or read_raw_table if generic
from .transform import transform
from .writer import write


@register_export_template("New Report Type")
class NewTypeExport(ExportTemplate):
    """
    Template plugin for a future report type.
    """

    def run(self, input_path: str, output_path: str, **options):
        # Load raw data (CSV/XLS/XLSX)
        df_raw = read_new_type_raw(input_path)

        # Transform into cleaned DataFrame
        df_processed = transform(df_raw, **options)

        # Write out to Excel (or other format your plugin chooses)
        write(df_processed, output_path, **options)
