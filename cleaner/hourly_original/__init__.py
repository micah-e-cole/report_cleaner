from ..templates import ExportTemplate, register_export_template
from ..common import read_raw_table
from .transform import transform
from .writer import write


@register_export_template("Hourly Room Utilization")
class HourlyExport(ExportTemplate):

    def run(self, input_path: str, output_path: str, **options):
        df_raw = read_raw_table(input_path)
        df_long = transform(df_raw, **options)
        write(df_long, output_path, **options)