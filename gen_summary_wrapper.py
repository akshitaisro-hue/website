from .gen_summary_core import run as _run_gen_summary


def run_gen_summary(master_summary_path, output_folder, p_id, tc_alias_series=None, tm_alias_series=None):
    return _run_gen_summary(master_summary_path, output_folder, p_id, tc_alias_series, tm_alias_series)
