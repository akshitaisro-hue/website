import re
import pandas as pd
import numpy as np
from pathlib import Path

SHEET_TC = "master_summary_tc"
SHEET_TM = "master_summary_tm"

SRC_FUNC_DESC = "FUNCTIONAL_DESCRIPTION"
SRC_CARD_TYPE = "TC_CARD_TYPE"
SRC_IP_ADDRESS = "IP_ADDRESS"
SRC_ASSIGNED_PIN = "ASSIGNED_TC_PIN"

SRC_TM_FUNC_DESC = "TM_FUNCTIONAL_DESCRIPTION"
SRC_TM_IP_ADDRESS = "IP_ADDRESS"
SRC_TM_CARD_TYPE = "TM_CARD_TYPE"
SRC_TM_ASSIGNED_PIN = "ASSIGNED_TM_PIN"
SRC_TM_LENGTH = "TM_LENGTH"


def load_sheet(file_path: Path, sheet_name: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except ValueError as e:
        print(f"Sheet '{sheet_name}' not found in {file_path.name}: {e}")
        xl = pd.ExcelFile(file_path)
        print(f"Available sheets: {xl.sheet_names}")
        raise
    return df


def get_col(df: pd.DataFrame, name: str) -> pd.Series:
    for c in df.columns:
        if str(c).strip().upper() == name.upper():
            return df[c]
    raise KeyError(f"Column '{name}' not found. Available: {list(df.columns)}")


def derive_gtc_tc_type(func_desc: str) -> str:
    if not isinstance(func_desc, str):
        return "P"
    text = func_desc.upper()
    if "ATTN" in text or "SOP" in text or "CLOCK" in text:
        return "S"
    if "ENABLE" in text or "DISABLE" in text:
        return "T"
    return "P"


# def derive_gtc_tc_alias(func_desc: str) -> str:
#     if not isinstance(func_desc, str):
#         return ""
#     text = func_desc.lower()
#     m = re.search(r"pos[\s\-_]*([123])", text)
#     if m:
#         return f"pos-{m.group(1)}"
#     if "helix trip" in text:
#         if "enable" in text:
#             return "helix trip enable"
#         if "disable" in text:
#             return "helix trip disable"
#     if re.search(r"\bon\b", text):
#         return "on"
#     if re.search(r"\boff\b", text):
#         return "off"
#     return ""


def derive_bus_read_time(func_desc: str):
    if not isinstance(func_desc, str):
        return np.nan
    text = func_desc.lower()
    if re.search(r"\bon\b", text) or re.search(r"\boff\b", text):
        return 2
    return np.nan


def build_all_tc(df: pd.DataFrame, out_path: Path, p_id) -> pd.DataFrame:
    tc_func_desc = get_col(df, SRC_FUNC_DESC)

    df_out = pd.DataFrame()
    df_out["tc_id"] = range(1, len(df) + 1)
    df_out["shape_id_fk"] = 0
    df_out["tc_description"] = tc_func_desc
    df_out["tm_id"] = 0
    df_out["p_id"] = p_id  # real project id, from selected/created project
    df_out["advantek_tc_address"] = get_col(df, SRC_IP_ADDRESS)
    df_out["word_pin"] = get_col(df, SRC_ASSIGNED_PIN)
    df_out["cmd_data"] = 64
    df_out["tc_priority"] = 1
    df_out["simulator_type"] = "ADVANTEK"
    df_out["desired_tm_raw"] = np.nan
    df_out["gtc_tc_type"] = tc_func_desc.apply(derive_gtc_tc_type)
    df_out["gtc_tc_alias"] = np.nan 
    df_out["bus_read_time"] = tc_func_desc.apply(derive_bus_read_time)
    df_out["rt"] = np.nan
    df_out["sa"] = np.nan
    df_out["advantek_tc_type"] = np.nan
    df_out["bus_ip"] = np.nan
    df_out["boa_cat_id"] = np.nan
    df_out["tc_operation_mode"] = "Normal"
    df_out["tc_ch_pos"] = np.nan
    df_out["ss_id"] = np.nan
    df_out["tc_length"] = np.nan
    df_out["tc_format"] = np.nan
    df_out["tc_club_id"] = np.nan
    df_out["tc_club_vale"] = np.nan
    df_out["tc_wait_time"] = np.nan

    df_out.to_excel(out_path, sheet_name="all_tc", index=False)
    print(f"Wrote: {out_path}")
    return df_out


def build_all_tm(df: pd.DataFrame, out_path: Path, p_id) -> pd.DataFrame:
    df_out = pd.DataFrame()
    df_out["tm_id"] = range(1, len(df) + 1)
    df_out["p_id"] = p_id  # real project id, from user's selected/created project
    df_out["tm_data_time"] = "null"
    df_out["tm_description"] = get_col(df, SRC_TM_FUNC_DESC)
    df_out["simulator_type"] = "ADVANTEK"
    df_out["rt_addres"] = "null"
    df_out["sa_addres"] = "null"
    df_out["word_count"] = "null"
    df_out["equation_type"] = "null"
    df_out["Advantek_tm_ip"] = get_col(df, SRC_TM_IP_ADDRESS)
    df_out["Advantek_card_type"] = get_col(df, SRC_TM_CARD_TYPE)
    df_out["tm_channel_position"] = get_col(df, SRC_TM_ASSIGNED_PIN)
    df_out["tm_length"] = get_col(df, SRC_TM_LENGTH)
    df_out["tm_value"] = "null"
    df_out["shape_id_fk"] = 0
    df_out["gtm_tm_type"] = "null"
    df_out["gtm_alias"] = "null"
    df_out["gtm_type_conversions"] = "null"
    df_out["gtm_tm_encoding"] = "null"
    df_out["ref_tm_id"] = "null"
    df_out["ref_tm_raw_value"] = "null"
    df_out["binary_decode_logic"] = "null"
    df_out["ss_id"] = "null"
    df_out["tm_type_1553"] = "processed"

    df_out.to_excel(out_path, sheet_name="all_tm", index=False)
    print(f"Wrote: {out_path}")
    return df_out


def run(master_summary_path: str, output_folder: str, p_id) -> dict:
    src_path = Path(master_summary_path)
    out_folder = Path(output_folder)

    df_tc = load_sheet(src_path, SHEET_TC)
    print(f"Loaded '{SHEET_TC}': {df_tc.shape[0]} rows, {df_tc.shape[1]} cols")
    all_tc_path = out_folder / "all_tc.xlsx"
    df_tc_out = build_all_tc(df_tc, all_tc_path, p_id)

    df_tm = load_sheet(src_path, SHEET_TM)
    print(f"Loaded '{SHEET_TM}': {df_tm.shape[0]} rows, {df_tm.shape[1]} cols")
    all_tm_path = out_folder / "all_tm.xlsx"
    df_tm_out = build_all_tm(df_tm, all_tm_path, p_id)

    return {
        "all_tc_path": str(all_tc_path),
        "all_tm_path": str(all_tm_path),
        "tc_count": len(df_tc_out),
        "tm_count": len(df_tm_out),
    }
