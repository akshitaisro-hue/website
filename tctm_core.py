import os
import re
import html
import pandas as pd


def increment_last_octet(ip_str):
    try:
        parts = ip_str.split('.')
        if len(parts) == 4:
            last_val = int(parts[3]) + 1
            parts[3] = f"{last_val:03d}" if len(parts[3]) == 3 else str(last_val)
            return '.'.join(parts)
    except Exception:
        pass
    return ip_str


def idx_2_column(n):
    if n <= 0:
        return ""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def increment_card_string(card_str, padding=2):
    return f"{int(card_str) + 1:0{padding}d}"


def find_sheet_header_and_bounds(file_path, sheet_name, standard_headers):
    try:
        raw_preview = pd.read_excel(file_path, sheet_name=sheet_name, nrows=61, header=None)
        for r in range(len(raw_preview)):
            row_content = raw_preview.iloc[r].astype(str).str.strip().tolist()
            row_content_lower = [val.lower() for val in row_content]

            if "pin" in row_content_lower:
                p_idx = row_content_lower.index("pin")
                if any("type" in val for val in row_content_lower):
                    header_row_idx = r
                    start_col_idx = p_idx

                    sim_matches = [i for i, val in enumerate(row_content_lower) if 'tctm simulator' in val]
                    if sim_matches:
                        end_col_idx = sim_matches[0]
                    else:
                        max_found_col = 0
                        for header in standard_headers:
                            header_target = header.lower().strip()
                            loc_matches = [i for i, val in enumerate(row_content_lower) if header_target in val]
                            if loc_matches and loc_matches[0] > max_found_col:
                                max_found_col = loc_matches[0]
                        end_col_idx = max_found_col + 1
                    return header_row_idx, start_col_idx, end_col_idx
    except Exception:
        pass
    return None, None, None


standard_headers = [
    "Pin", "Type", "Amp", "Functional Description",
    "Main Dest-1", "Patch-1", "Patch-2", "Main Dest-2", "Wire", "Gauge"
]


def scan_directory_files(folder_path):
    input_files = []
    output_cards = []
    all_dir_files = os.listdir(folder_path)
    for fname in all_dir_files:
        if fname.lower().endswith(('.xlsx', '.xlsv')):
            fname_lower = fname.lower()
            if ("_assigned" not in fname_lower and
                fname_lower != "master_summary_tctm.xlsx" and
                fname_lower != "start_tctm_cards.xlsx"):
                input_files.append(fname)
            if fname_lower == "start_tctm_cards.xlsx":
                output_cards.append(fname)
    return input_files, output_cards


def run(input_folder: str, output_folder: str, tctm_card_selection: int = 1, generate_pdf: bool = True) -> dict:
    """
    input_folder: uploaded raw excel files
    output_folder: where MASTER_SUMMARY_TCTM.xlsx / final_op_assigned.pdf / *_Assigned.xlsx go
    tctm_card_selection: 1 = default counters, 0 = load from START_TCTM_CARDS.xlsx (if present)
    generate_pdf: replaces the old input()-driven "generate PDF? y/n" prompt
    """
    os.makedirs(output_folder, exist_ok=True)
    folder_path = input_folder

    input_files, output_cards = scan_directory_files(folder_path)
    if not input_files:
        raise RuntimeError(f"No valid input workbooks found in the working dir: {folder_path}")

    c_1_name = output_cards[0] if output_cards else ""
    c_1 = os.path.join(folder_path, c_1_name)

    if tctm_card_selection == 0 and output_cards and os.path.exists(c_1):
        try:
            c_1_out = pd.read_excel(c_1, sheet_name="Cards")
        except Exception:
            print("Error reading Excel baseline tracker. Reverting to default counters.")
            tctm_card_selection = 1
    elif tctm_card_selection == 0:
        print("START_TCTM_CARDS.xlsx not discovered. Reverting to default baseline counters.")
        tctm_card_selection = 1

    if tctm_card_selection == 1:
        card_28V = '2801'
        pin_28V = 1
        card_5V = '0501'
        pin_5V = 1
        card_Lv1 = '01'
        pin_Lv1 = 1
        card_Data = '01'
        pin_Data = 1
        card_Analog = '01'
        pin_Analog = 1
        card_Digital = '01'
        pin_Digital = 1
        card_RF = '01'
        pin_RF = 1
        card_Therm = '01'
        pin_Therm = 1

        Initial_card_IP_28V = '192.168.001.100'
        Initial_card_IP_5V = '192.168.001.200'
        Initial_card_IP_Lv1 = '192.168.001.300'
        Initial_card_IP_Data = '192.168.001.400'
        Initial_card_IP_Analog = '192.168.001.500'
        Initial_card_IP_Digital = '192.168.001.600'
        Initial_card_IP_RF = '192.168.001.700'
        Initial_card_IP_Therm = '192.168.001.800'
    else:
        def parse_card_no(val):
            match = re.search(r'\d+', str(val))
            return str(int(float(match.group()))) if match else '1'

        card_5V = parse_card_no(c_1_out['card_no'].iloc[0])
        Initial_card_IP_5V = str(c_1_out['ip_address'].iloc[0])
        card_28V = parse_card_no(c_1_out['card_no'].iloc[1])
        Initial_card_IP_28V = str(c_1_out['ip_address'].iloc[1])
        card_Data = parse_card_no(c_1_out['card_no'].iloc[2])
        Initial_card_IP_Data = str(c_1_out['ip_address'].iloc[2])
        card_Lv1 = parse_card_no(c_1_out['card_no'].iloc[3])
        Initial_card_IP_Lv1 = str(c_1_out['ip_address'].iloc[3])
        card_RF = parse_card_no(c_1_out['card_no'].iloc[4])
        Initial_card_IP_RF = str(c_1_out['ip_address'].iloc[4])
        card_Digital = parse_card_no(c_1_out['card_no'].iloc[5])
        Initial_card_IP_Digital = str(c_1_out['ip_address'].iloc[5])
        card_Analog = parse_card_no(c_1_out['card_no'].iloc[6])
        Initial_card_IP_Analog = str(c_1_out['ip_address'].iloc[6])
        card_Therm = parse_card_no(c_1_out['card_no'].iloc[7])
        Initial_card_IP_Therm = str(c_1_out['ip_address'].iloc[7])
        pin_28V = pin_5V = pin_Lv1 = pin_Data = pin_Analog = pin_Digital = pin_RF = pin_Therm = 1

    TC_function, TM_function = '01', '02'
    card_type_5V, card_type_28V, card_type_Data, card_type_Level = '01', '02', '04', '03'
    card_type_Digital, card_type_RF, card_type_Analog, card_type_Therm = '05', '07', '06', '08'

    master_summary_tc = pd.DataFrame()
    master_summary_tm = pd.DataFrame()
    unassigned_list = []
    pdf_data_store = []

    for f_idx in range(len(input_files)):
        input_file_name = str(input_files[f_idx])
        input_file = os.path.join(folder_path, input_file_name)
        base_name, extension = os.path.splitext(input_file_name)
        output_file = os.path.join(output_folder, base_name + "_Assigned" + extension)

        print(f'\n=======================\n Starting Excel Workbook [{f_idx + 1}/{len(input_files)}]:{input_file_name}\n=======================')

        skip_entire_file = False
        if os.path.isfile(output_file):
            try:
                out_sheets = pd.ExcelFile(output_file).sheet_names
                valid_sheets_checked = completed_sheets = 0
                for sh_name in out_sheets:
                    if any(x in sh_name.lower() for x in ["master-summary", "summary", "list"]):
                        continue
                    valid_sheets_checked += 1
                    output_data = pd.read_excel(output_file, sheet_name=sh_name)
                    output_data.columns = output_data.columns.str.strip()
                    has_all_headers = all(header in output_data.columns for header in standard_headers)
                    if has_all_headers and "TCTM SIMULATOR" in output_data.columns:
                        simulator_headers = output_data["TCTM SIMULATOR"].fillna("").astype(str)
                        if not (simulator_headers == "").any():
                            completed_sheets += 1
                if valid_sheets_checked > 0 and (completed_sheets == valid_sheets_checked):
                    print(f'Message: Output workbook "{output_file}" already fully processed. Skipping.\n')
                    for sh_name in out_sheets:
                        if any(x in sh_name.lower() for x in ["master-summary", "summary", "list"]):
                            continue
                        existing_assigned_data = pd.read_excel(output_file, sheet_name=sh_name)
                        pdf_data_store.append([base_name, sh_name, existing_assigned_data])
                    skip_entire_file = True
            except Exception:
                skip_entire_file = False

        if skip_entire_file:
            continue

        all_sheets = pd.ExcelFile(input_file).sheet_names
        for s in range(len(all_sheets)):
            current_sheet_for_test = all_sheets[s]
            if any(x in current_sheet_for_test.lower() for x in ["master-summary", "summary", "list"]):
                continue

            print(f'Processing Sheet: {current_sheet_for_test}\n')
            header_row_idx, start_col_idx, end_col_idx = find_sheet_header_and_bounds(input_file, current_sheet_for_test, standard_headers)

            if header_row_idx is None:
                print(f'Valid header row not defined in {current_sheet_for_test}. Skipping.\n')
                continue

            l_start = idx_2_column(start_col_idx + 1)
            l_end = idx_2_column(end_col_idx + 1)

            try:
                input_data = pd.read_excel(
                    input_file, sheet_name=current_sheet_for_test,
                    header=header_row_idx, nrows=51, usecols="A:K"
                )
                if len(input_data) > 50:
                    input_data = input_data.iloc[:-1]
                input_data.columns = input_data.columns.str.strip()
                if "TCTM SIMULATOR" not in input_data.columns:
                    input_data["TCTM SIMULATOR"] = ""
                else:
                    input_data["TCTM SIMULATOR"] = input_data["TCTM SIMULATOR"].fillna("").astype(str)
            except Exception:
                print(f'Failed reading input table range for sheet {current_sheet_for_test}\n')
                continue

            type_col = input_data["Type"].fillna("").astype(str).str.strip().str.lower() if "Type" in input_data.columns else pd.Series([""]*len(input_data))
            amp_col = input_data["Amp"].fillna("").astype(str).str.strip().str.lower() if "Amp" in input_data.columns else pd.Series([""]*len(input_data))
            f_desc_col = input_data["Functional Description"].fillna("").astype(str).str.strip()
            f_desc_col_lower = f_desc_col.str.lower()
            dest1_col = input_data["Main Dest-1"].fillna("").astype(str).str.strip()
            dest2_col = input_data["Main Dest-2"].fillna("").astype(str).str.strip()

            d2_col_upper = dest2_col.str.upper()
            rf_row_group = [""] * len(input_data)
            rf_row_pos = [None] * len(input_data)
            last_valid_group_name = "UNKNOWN_RF"

            for k in range(len(input_data)):
                if input_data["TCTM SIMULATOR"].iloc[k] != "":
                    continue
                is_meta_gnd_rf = any(x in type_col.iloc[k] for x in ["lvl", "rfb"]) and any(x in amp_col.iloc[k] for x in ["0/5v", "5v", "5 v"])
                is_meta_empty = (type_col.iloc[k] in ["", "-"]) and (amp_col.iloc[k] in ["", "-"])
                has_rf_desc_keyword = any(x in f_desc_col_lower.iloc[k] for x in ["pos-", "pos", "position-", "position"])

                if is_meta_gnd_rf or (is_meta_empty and has_rf_desc_keyword):
                    pos_match = re.search(r'pos[- ](\d+)', f_desc_col_lower.iloc[k])
                    p_num = float(pos_match.group(1)) if pos_match else 1.0
                    id_match = re.search(r'([A-Za-z0-9]+-\d+)', d2_col_upper.iloc[k])
                    if id_match:
                        g_name = str(id_match.group(1))
                        last_valid_group_name = g_name
                    else:
                        g_name = last_valid_group_name if p_num > 1 else "unknown_rf"
                    rf_row_group[k] = g_name
                    rf_row_pos[k] = p_num

            seen_rf_groups = set()
            unique_rf_groups = [g for g in rf_row_group if g and not (g in seen_rf_groups or seen_rf_groups.add(g))]

            for cur_group in unique_rf_groups:
                match_idx = [i for i in range(len(input_data)) if rf_row_group[i] == cur_group]
                g_len = len(match_idx)
                if (pin_RF + g_len - 1) > 45:
                    card_RF = increment_card_string(card_RF, padding=2)
                    Initial_card_IP_RF = increment_last_octet(Initial_card_IP_RF)
                    pin_RF = 1
                for idx_pos in match_idx:
                    target_pos = rf_row_pos[idx_pos]
                    calculated_pin = int(pin_RF + (target_pos - 1))
                    sig_id = f"TMRF{int(card_RF):02d}-{calculated_pin:02d}"
                    input_data.at[idx_pos, "TCTM SIMULATOR"] = sig_id
                    if target_pos == 1:
                        raw_desc = f_desc_col.iloc[idx_pos]
                        clean_desc = re.sub(r'\s*pos[-/\s]*\d+([-/\s]*\d*)*', '', raw_desc, flags=re.IGNORECASE)
                        clean_desc = re.sub(r'\b(status|selected)\b', '', clean_desc, flags=re.IGNORECASE)
                        clean_desc = re.sub(r'[-_\s\(\)\[\]\/]+$', '', clean_desc)
                        rf_desc_clean = clean_desc.strip() + "Status"
                        new_tm_row = pd.DataFrame([{
                            'TM_CARD': f"TMRF{int(card_RF):02d}", 'IP_ADDRESS': Initial_card_IP_RF, 'TM_FUNCTION': TM_function,
                            'TM_CARD_TYPE': card_type_RF, 'ASSIGNED_TM_PIN': f"{calculated_pin:02d}", 'TM_FUNCTIONAL_DESCRIPTION': rf_desc_clean
                        }])
                        master_summary_tm = pd.concat([master_summary_tm, new_tm_row], ignore_index=True)
                pin_RF += g_len
                if pin_RF > 45:
                    card_RF = increment_card_string(card_RF, padding=2)
                    Initial_card_IP_RF = increment_last_octet(Initial_card_IP_RF)
                    pin_RF = 1

            for i in range(len(input_data)):
                if input_data["TCTM SIMULATOR"].iloc[i] == "":
                    is_pulse = any(x in type_col.iloc[i] for x in ["pulse", "puls", "pul"])
                    is_28V_family = any(x in amp_col.iloc[i] for x in ["28v","29v","30v","26","28","29","30","42","-42","26-","28-","29-","30-"])
                    if is_pulse and is_28V_family:
                        pulse_idx = i
                        matchfound = False
                        gnd_id = -1
                        TC_descrip_raw = f_desc_col.iloc[pulse_idx]
                        base_TC_descrip = re.sub(r'(\s+live|\s+pos-\d|\s+\(l1\))', '', f_desc_col_lower.iloc[pulse_idx], flags=re.IGNORECASE)
                        base_TC_descrip = base_TC_descrip.strip()
                        des1_pulse = dest1_col.iloc[pulse_idx]
                        des2_pulse = dest2_col.iloc[pulse_idx]
                        tokens1 = re.findall(r'\.(\d+)$', des1_pulse)
                        tokens2 = re.findall(r'\.(\d+)$', des2_pulse)
                        targetpinstrings = ""
                        target_pin_clean = ""
                        if tokens1 and 'pip' in des1_pulse.lower():
                            targetpinstrings = str(tokens1[0])
                        elif tokens2 and 'pip' in des2_pulse.lower():
                            targetpinstrings = str(tokens2[0])
                        elif tokens1:
                            targetpinstrings = str(tokens1[0])
                        elif tokens2:
                            targetpinstrings = str(tokens2[0])
                        if targetpinstrings:
                            target_pin_clean = str(int(float(targetpinstrings)))
                        conn1_pulse = re.sub(r'\.\d+$', '', des1_pulse).lower()
                        conn2_pulse = re.sub(r'\.\d+$', '', des2_pulse).lower()

                        for j in range(len(input_data)):
                            if input_data["TCTM SIMULATOR"].iloc[j] != "" or j == pulse_idx:
                                continue
                            gnd_descrip = f_desc_col_lower.iloc[j]
                            has_explicit_gnd_type = any(x in type_col.iloc[j] for x in ["gnd", "ground", "rtn", "return", "ret"])
                            is_meta_empty = (type_col.iloc[j] == "" or type_col.iloc[j] == "-") and (amp_col.iloc[j] == "" or amp_col.iloc[j] == "-")
                            has_desc_gnd_keyword = any(x in gnd_descrip for x in ["return", "ground", "rtn", "gnd", "ret"])
                            if has_explicit_gnd_type or (is_meta_empty and has_desc_gnd_keyword):
                                des1_gnd = dest1_col.iloc[j]
                                des2_gnd = dest2_col.iloc[j]
                                conn1_gnd = re.sub(r'\.\d+$', '', des1_gnd).lower()
                                conn2_gnd = re.sub(r'\.\d+$', '', des2_gnd).lower()
                                if targetpinstrings != "" and (
                                    ("pin-" + targetpinstrings) in gnd_descrip or
                                    ("pin-" + target_pin_clean) in gnd_descrip or
                                    ("pin " + target_pin_clean) in gnd_descrip
                                ):
                                    gnd_id = j
                                    matchfound = True
                                    break
                                else:
                                    is_connector_match = False
                                    if conn1_pulse != "" and conn1_pulse != "nc":
                                        if conn1_pulse in [conn1_gnd, conn2_gnd]:
                                            is_connector_match = True
                                    if not is_connector_match and conn2_pulse != "" and conn2_pulse != "nc":
                                        if conn2_pulse in [conn1_gnd, conn2_gnd]:
                                            is_connector_match = True
                                    if is_connector_match:
                                        words_pulse = re.findall(r'\w+', base_TC_descrip)
                                        if words_pulse and any(word in gnd_descrip for word in words_pulse):
                                            gnd_id = j
                                            matchfound = True
                                            break

                        if matchfound:
                            hw_id_pulse = f"TC{int(card_28V)}-{int(pin_28V):02d}"
                            hw_id_gnd = f"TC{int(card_28V)}-{int(pin_28V + 1):02d}"
                            input_data.at[pulse_idx, "TCTM SIMULATOR"] = hw_id_pulse
                            input_data.at[gnd_id, "TCTM SIMULATOR"] = hw_id_gnd
                            new_tc_row = pd.DataFrame([{
                                'TC_CARD': f"TC{int(card_28V)}", 'IP_ADDRESS': Initial_card_IP_28V, 'TC_FUNCTION': TC_function,
                                'TC_CARD_TYPE': card_type_28V, 'ASSIGNED_TC_PIN': f"{int(pin_28V):02d}", 'FUNCTIONAL_DESCRIPTION': TC_descrip_raw
                            }])
                            master_summary_tc = pd.concat([master_summary_tc, new_tc_row], ignore_index=True)
                            pin_28V = pin_28V + 2
                            if pin_28V > 49:
                                card_28V = str(int(card_28V) + 1)
                                Initial_card_IP_28V = increment_last_octet(Initial_card_IP_28V)
                                pin_28V = 1
                        else:
                            unassigned_list.append(f"28V comand missing GND -> File: {input_file_name} | Sheet: {current_sheet_for_test} | Row: {pulse_idx + 1} | Description: {TC_descrip_raw}")

            is_data_cmd = [False] * len(input_data)
            group_ids = [""] * len(input_data)
            for k in range(len(input_data)):
                if input_data["TCTM SIMULATOR"].iloc[k] == "":
                    is_clock = "clock" in f_desc_col_lower.iloc[k]
                    is_data = "data" in f_desc_col_lower.iloc[k]
                    is_strobe = any(x in f_desc_col_lower.iloc[k] for x in ["strobe", "tp", "transfer pulse", " transfer_pulse", "transfer"])
                    is_amp_data = any(x in amp_col.iloc[k] for x in ["5v", "5v.", "0-5", "0/5", "5 v", "5 v."]) and not any(x in amp_col.iloc[k] for x in ["+/-", "+ / -", "+/ -", "+ /-"])
                    is_type_data = any(x in type_col.iloc[k] for x in ["pulse", "cmos"])
                    is_meta_empty = (type_col.iloc[k] == "" or type_col.iloc[k] == "-") and (amp_col.iloc[k] == "" or amp_col.iloc[k] == "-")
                    if ((is_amp_data and is_type_data) or is_clock or is_data or is_strobe) or (is_meta_empty and (is_clock or is_data or is_strobe)):
                        is_data_cmd[k] = True
                        tokens1 = re.findall(r'([A-Za-z0-9]+-\d+)', dest1_col.iloc[k])
                        tokens2 = re.findall(r'([A-Za-z0-9]+-\d+)', dest2_col.iloc[k])
                        group_ids[k] = str(tokens2[0]).upper() if tokens2 else (str(tokens1[0]).upper() if tokens1 else "UNKNOWN")

            seen_groups = set()
            unique_data_groups = [grp for idx, grp in enumerate(group_ids) if is_data_cmd[idx] and grp not in seen_groups and not seen_groups.add(grp)]

            for g in range(len(unique_data_groups)):
                g_indices = [i for i, (cmd, grp) in enumerate(zip(is_data_cmd, group_ids)) if cmd and grp == unique_data_groups[g]]
                if g_indices:
                    if pin_Data > 43:
                        card_Data = increment_card_string(card_Data, padding=2)
                        Initial_card_IP_Data = increment_last_octet(Initial_card_IP_Data)
                        pin_Data = 1
                    clock_row_id = -1
                    for w_idx in g_indices:
                        d_text = f_desc_col_lower.iloc[w_idx]
                        if "clock" in d_text:
                            sig_id, clock_row_id = f"TCDATA{int(card_Data):02d}-{int(pin_Data):02d}", w_idx
                        elif any(x in d_text for x in ["strobe", "tp", "transfer pulse", "transfer_pulse", "transfer"]):
                            sig_id = f"TCDATA{int(card_Data):02d}-{int(pin_Data + 2):02d}"
                        elif "data" in d_text:
                            sig_id = f"TCDATA{int(card_Data):02d}-{int(pin_Data + 1):02d}"
                        input_data.at[w_idx, "TCTM SIMULATOR"] = sig_id
                    if clock_row_id != -1:
                        clean_desc = re.sub(r'\s*[\(\[-]?\s*[\)\]-]?\s*', '', f_desc_col.iloc[clock_row_id], flags=re.IGNORECASE).strip()
                        new_tc_row = pd.DataFrame([{'TC_CARD': f"TCDATA{int(card_Data):02d}", 'IP_ADDRESS': Initial_card_IP_Data, 'TC_FUNCTION': TC_function, 'TC_CARD_TYPE': card_type_Data, 'ASSIGNED_TC_PIN': f"{int((pin_Data + 2) / 3):02d}", 'FUNCTIONAL_DESCRIPTION': clean_desc + " BOA"}])
                        master_summary_tc = pd.concat([master_summary_tc, new_tc_row], ignore_index=True)
                    pin_Data += 3

            for k in range(len(input_data)):
                if input_data["TCTM SIMULATOR"].iloc[k] == "":
                    if any(x in amp_col.iloc[k] for x in ["5v", "5v.", "5 v", "5 v."]):
                        if any(x in type_col.iloc[k] for x in ["pulse", "pulse.", "pul", "pul."]) or ("cmos" in type_col.iloc[k] and any(x in f_desc_col_lower.iloc[k] for x in ["clock", "data", "strobe", "tp", "transfer pulse", "transfer_pulse", "transfer"])):
                            input_data.at[k, "TCTM SIMULATOR"] = f"TC{int(card_5V):04d}-{int(pin_5V):02d}"
                            new_tc_row = pd.DataFrame([{'TC_CARD': f"TC{int(card_5V):04d}", 'IP_ADDRESS': Initial_card_IP_5V, 'TC_FUNCTION': TC_function, 'TC_CARD_TYPE': card_type_5V, 'ASSIGNED_TC_PIN': f"{int(pin_5V):02d}", 'FUNCTIONAL_DESCRIPTION': f_desc_col.iloc[k]}])
                            master_summary_tc = pd.concat([master_summary_tc, new_tc_row], ignore_index=True)
                            pin_5V += 1
                            if pin_5V > 45:
                                card_5V = increment_card_string(card_5V, padding=2)
                                Initial_card_IP_5V = increment_last_octet(Initial_card_IP_5V)
                                pin_5V = 1

            for k in range(len(input_data)):
                if input_data["TCTM SIMULATOR"].iloc[k] == "":
                    is_level_type = any(x in type_col.iloc[k] for x in ["level", "lev"])
                    is_amp_level = any(x in amp_col.iloc[k] for x in ["5v", "28v", "70v", "26", "28", "29", "30", "70", "42", "5"])
                    is_meta_empty = (type_col.iloc[k] == "" or type_col.iloc[k] == "-") and (amp_col.iloc[k] == "" or amp_col.iloc[k] == "-")
                    has_level_keyword = any(x in f_desc_col_lower.iloc[k] for x in ["level", "lev", "lvl"])
                    if (is_level_type and is_amp_level) or (is_meta_empty and has_level_keyword):
                        input_data.at[k, "TCTM SIMULATOR"] = f"TCLVL{int(card_Lv1):02d}-{int(pin_Lv1):02d}"
                        for suff in ["[SET]", "[RESET]"]:
                            new_tc_row = pd.DataFrame([{'TC_CARD': f"TCLVL{int(card_Lv1):02d}", 'IP_ADDRESS': Initial_card_IP_Lv1, 'TC_FUNCTION': TC_function, 'TC_CARD_TYPE': card_type_Level, 'ASSIGNED_TC_PIN': f"{int(pin_Lv1):02d}", 'FUNCTIONAL_DESCRIPTION': f_desc_col.iloc[k] + suff}])
                            master_summary_tc = pd.concat([master_summary_tc, new_tc_row], ignore_index=True)
                        pin_Lv1 += 1
                        if pin_Lv1 > 45:
                            card_Lv1 = increment_card_string(card_Lv1, padding=2)
                            Initial_card_IP_Lv1 = increment_last_octet(Initial_card_IP_Lv1)
                            pin_Lv1 = 1

            for k in range(len(input_data)):
                if input_data["TCTM SIMULATOR"].iloc[k] == "":
                    is_dig_type = any(x in type_col.iloc[k] for x in ["dig", "dig.", "digital", "digital.", "nb"])
                    is_amp_dig = any(x in amp_col.iloc[k] for x in ["5v", "5v.", "0/5v", "0/5v.", "0/5"])
                    is_meta_empty = (type_col.iloc[k] == "" or type_col.iloc[k] == "-") and (amp_col.iloc[k] == "" or amp_col.iloc[k] == "-")
                    has_dig_keyword = any(x in f_desc_col_lower.iloc[k] for x in ["digital", "dig", "dig."])
                    if (is_dig_type and is_amp_dig) or (is_meta_empty and has_dig_keyword):
                        input_data.at[k, "TCTM SIMULATOR"] = f"TMDIG{int(card_Digital):02d}-{int(pin_Digital):02d}"
                        new_tm_row = pd.DataFrame([{'TM_CARD': f"TMDIG{int(card_Digital):02d}", 'IP_ADDRESS': Initial_card_IP_Digital, 'TM_FUNCTION': TM_function, 'TM_CARD_TYPE': card_type_Digital, 'ASSIGNED_TM_PIN': f"{int(pin_Digital):02d}", 'FUNCTIONAL_DESCRIPTION': f_desc_col.iloc[k]}])
                        master_summary_tm = pd.concat([master_summary_tm, new_tm_row], ignore_index=True)
                        pin_Digital += 1
                        if pin_Digital > 45:
                            card_Digital = increment_card_string(card_Digital, padding=2)
                            Initial_card_IP_Digital = increment_last_octet(Initial_card_IP_Digital)
                            pin_Digital = 1

            for k in range(len(input_data)):
                if input_data["TCTM SIMULATOR"].iloc[k] == "":
                    is_ang_type = any(x in type_col.iloc[k] for x in ["ana", "ana.", "analog", "analog.", "ang", "ang."])
                    is_amp_ang = any(x in amp_col.iloc[k] for x in ["5v", "5v.", "0/5v", "0/5v.", "+/-5v", "+/-5v.", "+/- 5v", "+/- 5v.", "5 v", "5 v."])
                    is_meta_empty = (type_col.iloc[k] == "" or type_col.iloc[k] == "-") and (amp_col.iloc[k] == "" or amp_col.iloc[k] == "-")
                    has_ang_keyword = any(x in f_desc_col_lower.iloc[k] for x in ["analog", "ang", "ang."])
                    if (is_ang_type and is_amp_ang) or (is_meta_empty and has_ang_keyword):
                        input_data.at[k, "TCTM SIMULATOR"] = f"TMANG{int(card_Analog):02d}-{int(pin_Analog):02d}"
                        new_tm_row = pd.DataFrame([{'TM_CARD': f"TMANG{int(card_Analog):02d}", 'IP_ADDRESS': Initial_card_IP_Analog, 'TM_FUNCTION': TM_function, 'TM_CARD_TYPE': card_type_Analog, 'ASSIGNED_TM_PIN': f"{int(pin_Analog):02d}", 'FUNCTIONAL_DESCRIPTION': f_desc_col.iloc[k]}])
                        master_summary_tm = pd.concat([master_summary_tm, new_tm_row], ignore_index=True)
                        pin_Analog += 1
                        if pin_Analog > 45:
                            card_Analog = increment_card_string(card_Analog, padding=2)
                            Initial_card_IP_Analog = increment_last_octet(Initial_card_IP_Analog)
                            pin_Analog = 1

            is_thermistor_live = [False] * len(input_data)
            is_thermistor_rtn = [False] * len(input_data)
            thermistor_assigned = [""] * len(input_data)
            valid_amps = ["5k", "10k", "0-5", "0/5", "5v", "+/-5", "+/- 5", "+_5", "+_ 5"]
            valid_types = ["thermis", "thm", "thr", "th"]

            for k in range(len(input_data)):
                if input_data["TCTM SIMULATOR"].iloc[k] == "":
                    f_desc, t_str, a_str = f_desc_col_lower.iloc[k], type_col.iloc[k], amp_col.iloc[k]
                    is_meta_empty = (t_str == "" or t_str == "-") and (a_str == "" or a_str == "-")
                    has_desc_live = is_meta_empty and any(x in f_desc for x in ["temp", "thermis", "thm", "thr"]) and "live" in f_desc
                    if (any(x in a_str for x in valid_amps) and any(x in t_str for x in valid_types)) or has_desc_live or (is_meta_empty and bool(re.search(r'(thermistor|temperature|temp|thm|thr)[-_\s]?1$', f_desc))):
                        is_thermistor_live[k] = True
                        continue
                    if (any(x in a_str for x in valid_amps) and any(x in t_str for x in valid_types)) or (is_meta_empty and any(x in f_desc for x in ["return", "ret", "rtn"])) or (is_meta_empty and bool(re.search(r'(thermistor|temperature|temp|thm|thr)[-_\s]?2$', f_desc))) or is_meta_empty:
                        is_thermistor_rtn[k] = True

            live_indices = [idx for idx, val in enumerate(is_thermistor_live) if val]
            for live_idx in live_indices:
                clean_id = re.sub(r'[-_\s](lead[-_\s]?1|live|pos[-_\s]?\d|thermistor_?1|temperature_?1|temp_?1|1)$', '', f_desc_col_lower.iloc[live_idx]).strip()
                match_rtn_idx = -1
                for j in range(len(input_data)):
                    if j == live_idx or not is_thermistor_rtn[j] or thermistor_assigned[j] != "":
                        continue
                    search_desc = f_desc_col_lower.iloc[j]
                    clean_rtn = re.sub(r'[-_\s](rtn|return|ret|lead[-_\s]?2|thermistor_?2|temperature_?2|temp_?2|2)$', '', search_desc)
                    if re.sub(r'(thermistor|temperature|temp|thm|thr)[-_\s]?2$', '', clean_rtn) == clean_id or clean_id in search_desc:
                        match_rtn_idx = j
                        break

                if pin_Therm > 49:
                    card_Therm = increment_card_string(card_Therm, padding=2)
                    Initial_card_IP_Therm = increment_last_octet(Initial_card_IP_Therm)
                    pin_Therm = 1

                thermistor_assigned[live_idx] = f"TMTHR{int(card_Therm):02d}-{int(pin_Therm):02d}"
                new_tm_row = pd.DataFrame([{'TM_CARD': f"TMTHR{int(card_Therm):02d}", 'IP_ADDRESS': Initial_card_IP_Therm, 'TM_FUNCTION': TM_function, 'TM_CARD_TYPE': card_type_Therm, 'ASSIGNED_TM_PIN': f"{int(pin_Therm):02d}", 'FUNCTIONAL_DESCRIPTION': f_desc_col.iloc[live_idx]}])
                master_summary_tm = pd.concat([master_summary_tm, new_tm_row], ignore_index=True)
                pin_Therm += 1

                if match_rtn_idx != -1:
                    thermistor_assigned[match_rtn_idx] = f"TMTHR{int(card_Therm):02d}-{int(pin_Therm):02d}"
                    pin_Therm += 1
                else:
                    pin_Therm += 1
                    unassigned_list.append(f"Thermistor Missing Return Row -> File : {input_file_name} | Sheet: {current_sheet_for_test} | Row: {live_idx + 1}")

            for k in range(len(input_data)):
                if thermistor_assigned[k] != "":
                    input_data.at[k, "TCTM SIMULATOR"] = thermistor_assigned[k]

            # bug fixed: original had mixed tabs/spaces here
            file_mode = 'a' if os.path.isfile(output_file) else 'w'
            write_kwargs = {'engine': 'openpyxl', 'mode': file_mode}
            if file_mode == 'a':
                write_kwargs['if_sheet_exists'] = 'replace'
            try:
                with pd.ExcelWriter(output_file, **write_kwargs) as writer:
                    input_data.to_excel(writer, sheet_name=current_sheet_for_test, index=False)
            except Exception as e:
                print(f'Warning: Sheet "{output_file}" is locked. {e}')

            pdf_data_store.append([base_name, current_sheet_for_test, input_data.copy()])

        if pin_28V > 1:
            card_28V = str(int(card_28V) + 1)
            Initial_card_IP_28V = increment_last_octet(Initial_card_IP_28V)
            pin_28V = 1
        if pin_5V > 1:
            card_5V = f"{int(card_5V) + 1:04d}"
            Initial_card_IP_5V = increment_last_octet(Initial_card_IP_5V)
            pin_5V = 1
        if pin_Lv1 > 1:
            card_Lv1 = f"{int(card_Lv1) + 1:02d}"
            Initial_card_IP_Lv1 = increment_last_octet(Initial_card_IP_Lv1)
            pin_Lv1 = 1
        if pin_Data > 1:
            card_Data = f"{int(card_Data) + 1:02d}"
            Initial_card_IP_Data = increment_last_octet(Initial_card_IP_Data)
            pin_Data = 1
        if pin_Digital > 1:
            card_Digital = f"{int(card_Digital) + 1:02d}"
            Initial_card_IP_Digital = increment_last_octet(Initial_card_IP_Digital)
            pin_Digital = 1
        if pin_Analog > 1:
            card_Analog = f"{int(card_Analog) + 1:02d}"
            Initial_card_IP_Analog = increment_last_octet(Initial_card_IP_Analog)
            pin_Analog = 1
        if pin_Therm > 1:
            card_Therm = f"{int(card_Therm) + 1:02d}"
            Initial_card_IP_Therm = increment_last_octet(Initial_card_IP_Therm)
            pin_Therm = 1
        if pin_RF > 1:
            card_RF = f"{int(card_RF) + 1:02d}"
            Initial_card_IP_RF = increment_last_octet(Initial_card_IP_RF)
            pin_RF = 1

        if len(input_files) > 1:
            print('\nCompleted Processing Workbook')
        else:
            print(f'\nCompleted Processing Workbook "{input_file_name}". Advancing baseline card counters for next excel file.')

    output_master_workbook = os.path.join(output_folder, "MASTER_SUMMARY_TCTM.xlsx")

    if os.path.isfile(output_master_workbook):
        try:
            os.remove(output_master_workbook)
        except Exception as e:
            print(f"Warning: Could not remove old master summary file: {e}")

    with pd.ExcelWriter(output_master_workbook, engine='openpyxl') as writer:
        if not master_summary_tc.empty:
            master_summary_tc['TC_CARD'] = master_summary_tc['TC_CARD'].astype(str)
            master_summary_tc['ASSIGNED_TC_PIN'] = master_summary_tc['ASSIGNED_TC_PIN'].astype(str)
            if 'FUNCTIONAL_DESCRIPTION' in master_summary_tc.columns:
                master_summary_tc['FUNCTIONAL_DESCRIPTION'] = master_summary_tc['FUNCTIONAL_DESCRIPTION'].astype(str)
            master_summary_tc = master_summary_tc.sort_values(by=['TC_CARD', 'ASSIGNED_TC_PIN'])
            master_summary_tc.to_excel(writer, sheet_name="master_summary_tc", index=False)

        if not master_summary_tm.empty:
            master_summary_tm['TM_CARD'] = master_summary_tm['TM_CARD'].astype(str)
            master_summary_tm['ASSIGNED_TM_PIN'] = master_summary_tm['ASSIGNED_TM_PIN'].astype(str)
            if 'TM_FUNCTIONAL_DESCRIPTION' in master_summary_tm.columns:
                master_summary_tm['TM_FUNCTIONAL_DESCRIPTION'] = master_summary_tm['TM_FUNCTIONAL_DESCRIPTION'].astype(str)
            master_summary_tm = master_summary_tm.sort_values(by=['TM_CARD', 'ASSIGNED_TM_PIN'])
            master_summary_tm.to_excel(writer, sheet_name="master_summary_tm", index=False)

    print(f'\n>>> Combined TC & TM Files Saved Successfully to: {output_master_workbook}')

    unique_card_tc = sorted(master_summary_tc['TC_CARD'].dropna().unique().tolist()) if not master_summary_tc.empty else []
    unique_card_tm = sorted(master_summary_tm['TM_CARD'].dropna().unique().tolist()) if not master_summary_tm.empty else []

    summary_rows = []
    for uc in unique_card_tc:
        matching_rows = master_summary_tc[master_summary_tc['TC_CARD'] == uc]
        pins = pd.to_numeric(matching_rows['ASSIGNED_TC_PIN'], errors='coerce').dropna()
        if not pins.empty:
            summary_rows.append([uc, f"{int(pins.min()):02d}", f"{int(pins.max()):02d}"])
    for uc in unique_card_tm:
        matching_rows = master_summary_tm[master_summary_tm['TM_CARD'] == uc]
        pins = pd.to_numeric(matching_rows['ASSIGNED_TM_PIN'], errors='coerce').dropna()
        if not pins.empty:
            summary_rows.append([uc, f"{int(pins.min()):02d}", f"{int(pins.max()):02d}"])

    if summary_rows:
        hardware_summary_table = pd.DataFrame(summary_rows, columns=['TC_TM_CARD', 'START_PIN', 'END_PIN']).sort_values(by=['TC_TM_CARD'])
    else:
        hardware_summary_table = pd.DataFrame(columns=['TC_TM_CARD', 'START_PIN', 'END_PIN'])

    with pd.ExcelWriter(output_master_workbook, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        hardware_summary_table.to_excel(writer, sheet_name="CARD_PIN_SUMMARY", index=False)

    pdf_path = None
    if generate_pdf:
        pdf_path = _generate_pdf(output_folder, pdf_data_store, hardware_summary_table)

    return {
        "master_summary_path": output_master_workbook,
        "pdf_path": pdf_path,
        "tc_count": len(master_summary_tc),
        "tm_count": len(master_summary_tm),
        "unassigned": unassigned_list,
    }


def _generate_pdf(output_folder, pdf_data_store, hardware_summary_table):
    """Bugs fixed vs original: undefined usable_height, missing comma in TableStyle, misindented if-block."""
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, LongTable, PageBreak, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from datetime import date

    output_pdf_file = os.path.join(output_folder, "final_op_assigned.pdf")
    doc = SimpleDocTemplate(output_pdf_file, pagesize=letter, leftMargin=21, rightMargin=21, topMargin=14, bottomMargin=14)
    usable_height = letter[1] - doc.topMargin - doc.bottomMargin

    styles = getSampleStyleSheet()
    elements = []

    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], alignment=1, fontSize=18, textColor=colors.HexColor('#8B0000'))
    sub_style = ParagraphStyle('SubStyle', parent=styles['Heading2'], alignment=1, fontSize=12, textColor=colors.HexColor('#696969'))

    elements.append(Spacer(1, usable_height / 2 - 50))
    elements.append(Paragraph("TCTM SIMULATOR ASSIGNMENT SUMMARY", title_style))
    elements.append(Spacer(1, 20))
    elements.append(Paragraph(f"Generation Date: {date.today()}", sub_style))
    elements.append(PageBreak())

    for idx, item in enumerate(pdf_data_store):
        wb_base_name, sheet_name_str, current_data = item
        if current_data is None or len(current_data) == 0:
            continue

        headers = list(current_data.columns)
        data_matrix = [headers] + current_data.astype(str).values.tolist()

        cell_style = ParagraphStyle('CellStyle', fontSize=4.5, alignment=1, wordWrap="CJK", leading=5)
        formatted_matrix = [[Paragraph(html.escape(str(cell)), cell_style) for cell in row] for row in data_matrix]

        num_cols = max(len(headers), 1)
        col_width = doc.width / num_cols
        calculated_col_widths = [col_width] * num_cols

        table = LongTable(formatted_matrix, repeatRows=1, splitByRow=True, colWidths=calculated_col_widths, rowHeights=[19] * len(data_matrix))
        table.setStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgreen),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ])

        elements.append(Paragraph(f"{wb_base_name} | {sheet_name_str}", styles['Normal']))
        elements.append(Spacer(1, 5))
        elements.append(table)
        if idx != len(pdf_data_store) - 1:
            elements.append(PageBreak())

    if not hardware_summary_table.empty:
        final_title_style = ParagraphStyle('FinalTitle', parent=styles['Heading1'], alignment=1, fontSize=12, textColor=colors.HexColor('#006400'))
        elements.append(PageBreak())
        elements.append(Paragraph("CARD PIN ALLOCATION SUMMARY", final_title_style))
        elements.append(Spacer(1, 5))

        sum_headers = list(hardware_summary_table.columns)
        sum_matrix = [sum_headers] + hardware_summary_table.astype(str).values.tolist()

        cell_style = ParagraphStyle('SummaryCellStyle', fontSize=6, alignment=1, leading=7, spaceBefore=0, spaceAfter=0)
        formatted_sum_matrix = [[Paragraph(html.escape(str(cell)), cell_style) for cell in row] for row in sum_matrix]

        t_sum = LongTable(
            formatted_sum_matrix, hAlign='CENTER',
            colWidths=[doc.width / len(sum_headers)] * len(sum_headers),
            rowHeights=[0.25 * 72] * len(sum_matrix), splitByRow=True
        )
        t_sum.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E0FFFF')),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ('FONTSIZE', (0, 0), (-1, -1), 7),
        ]))
        elements.append(t_sum)

    doc.build(elements)
    print("PDF CREATED:", output_pdf_file)
    return output_pdf_file
