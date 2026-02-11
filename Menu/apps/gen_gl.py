import streamlit as st
import pandas as pd
import zipfile
import io
import re
from datetime import datetime
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# =========================
# --- CONFIG (เหมือนเดิม) ---
# =========================

# source_folder = 'CDM'  # ในเวอร์ชัน ZIP จะใช้เป็น label เฉย ๆ
DEFAULT_SOURCE_FOLDER_NAME = "CDM"

tlf_reserved_rows = 2
gl_reserved_rows = 10
gap_rows = 3

exclude_tlf_columns = ['from_acct', 'to_acct', 'auth_branch_from']
TLF_LABEL = "TLF or Database(ATMI)"

gl_columns_letters = ['J', 'K', 'L', 'M', 'N', 'P', 'AM', 'AN', 'AZ']
gl_base_headers = ['RC', 'OC', 'CH', 'Product Code', 'Account Code', 'Tax', 'DR', 'CR', 'Seq', 'Details']
gl_new_headers = gl_base_headers

tlf_columns_letters = [
    'F', 'G', 'I', 'J', 'K', 'M', 'O', 'V',
    'AF', 'AS', 'AT', 'AU', 'AV', 'AX', 'AZ', 'CU', 'DP', 'BH'
]

# =========================
# --- Helper Functions (เหมือนเดิม) ---
# =========================

def excel_col_to_index(col_str):
    num = 0
    for c in col_str:
        if c in "0123456789":
            continue
        num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num - 1

def convert_implied_decimal(val):
    """
    แปลง implied decimal (หาร 100) เฉพาะกรณี:
    - เป็นเลขล้วน และ
    - มี "00" นำหน้า
    ถ้าไม่เข้าเงื่อนไข -> คืนค่าตรง ๆ (ไม่แปลง)
    """
    try:
        if val is None:
            return val
        val_str = str(val).strip()
        if not val_str.isdigit() or not val_str.startswith("00"):
            return val_str
        return float(val_str) / 100.0
    except:
        return val

def extract_seq_num(val):
    text = str(val)
    match = re.search(r'seq_num:(\d+)', text)
    if match:
        return match.group(1)
    return str(val).strip()

def strip_d_suffix_for_tlf_sheet(name_no_ext: str):
    return re.sub(r'[-_]?D\d{6}.*$', '', name_no_ext, flags=re.IGNORECASE).strip()

def make_unique_sheet_name(book, desired_name: str):
    base = (desired_name or "Sheet")[:31]
    name = base
    i = 2
    while name in book.sheetnames:
        suffix = f"_{i}"
        name = (base[:31 - len(suffix)] + suffix)[:31]
        i += 1
    return name

def max_k_from_searchkey(series: pd.Series) -> int:
    """
    หา max เลขท้ายของรูปแบบ xxxx|k จากคอลัมน์ _SearchKey
    ถ้าไม่เจอคืน 1
    """
    max_k = 1
    try:
        k_series = series.astype(str).str.extract(r'\|(\d+)\s*$')[0]
        k_series = pd.to_numeric(k_series, errors='coerce')
        max_k_val = k_series.max()
        if pd.notna(max_k_val):
            max_k = int(max_k_val)
    except:
        max_k = 1
    return max_k

def to_yymmdd(value):
    if value is None:
        return None

    if isinstance(value, (datetime, pd.Timestamp)):
        return value.strftime('%y%m%d')

    s = str(value).strip()
    if not s or s.lower() in ['nan', 'none', 'null']:
        return None

    digits = re.sub(r'[^0-9]', '', s)

    if len(digits) == 6:
        return digits

    if len(digits) == 8:
        first4 = int(digits[:4])
        last4 = int(digits[-4:])

        if 1900 <= first4 <= 2099:
            try:
                dt = datetime.strptime(digits, '%Y%m%d')
                return dt.strftime('%y%m%d')
            except:
                pass

        if 1900 <= last4 <= 2099:
            try:
                dt = datetime.strptime(digits, '%d%m%Y')
                return dt.strftime('%y%m%d')
            except:
                pass

    try:
        dt = pd.to_datetime(s, errors='coerce', dayfirst=True)
        if pd.isna(dt):
            return None
        return dt.strftime('%y%m%d')
    except:
        return None

# =========================
# ZIP Version: แทน pick_date_from_csv_col_c(file_path) ด้วย bytes
# (logic เดิมทุกประการ)
# =========================

def pick_date_from_csv_col_c_bytes(file_bytes: bytes, nrows=20):
    df_c = None
    for enc in ['utf-8', 'cp874']:
        try:
            bio = io.BytesIO(file_bytes)
            df_c = pd.read_csv(
                bio,
                header=None,
                usecols=[2],
                nrows=nrows,
                encoding=enc,
                dtype=str,
                engine='python'
            )
            break
        except:
            df_c = None

    if df_c is None or df_c.empty:
        return None

    yys = [to_yymmdd(v) for v in df_c.iloc[:, 0].tolist()]
    yys = [v for v in yys if v]
    if not yys:
        return None

    vc = pd.Series(yys).value_counts()
    max_count = vc.max()
    candidates = [d for d, c in vc.items() if c == max_count]
    return max(candidates, key=lambda x: int(x))

# =========================
# Indices / pos AZ CU (เหมือนเดิม)
# =========================

gl_indices = [excel_col_to_index(c) for c in gl_columns_letters]
tlf_indices = [excel_col_to_index(c) for c in tlf_columns_letters]

def get_col_pos_in_tlf(target_letter):
    sorted_letters = sorted(tlf_columns_letters, key=lambda x: excel_col_to_index(x))
    try:
        return sorted_letters.index(target_letter)
    except:
        return -1

pos_AZ = get_col_pos_in_tlf('AZ')
pos_CU = get_col_pos_in_tlf('CU')

# Styles (เหมือนเดิม)
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                     top=Side(style='thin'), bottom=Side(style='thin'))
align_center = Alignment(horizontal='center', vertical='center')
align_right = Alignment(horizontal='right', vertical='center')
align_left = Alignment(horizontal='left', vertical='center')
header_font = Font(bold=True)
title_font = Font(bold=True, size=14, color="000000")
search_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# =========================
# ✅ CORE: process_combined_data_from_zip
# = logic เดิม 100% แต่เปลี่ยน input เป็น zip bytes
# =========================

def process_combined_data_from_zip(zip_bytes: bytes, source_folder_name: str):
    output_filename = f'GL_{source_folder_name}.xlsx'

    log_lines = []

    def log(msg: str):
        log_lines.append(msg)

    with zipfile.ZipFile(io.BytesIO(zip_bytes), 'r') as z:
        names = [n for n in z.namelist() if not n.endswith('/')]

        # ✅ รอบนี้ไม่เป็น folder แล้ว: ทุกไฟล์อยู่ใน zip เดียวกัน
        # เลือกเฉพาะไฟล์ระดับเดียวกัน (กัน path)
        flat_names = [n for n in names if '/' not in n and '\\' not in n]

        tlf_files = [n for n in flat_names if n.lower().endswith('.xlsx')]
        csv_files = [n for n in flat_names if n.lower().endswith('.csv')]

        if not tlf_files:
            raise ValueError("ไม่พบไฟล์ TLF (.xlsx) ใน ZIP (ต้องอยู่ระดับเดียวกัน ไม่อยู่ใน folder)")
        if not csv_files:
            raise ValueError("ไม่พบไฟล์ .csv ใน ZIP (ต้องอยู่ระดับเดียวกัน ไม่อยู่ใน folder)")

        csv_files.sort(key=lambda p: p.lower())
        tlf_files.sort(key=lambda p: p.lower())

        log(f"กำลังประมวลผล... (Auto Width & Reorder Column)")
        log(f"Source folder: {source_folder_name}")
        log(f"TLF files: {tlf_files}")
        log(f"CSV files: {csv_files}")

        # ✅ Load หลายไฟล์ TLF (เหมือนเดิม แต่จาก bytes)
        tlf_books = []
        for f in tlf_files:
            try:
                tlf_bytes = z.read(f)
                book = pd.ExcelFile(io.BytesIO(tlf_bytes))
                tlf_books.append((f, book))
            except Exception as e:
                raise RuntimeError(f"Error อ่านไฟล์ TLF {f}: {e}")

        log("-" * 30)
        log("ไฟล์ที่จะประมวลผล (หา date จาก Column C1-C20 แยกไฟล์):")
        csv_bytes_map = {}
        for fn in csv_files:
            b = z.read(fn)
            csv_bytes_map[fn] = b
            d = pick_date_from_csv_col_c_bytes(b, nrows=20)
            log(f" - {fn} | chosen_date(YYMMDD)={d}")
        log("-" * 30)

        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            for filename in csv_files:
                file_bytes = csv_bytes_map[filename]

                chosen_date = pick_date_from_csv_col_c_bytes(file_bytes, nrows=20)

                desired_sheet_name = chosen_date if chosen_date else re.sub(r'\.csv$', '', filename, flags=re.IGNORECASE)
                desired_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', desired_sheet_name)

                log(f">> Processing: {filename} -> Sheet: {desired_sheet_name}")

                clean_name = re.sub(r'GL', '', filename, flags=re.IGNORECASE)
                clean_name = clean_name.replace('.csv', '').replace('.CSV', '').strip()
                fallback_lookup_name = strip_d_suffix_for_tlf_sheet(clean_name)

                tlf_lookup_candidates = []
                if chosen_date:
                    tlf_lookup_candidates.append(chosen_date)
                    tlf_lookup_candidates.append("D" + chosen_date)
                tlf_lookup_candidates.append(fallback_lookup_name)

                tlf_sheet_to_use = None
                tlf_book_to_use = None
                tlf_book_filename = None

                for book_fn, book in tlf_books:
                    for cand in tlf_lookup_candidates:
                        if cand and cand in book.sheet_names:
                            tlf_sheet_to_use = cand
                            tlf_book_to_use = book
                            tlf_book_filename = book_fn
                            break
                    if tlf_sheet_to_use:
                        break

                try:
                    # --- 1. Load Data ---
                    # TLF
                    tlf_df = pd.DataFrame()
                    effective_tlf_reserved_rows = tlf_reserved_rows
                    max_k_tlf = 1

                    if tlf_sheet_to_use:
                        tlf_df = pd.read_excel(tlf_book_to_use, sheet_name=tlf_sheet_to_use, usecols=tlf_indices, dtype=str)
                        for col in tlf_df.columns:
                            tlf_df[col] = tlf_df[col].astype(str).str.strip()

                        if pos_AZ != -1 and pos_AZ < len(tlf_df.columns):
                            tlf_df.iloc[:, pos_AZ] = tlf_df.iloc[:, pos_AZ].apply(convert_implied_decimal)
                        if pos_CU != -1 and pos_CU < len(tlf_df.columns):
                            tlf_df.iloc[:, pos_CU] = tlf_df.iloc[:, pos_CU].apply(convert_implied_decimal)

                        if not tlf_df.empty and len(tlf_df.columns) > 8:
                            search_col = tlf_df.iloc[:, 8].astype(str).str.strip()
                            tlf_df['_SearchKey'] = search_col + '|' + (tlf_df.groupby(search_col).cumcount() + 1).astype(str)

                            max_k_tlf = max_k_from_searchkey(tlf_df['_SearchKey'])
                            effective_tlf_reserved_rows = max(tlf_reserved_rows, max_k_tlf)

                        log(f"   ✓ TLF sheet used: {tlf_sheet_to_use} (file: {tlf_book_filename})")
                    else:
                        log(f"   ! ไม่พบชีตใน TLF (ลองแล้ว: {tlf_lookup_candidates}) -> ข้ามส่วน TLF")

                    # ATMI (GL) - อ่านจาก bytes แต่พารามิเตอร์เดิม
                    try:
                        gl_df = pd.read_csv(
                            io.BytesIO(file_bytes),
                            header=None,
                            usecols=gl_indices,
                            encoding='utf-8',
                            dtype=str,
                            engine='python'
                        )
                    except:
                        gl_df = pd.read_csv(
                            io.BytesIO(file_bytes),
                            header=None,
                            usecols=gl_indices,
                            encoding='cp874',
                            dtype=str,
                            engine='python'
                        )

                    gl_source_headers = ['RC', 'OC', 'CH', 'Product Code', 'Account Code', 'Tax', 'DR', 'CR', 'AZ_RAW']
                    if len(gl_df.columns) == len(gl_source_headers):
                        gl_df.columns = gl_source_headers

                    gl_df['Details'] = gl_df['AZ_RAW']
                    gl_df['Seq'] = gl_df['AZ_RAW'].apply(extract_seq_num).astype(str).str.strip()

                    gl_df['RC'] = gl_df['RC'].astype(str).str.strip()
                    gl_df['CH'] = gl_df['CH'].astype(str).str.strip()
                    gl_df['DR'] = pd.to_numeric(gl_df['DR'], errors='coerce').fillna(0)
                    gl_df['CR'] = pd.to_numeric(gl_df['CR'], errors='coerce').fillna(0)

                    gl_df = gl_df[gl_base_headers]

                    cols_to_sort = ['CH', 'RC', 'OC', 'Product Code']
                    gl_df = gl_df.sort_values(by=cols_to_sort, ascending=[True, True, True, True])

                    if not gl_df.empty:
                        search_col_gl = gl_df['Seq'].astype(str)
                        gl_df['_SearchKey'] = search_col_gl + '|' + (gl_df.groupby(search_col_gl).cumcount() + 1).astype(str)

                    max_k_gl = 1
                    if not gl_df.empty and '_SearchKey' in gl_df.columns:
                        max_k_gl = max_k_from_searchkey(gl_df['_SearchKey'])
                    effective_gl_reserved_rows = max(gl_reserved_rows, max_k_gl)

                    # --- 2. Setup Layout ---
                    target_sheet_name = make_unique_sheet_name(writer.book, desired_sheet_name)

                    worksheet = writer.book.create_sheet(target_sheet_name)
                    writer.sheets[target_sheet_name] = worksheet
                    ws = writer.sheets[target_sheet_name]

                    search_ui_start_row = 1
                    tlf_ui_height = 2 + (effective_tlf_reserved_rows if not tlf_df.empty else 0)
                    gl_ui_height = 2 + (effective_gl_reserved_rows if not gl_df.empty else 0)
                    raw_data_start_row = search_ui_start_row + tlf_ui_height + gap_rows + gl_ui_height + 5

                    current_raw_row = raw_data_start_row

                    # Store ranges
                    tlf_data_start = tlf_data_end = None
                    tlf_key_col_letter = 'A'
                    gl_data_start = gl_data_end = None
                    gl_key_col_letter = 'A'

                    # --- 3. Write Raw Data ---
                    if not tlf_df.empty:
                        ws.cell(row=current_raw_row - 1, column=1, value=TLF_LABEL).font = Font(bold=True, italic=True)
                        tlf_df.to_excel(writer, sheet_name=target_sheet_name, startrow=current_raw_row - 1, index=False)
                        tlf_data_start = current_raw_row + 1
                        tlf_data_end = current_raw_row + len(tlf_df)
                        tlf_key_col_letter = get_column_letter(len(tlf_df.columns))

                        for row in range(current_raw_row, tlf_data_end + 1):
                            for col in range(1, len(tlf_df.columns)):
                                cell = ws.cell(row=row, column=col)
                                cell.border = thin_border
                                if row == current_raw_row:
                                    cell.alignment = align_center
                                    cell.font = header_font
                                else:
                                    cell.alignment = align_right if isinstance(cell.value, (int, float)) else align_center
                                    if col == 9:
                                        cell.number_format = '@'

                        current_raw_row += len(tlf_df) + 4

                    if not gl_df.empty:
                        ws.cell(row=current_raw_row - 1, column=1, value="--- Raw ATMI Data ---").font = Font(bold=True, italic=True)
                        gl_df.to_excel(writer, sheet_name=target_sheet_name, startrow=current_raw_row - 1, index=False)
                        gl_data_start = current_raw_row + 1
                        gl_data_end = current_raw_row + len(gl_df)
                        gl_key_col_letter = get_column_letter(len(gl_df.columns))

                        for row in range(current_raw_row, gl_data_end + 1):
                            for col in range(1, len(gl_df.columns) + 1):
                                cell = ws.cell(row=row, column=col)
                                cell.border = thin_border
                                if row == current_raw_row:
                                    cell.alignment = align_center
                                    cell.font = header_font
                                else:
                                    if col in [7, 8]:
                                        cell.alignment = align_right
                                        cell.number_format = '#,##0.00'
                                    elif col == gl_base_headers.index('Details') + 1:
                                        cell.alignment = align_left
                                        cell.number_format = '@'
                                    else:
                                        cell.alignment = align_center

                                    if col in [gl_base_headers.index('Seq') + 1, gl_base_headers.index('Details') + 1]:
                                        cell.number_format = '@'

                    # --- 4. Search UI ---
                    ws[f'A{search_ui_start_row}'] = "🔍 ค้นหาข้อมูล SEQ"
                    ws[f'A{search_ui_start_row}'].font = Font(bold=True, size=12)
                    ws[f'A{search_ui_start_row}'].alignment = Alignment(horizontal='right')

                    input_cell_ref = f'$B${search_ui_start_row}'
                    input_cell = ws[f'B{search_ui_start_row}']
                    input_cell.fill = search_fill
                    input_cell.border = thin_border
                    input_cell.alignment = align_center
                    input_cell.number_format = '@'

                    report_row = search_ui_start_row + 2

                    # 4.2 TLF Result
                    if not tlf_df.empty:
                        ws[f'A{report_row}'] = TLF_LABEL
                        ws[f'A{report_row}'].font = title_font

                        display_cols = [c for c in tlf_df.columns if c != '_SearchKey' and c not in exclude_tlf_columns]

                        if 'amt_1_full' in display_cols and 'resp_byte' in display_cols:
                            idx1 = display_cols.index('amt_1_full')
                            idx2 = display_cols.index('resp_byte')
                            display_cols[idx1], display_cols[idx2] = display_cols[idx2], display_cols[idx1]

                        current_col_idx = 1
                        tlf_key_range_str = f"${tlf_key_col_letter}${tlf_data_start}:${tlf_key_col_letter}${tlf_data_end}"

                        for col_name in display_cols:
                            cell = ws.cell(row=report_row + 1, column=current_col_idx)
                            cell.value = col_name
                            cell.font = Font(bold=True)
                            cell.border = thin_border
                            cell.alignment = align_center
                            cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                            current_col_idx += 1

                        data_start_row = report_row + 2
                        for r_offset in range(effective_tlf_reserved_rows):
                            current_formula_row = data_start_row + r_offset
                            k_value = r_offset + 1
                            match_logic = f'MATCH({input_cell_ref}&"|"&{k_value}, {tlf_key_range_str}, 0)'

                            for i, col_name in enumerate(display_cols, 1):
                                original_col_idx = tlf_df.columns.get_loc(col_name)
                                col_letter = get_column_letter(original_col_idx + 1)
                                data_col_range = f"${col_letter}${tlf_data_start}:${col_letter}${tlf_data_end}"
                                formula = f'=IFERROR(INDEX({data_col_range}, {match_logic}), "")'

                                cell = ws.cell(row=current_formula_row, column=i)
                                cell.value = formula
                                cell.border = thin_border
                                cell.alignment = align_center

                        report_row = data_start_row + effective_tlf_reserved_rows

                    report_row += gap_rows

                    # 4.3 ATMI Result
                    if not gl_df.empty:
                        ws[f'A{report_row}'] = "ATMI"
                        ws[f'A{report_row}'].font = title_font

                        current_col_idx = 1
                        for col_name in gl_df.columns:
                            if col_name == '_SearchKey':
                                continue
                            cell = ws.cell(row=report_row + 1, column=current_col_idx)
                            cell.value = col_name
                            cell.font = Font(bold=True)
                            cell.border = thin_border
                            cell.alignment = align_center
                            cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                            current_col_idx += 1

                        data_start_row = report_row + 2
                        gl_key_range_str = f"${gl_key_col_letter}${gl_data_start}:${gl_key_col_letter}${gl_data_end}"
                        gl_display_cols = [c for c in gl_df.columns if c != '_SearchKey']

                        for r_offset in range(effective_gl_reserved_rows):
                            current_formula_row = data_start_row + r_offset
                            k_value = r_offset + 1
                            match_logic = f'MATCH({input_cell_ref}&"|"&{k_value}, {gl_key_range_str}, 0)'

                            for out_col_idx, col_name in enumerate(gl_display_cols, 1):
                                original_idx = gl_df.columns.get_loc(col_name) + 1
                                col_letter = get_column_letter(original_idx)
                                data_col_range = f"${col_letter}${gl_data_start}:${col_letter}${gl_data_end}"

                                formula = f'=IFERROR(INDEX({data_col_range}, {match_logic}), "")'

                                cell = ws.cell(row=current_formula_row, column=out_col_idx)
                                cell.value = formula
                                cell.border = thin_border

                                if col_name in ['DR', 'CR']:
                                    cell.number_format = '#,##0.00'
                                    cell.alignment = align_right
                                elif col_name == 'Details':
                                    cell.number_format = '@'
                                    cell.alignment = align_left
                                else:
                                    cell.alignment = align_center
                                    if col_name == 'Seq':
                                        cell.number_format = '@'

                    # --- 5. Smart Auto Width ---
                    col_widths = {}

                    def update_max_width(df, start_col_idx=1, skip_cols=None):
                        skip_cols = set(skip_cols or [])
                        for i, col_name in enumerate(df.columns):
                            if col_name in skip_cols:
                                continue
                            current_idx = start_col_idx + i
                            max_len = len(str(col_name))
                            if not df.empty:
                                try:
                                    data_len = df[col_name].astype(str).map(len).max()
                                    if pd.notna(data_len):
                                        max_len = max(max_len, data_len)
                                except:
                                    pass
                            existing = col_widths.get(current_idx, 0)
                            col_widths[current_idx] = max(existing, max_len + 3)

                    if not tlf_df.empty:
                        update_max_width(tlf_df, start_col_idx=1)
                        if 'display_cols' in locals():
                            for idx, col_name in enumerate(display_cols, 1):
                                header_len = len(str(col_name)) + 3
                                existing = col_widths.get(idx, 0)
                                col_widths[idx] = max(existing, header_len)

                    if not gl_df.empty:
                        update_max_width(gl_df, start_col_idx=1, skip_cols={'Details'})

                    for col_idx, width in col_widths.items():
                        col_letter = get_column_letter(col_idx)
                        final_width = max(12, min(width, 60))
                        writer.sheets[target_sheet_name].column_dimensions[col_letter].width = final_width

                    writer.sheets[target_sheet_name].column_dimensions['A'].width = max(col_widths.get(1, 20), 30)
                    writer.sheets[target_sheet_name].column_dimensions['B'].width = max(col_widths.get(2, 20), 25)

                    if 'Details' in gl_df.columns:
                        details_col_idx = gl_df.columns.get_loc('Details') + 1
                        details_col_letter = get_column_letter(details_col_idx)
                        writer.sheets[target_sheet_name].column_dimensions[details_col_letter].width = 12

                    log(
                        f"   ✓ เสร็จสิ้น: {filename} -> {target_sheet_name} | "
                        f"chosen_date={chosen_date} | "
                        f"TLF max|k={max_k_tlf} ui_rows={effective_tlf_reserved_rows} | "
                        f"ATMI max|k={max_k_gl} ui_rows={effective_gl_reserved_rows}"
                    )

                except Exception as e:
                    log(f"X Error ไฟล์ {filename}: {e}")
                    import traceback
                    log(traceback.format_exc())

            if 'Sheet' in writer.book.sheetnames and len(writer.book.sheetnames) > 1:
                del writer.book['Sheet']

        out.seek(0)
        log("-" * 30)
        log(f"บันทึกไฟล์เรียบร้อยที่: {output_filename}")

        return out, output_filename, "\n".join(log_lines)

# =========================
# Streamlit UI
# =========================

st.write("อัปโหลด ZIP ที่มีไฟล์ .csv + .xlsx อยู่ **ระดับเดียวกันใน zip** (ไม่มี folder ย่อย)")

source_name = st.text_input("Source name (ใช้ตั้งชื่อไฟล์ output: GL_<source>.xlsx)", value=DEFAULT_SOURCE_FOLDER_NAME)
uploaded = st.file_uploader("Upload ZIP", type=["zip"])

if uploaded is not None:
    try:
        out_bytes, out_name, logs = process_combined_data_from_zip(uploaded.getvalue(), source_name)

        st.success("✅ ประมวลผลเสร็จสิ้น")
        st.download_button(
            label=f"⬇️ ดาวน์โหลดไฟล์ {out_name}",
            data=out_bytes.getvalue(),
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        with st.expander("ดู Log"):
            st.code(logs)

    except Exception as e:
        st.error(f"❌ Error: {e}")
