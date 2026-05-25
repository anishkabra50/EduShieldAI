import re
import pandas as pd
import xlrd
import openpyxl
import io

def parse_attendance_cell(text):
    """
    Parses a cell text like '25 / 35 ( 71.43 % ) ', 'pouring Attendance :28.0/ 33 (84.85) ',
    '134/256 (52.34)', or '0 / 7' and returns (attended, total, percentage).
    """
    text = str(text).strip()
    if not text or text.lower() == 'nan':
        return None, None, None
    
    # Try to find a fraction like "28.0/ 33" or "25 / 35"
    fraction_match = re.search(r"([\d\.]+)\s*/\s*([\d\.]+)", text)
    if fraction_match:
        attended = float(fraction_match.group(1))
        total = float(fraction_match.group(2))
        pct = (attended / total * 100) if total > 0 else 0.0
        
        # If there's an explicit percentage in parentheses, use it for exact precision
        pct_match = re.search(r"\(([\d\.]+)\s*%?\)", text)
        if pct_match:
            pct = float(pct_match.group(1))
        return attended, total, round(pct, 2)
    
    # Try to just find a percentage like "75%" or "75.5"
    pct_match = re.search(r"([\d\.]+)\s*%", text)
    if pct_match:
        pct = float(pct_match.group(1))
        return None, None, round(pct, 2)
        
    return None, None, None

def parse_attendance_sheet(file_bytes, file_name):
    """
    Parses an uploaded attendance file (.xls or .xlsx).
    Handles the 2-row per student block layout.
    Returns (students_list, subject_columns).
    """
    # Load workbook using xlrd for .xls and openpyxl for .xlsx
    if file_name.endswith('.xls'):
        book = xlrd.open_workbook(file_contents=file_bytes)
        sheet = book.sheet_by_index(0)
        nrows = sheet.nrows
        ncols = sheet.ncols
        
        # Get cell values
        def get_val(r, c):
            return str(sheet.cell_value(r, c)).strip()
    else:
        # xlsx
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        sheet = wb.active
        nrows = sheet.max_row
        ncols = sheet.max_column
        
        def get_val(r, c):
            # openpyxl is 1-indexed
            val = sheet.cell(row=r+1, column=c+1).value
            return str(val).strip() if val is not None else ""

    if nrows < 1:
        return [], []

    # Extract headers from row 0
    headers = [get_val(0, c) for c in range(ncols)]
    
    # Identify metadata columns (case-insensitive)
    roll_col = next((i for i, h in enumerate(headers) if "roll" in h.lower()), -1)
    id_col = next((i for i, h in enumerate(headers) if "unique" in h.lower() or "email" in h.lower()), -1)
    seat_col = next((i for i, h in enumerate(headers) if "seat" in h.lower() or "urn" in h.lower() or "prn" in h.lower()), -1)
    name_col = next((i for i, h in enumerate(headers) if "name" in h.lower()), -1)
    overall_col = next((i for i, h in enumerate(headers) if "overall" in h.lower() or "aggregate" in h.lower()), -1)
    
    # Ensure critical columns are found; if not, fall back to indices
    if roll_col == -1: roll_col = 0
    if id_col == -1: id_col = 1
    if seat_col == -1: seat_col = 2
    if name_col == -1: name_col = 3
    if overall_col == -1: overall_col = 4
    
    metadata_indices = {roll_col, id_col, seat_col, name_col, overall_col}
    
    # Subject columns are those that have a name and are not metadata
    subject_cols = {headers[i]: i for i in range(ncols) if i not in metadata_indices and headers[i] != ""}
    subject_names = list(subject_cols.keys())
    
    students = []
    current_student = None
    
    for r in range(1, nrows):
        # Check if row is completely empty
        is_row_empty = True
        for c in range(ncols):
            if get_val(r, c) != "":
                is_row_empty = False
                break
                
        if is_row_empty:
            current_student = None
            continue
            
        roll_val = get_val(r, roll_col)
        name_val = get_val(r, name_col)
        
        if roll_val or name_val:
            # Row A (main student details + pouring/adjusted attendance)
            current_student = {
                'roll_no': roll_val,
                'unique_id': get_val(r, id_col),
                'seat_no': get_val(r, seat_col),
                'student_name': name_val,
                'overall_attendance_raw': get_val(r, overall_col),
                'subjects': {}
            }
            # Parse overall attendance
            att, tot, pct = parse_attendance_cell(current_student['overall_attendance_raw'])
            current_student['overall_attended'] = att
            current_student['overall_total'] = tot
            current_student['overall_pct'] = pct
            
            # Parse subjects for Row A
            for sub_name, col_idx in subject_cols.items():
                cell_val = get_val(r, col_idx)
                att_s, tot_s, pct_s = parse_attendance_cell(cell_val)
                if pct_s is not None:
                    current_student['subjects'][sub_name] = {
                        'pouring_attended': att_s,
                        'pouring_total': tot_s,
                        'pouring_pct': pct_s,
                        'normal_attended': att_s,  # Default to pouring
                        'normal_total': tot_s,
                        'normal_pct': pct_s
                    }
            students.append(current_student)
        else:
            # Row B (normal/raw attendance values)
            if current_student is not None:
                for sub_name, col_idx in subject_cols.items():
                    cell_val = get_val(r, col_idx)
                    att_s, tot_s, pct_s = parse_attendance_cell(cell_val)
                    if pct_s is not None:
                        if sub_name in current_student['subjects']:
                            current_student['subjects'][sub_name]['normal_attended'] = att_s
                            current_student['subjects'][sub_name]['normal_total'] = tot_s
                            current_student['subjects'][sub_name]['normal_pct'] = pct_s
                        else:
                            current_student['subjects'][sub_name] = {
                                'pouring_attended': None,
                                'pouring_total': None,
                                'pouring_pct': None,
                                'normal_attended': att_s,
                                'normal_total': tot_s,
                                'normal_pct': pct_s
                            }
                            
    return students, subject_names
