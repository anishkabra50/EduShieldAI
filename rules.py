def get_section(roll_no):
    if not roll_no:
        return ""
    parts = str(roll_no).split('_')
    if len(parts) > 1:
        return parts[0]
    return ""

def get_table_i(students, use_pouring=True):
    """
    Table I: Students having aggregate attendance less than 75%.
    Columns needed: Exam Seat No, Name, Aggregate Attendance (%)
    """
    table_data = []
    for s in students:
        # Determine aggregate percentage
        pct = s['overall_pct']
        if pct is not None and pct < 75.0:
            table_data.append({
                'Exam Seat No': s['seat_no'],
                'Name': s['student_name'],
                'Aggregate Attendance\n(%)': f"{pct}%" if pct is not None else "N/A",
                'raw_pct': pct
            })
    # Sort by Seat No
    table_data.sort(key=lambda x: x['Exam Seat No'])
    return table_data

def get_table_ii(students, use_pouring=True):
    """
    Table II: Students having aggregate attendance between 60% and 75%.
    Columns needed: Exam Seat No, Student name, OverAll Attendance
    """
    table_data = []
    for s in students:
        pct = s['overall_pct']
        if pct is not None and 60.0 <= pct < 75.0:
            table_data.append({
                'Exam Seat No': s['seat_no'],
                'Student name': s['student_name'],
                'OverAll Attendance': s['overall_attendance_raw'],
                'raw_pct': pct
            })
    # Sort by Seat No
    table_data.sort(key=lambda x: x['Exam Seat No'])
    return table_data

def get_table_iii_a(students, subject_cols, course_map, use_pouring=True):
    """
    Table III(A): Course-wise detention.
    For each course, list the students detained (subject attendance < 75%).
    Columns: SN, Course Code, Course Name, Sec/Div., Exam Seat No, Name, Attendance (%)
    """
    # Group by course
    course_detentions = {}
    
    # Initialize course detentions lists
    for sub in subject_cols:
        course_detentions[sub] = []
        
    for s in students:
        sec = get_section(s['roll_no'])
        for sub in subject_cols:
            if sub in s['subjects']:
                sub_info = s['subjects'][sub]
                pct = sub_info['pouring_pct'] if use_pouring else sub_info['normal_pct']
                attended = sub_info['pouring_attended'] if use_pouring else sub_info['normal_attended']
                total = sub_info['pouring_total'] if use_pouring else sub_info['normal_total']
                
                if pct is not None and pct < 75.0:
                    course_detentions[sub].append({
                        'Sec/Div.': sec,
                        'Exam Seat No': s['seat_no'],
                        'Name': s['student_name'],
                        'Attendance (%)': f"{attended}/{total} ({pct}%)" if (attended is not None and total is not None) else f"{pct}%",
                        'raw_pct': pct
                    })
                    
    # Format into a flat list with serial numbers
    flat_table = []
    sn_counter = 1
    
    for sub in sorted(subject_cols):
        records = course_detentions[sub]
        if not records:
            continue
            
        # Get code and name from map, default to sub abbreviation
        c_code = course_map.get(sub, {}).get('code', sub)
        c_name = course_map.get(sub, {}).get('name', sub)
        
        # Sort records for this course by Exam Seat No
        records.sort(key=lambda x: x['Exam Seat No'])
        
        for idx, rec in enumerate(records):
            flat_table.append({
                'SN': str(sn_counter),
                'Course Code': c_code,
                'Course Name': c_name,
                'Sec/Div.': rec['Sec/Div.'],
                'Exam Seat No': rec['Exam Seat No'],
                'Name': rec['Name'],
                'Attendance (%)': rec['Attendance (%)']
            })
        
        sn_counter += 1
        
    return flat_table

def get_table_iii_b(students, subject_cols, course_map, use_pouring=True):
    """
    Table III(B): Student-wise detention.
    For each student, list all courses they are detained in.
    Columns: SN, Sec/Div, Exam Seat No, Student Name, Overall Attendance (%), Course Code, Course Name, Attendance (%)
    """
    student_detentions = []
    
    for s in students:
        sec = get_section(s['roll_no'])
        detained_courses = []
        
        for sub in subject_cols:
            if sub in s['subjects']:
                sub_info = s['subjects'][sub]
                pct = sub_info['pouring_pct'] if use_pouring else sub_info['normal_pct']
                attended = sub_info['pouring_attended'] if use_pouring else sub_info['normal_attended']
                total = sub_info['pouring_total'] if use_pouring else sub_info['normal_total']
                
                if pct is not None and pct < 75.0:
                    c_code = course_map.get(sub, {}).get('code', sub)
                    c_name = course_map.get(sub, {}).get('name', sub)
                    detained_courses.append({
                        'Course Code': c_code,
                        'Course Name': c_name,
                        'Attendance (%)': f"{attended}/{total} ({pct}%)" if (attended is not None and total is not None) else f"{pct}%"
                    })
                    
        if detained_courses:
            # Sort detained courses by Course Code
            detained_courses.sort(key=lambda x: x['Course Code'])
            student_detentions.append({
                'Sec/Div': sec,
                'Exam Seat No': s['seat_no'],
                'Student Name': s['student_name'],
                'Overall Attendance (%)': f"{s['overall_pct']}%" if s['overall_pct'] is not None else s['overall_attendance_raw'],
                'courses': detained_courses
            })
            
    # Sort students by Exam Seat No
    student_detentions.sort(key=lambda x: x['Exam Seat No'])
    
    # Flatten with repeated serial numbers for a single student's rows
    flat_table = []
    sn_counter = 1
    
    for item in student_detentions:
        for rec in item['courses']:
            flat_table.append({
                'SN': str(sn_counter),
                'Sec/\nDiv': item['Sec/Div'],
                'Exam Seat No': item['Exam Seat No'],
                'Student Name': item['Student Name'],
                'Overall Attendance (%)': item['Overall Attendance (%)'],
                'Course Code': rec['Course Code'],
                'Course Name': rec['Course Name'],
                'Attendance (%)': rec['Attendance (%)']
            })
        sn_counter += 1
        
    return flat_table
