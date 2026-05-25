import streamlit as st
import datetime
import pandas as pd
from parser import parse_attendance_sheet
from rules import get_table_i, get_table_ii, get_table_iii_a, get_table_iii_b
from export import fill_docx_template, generate_excel, generate_pdf

# Set Page Config
st.set_page_config(
    page_title="EduShield - Attendance Automation",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Premium Custom CSS for Rich Aesthetics
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif;
    }
    
    /* Global Styles */
    .stApp {
        background: radial-gradient(circle at 50% 50%, #151828 0%, #090a10 100%);
        color: #f1f5f9;
    }
    
    /* Custom Scrollbar for Premium Feel */
    ::-webkit-scrollbar {
        width: 6px;
        height: 6px;
    }
    ::-webkit-scrollbar-track {
        background: rgba(15, 23, 42, 0.3);
    }
    ::-webkit-scrollbar-thumb {
        background: rgba(255, 255, 255, 0.08);
        border-radius: 4px;
    }
    ::-webkit-scrollbar-thumb:hover {
        background: rgba(255, 255, 255, 0.18);
    }
    
    /* Transparent Streamlit Default Header */
    header[data-testid="stHeader"] {
        background-color: transparent !important;
    }
    
    /* Sidebar styling with Glassmorphism */
    section[data-testid="stSidebar"] {
        background-color: rgba(9, 10, 16, 0.85) !important;
        backdrop-filter: blur(25px);
        border-right: 1px solid rgba(255, 255, 255, 0.05) !important;
    }
    
    /* Header Container styling - Premium Glassmorphic banner */
    .header-container {
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.4) 0%, rgba(15, 23, 42, 0.6) 100%);
        backdrop-filter: blur(20px);
        padding: 3rem 2.5rem;
        border-radius: 24px;
        margin-bottom: 2.5rem;
        border: 1px solid rgba(255, 255, 255, 0.05);
        box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
        position: relative;
        overflow: hidden;
    }
    
    .header-container::after {
        content: '';
        position: absolute;
        bottom: 0;
        left: 0;
        width: 100%;
        height: 5px;
        background: linear-gradient(90deg, #6366f1, #3b82f6, #10b981);
    }
    
    .header-title {
        font-size: 3.25rem;
        font-weight: 800;
        margin: 0;
        background: linear-gradient(90deg, #ffffff 0%, #cbd5e1 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        letter-spacing: -0.03em;
        display: flex;
        align-items: center;
        gap: 16px;
    }
    
    .header-subtitle {
        font-size: 1.15rem;
        color: #94a3b8;
        margin-top: 0.75rem;
        font-weight: 400;
        letter-spacing: 0.01em;
        max-width: 800px;
        line-height: 1.6;
    }
    
    /* File Uploader styling customisation */
    div[data-testid="stFileUploader"] {
        background-color: rgba(15, 23, 42, 0.3) !important;
        border: 2px dashed rgba(255, 255, 255, 0.08) !important;
        border-radius: 16px !important;
        padding: 1.5rem !important;
        transition: all 0.3s ease;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: rgba(59, 130, 246, 0.4) !important;
        background-color: rgba(15, 23, 42, 0.45) !important;
    }
    
    /* Dashboard Cards with Glowing Hover Effects */
    .card-container {
        display: flex;
        gap: 1.5rem;
        margin-bottom: 2.5rem;
    }
    
    .kpi-card {
        flex: 1;
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.25) 0%, rgba(15, 23, 42, 0.45) 100%);
        border: 1px solid rgba(255, 255, 255, 0.04);
        border-radius: 20px;
        padding: 2rem 1.75rem;
        text-align: left;
        backdrop-filter: blur(20px);
        transition: all 0.4s cubic-bezier(0.16, 1, 0.3, 1);
        box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3);
    }
    
    .kpi-card:hover {
        transform: translateY(-8px);
        border-color: rgba(99, 102, 241, 0.3);
        box-shadow: 0 30px 45px -15px rgba(99, 102, 241, 0.25);
    }
    
    .kpi-value {
        font-size: 2.75rem;
        font-weight: 800;
        color: #ffffff;
        line-height: 1;
        margin-bottom: 0.6rem;
        letter-spacing: -0.02em;
    }
    
    .kpi-label {
        font-size: 0.8rem;
        font-weight: 600;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.1em;
    }
    
    /* Metric variations and color themes */
    .metric-detained .kpi-value {
        color: #f87171;
    }
    .metric-detained:hover {
        border-color: rgba(239, 68, 68, 0.3);
        box-shadow: 0 30px 45px -15px rgba(239, 68, 68, 0.25);
    }
    
    .metric-condone .kpi-value {
        color: #fbbf24;
    }
    .metric-condone:hover {
        border-color: rgba(245, 158, 11, 0.3);
        box-shadow: 0 30px 45px -15px rgba(245, 158, 11, 0.25);
    }
    
    .metric-passed .kpi-value {
        color: #34d399;
    }
    .metric-passed:hover {
        border-color: rgba(16, 185, 129, 0.3);
        box-shadow: 0 30px 45px -15px rgba(16, 185, 129, 0.25);
    }
    
    /* Custom tab buttons styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 16px;
        background-color: rgba(15, 23, 42, 0.4);
        padding: 8px;
        border-radius: 16px;
        border: 1px solid rgba(255, 255, 255, 0.05);
        margin-bottom: 2rem;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 48px;
        white-space: pre-wrap;
        background-color: transparent;
        border-radius: 12px;
        color: #94a3b8;
        font-weight: 600;
        font-size: 0.95rem;
        transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
        padding: 0 24px;
        border: none !important;
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        color: #f8fafc;
        background-color: rgba(255, 255, 255, 0.03);
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #6366f1 0%, #3b82f6 100%) !important;
        color: #ffffff !important;
        box-shadow: 0 10px 20px rgba(59, 130, 246, 0.35);
    }
    
    /* Info Message boxes */
    .stAlert {
        background-color: rgba(30, 41, 59, 0.35) !important;
        border: 1px solid rgba(255, 255, 255, 0.06) !important;
        color: #e2e8f0 !important;
        border-radius: 16px !important;
        box-shadow: 0 10px 25px rgba(0,0,0,0.2);
    }
    
    /* Premium Button & Download Button styles */
    div.stButton>button, div.stDownloadButton>button {
        background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%) !important;
        color: white !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        border: 1px solid rgba(255, 255, 255, 0.08) !important;
        padding: 0.75rem 2rem !important;
        border-radius: 14px !important;
        transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1) !important;
        box-shadow: 0 8px 20px rgba(37, 99, 235, 0.25) !important;
        width: 100% !important;
    }
    
    div.stButton>button:hover, div.stDownloadButton>button:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 15px 30px rgba(37, 99, 235, 0.4) !important;
        border-color: rgba(255, 255, 255, 0.2) !important;
    }
    
    /* Input field styling */
    .stTextInput input, .stSelectbox select, .stDateInput input {
        background-color: rgba(15, 23, 42, 0.45) !important;
        border: 1px solid rgba(255, 255, 255, 0.06) !important;
        border-radius: 14px !important;
        color: #f8fafc !important;
        padding: 0.6rem 1.1rem !important;
        transition: all 0.25s ease !important;
    }
    
    .stTextInput input:focus, .stSelectbox select:focus, .stDateInput input:focus {
        border-color: #3b82f6 !important;
        box-shadow: 0 0 0 2px rgba(59, 130, 246, 0.2) !important;
    }
    
    /* Highlight containers */
    .highlight-box {
        background-color: rgba(99, 102, 241, 0.03);
        border: 1px solid rgba(99, 102, 241, 0.08);
        padding: 1.5rem;
        border-radius: 16px;
        margin-bottom: 2rem;
    }
</style>
""", unsafe_allow_html=True)

# App Header
st.markdown("""
<div class="header-container">
    <div class="header-title">🛡️ EduShield Attendance System</div>
    <div class="header-subtitle">Automate attendance sheets parsing, verify university detention policies, and generate official reports.</div>
</div>
""", unsafe_allow_html=True)

# Initialize Session States
if 'parsed_data' not in st.session_state:
    st.session_state['parsed_data'] = None
if 'subject_cols' not in st.session_state:
    st.session_state['subject_cols'] = []
if 'course_map' not in st.session_state:
    st.session_state['course_map'] = {}
if 'students' not in st.session_state:
    st.session_state['students'] = []

# Sidebar Configuration & Uploads
with st.sidebar:
    st.markdown("### 📋 Configuration Panel")
    
    # School / Department
    school_options = [
        "School of Computer Science & Engineering",
        "School of Electrical & Electronics Engineering",
        "School of Mechanical Engineering",
        "School of Civil Engineering",
        "School of Management",
        "Custom (Type below)..."
    ]
    school_sel = st.selectbox("School / Department", options=school_options, index=0, help="Select or type school name")
    if school_sel == "Custom (Type below)...":
        school = st.text_input("Enter Custom School / Department", "")
    else:
        school = school_sel
        
    # Programme
    prog_options = [
        "B.Tech. Computer Science & Engineering",
        "B.Tech. Computer Science & Engineering (AIML)",
        "B.Tech. Information Technology",
        "B.Tech. Electronics & Communication Engineering",
        "Master of Computer Applications (MCA)",
        "Master of Business Administration (MBA)",
        "Custom (Type below)..."
    ]
    prog_sel = st.selectbox("Programme", options=prog_options, index=0, help="Select or type programme name")
    if prog_sel == "Custom (Type below)...":
        programme = st.text_input("Enter Custom Programme", "")
    else:
        programme = prog_sel
        
    # Semester
    sem_options = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "Custom (Type below)..."]
    sem_sel = st.selectbox("Semester", options=sem_options, index=3, help="Select or type semester")
    if sem_sel == "Custom (Type below)...":
        semester = st.text_input("Enter Custom Semester", "")
    else:
        semester = sem_sel
        
    # Exam Name
    exam_options = [
        "End Semester Theory Examination, Summer 2026",
        "End Semester Practical Examination, Summer 2026",
        "End Semester Theory Examination, Winter 2026",
        "End Semester Practical Examination, Winter 2026",
        "Make-up Examination, Summer 2026",
        "Custom (Type below)..."
    ]
    exam_sel = st.selectbox("Exam Name", options=exam_options, index=0, help="Select or type exam identifier")
    if exam_sel == "Custom (Type below)...":
        exam_name = st.text_input("Enter Custom Exam Name", "")
    else:
        exam_name = exam_sel
        
    # Report Date
    report_date = st.date_input("Report Date", datetime.date.today(), help="Select report publication date")
    
    st.markdown("---")
    st.markdown("### 📁 Upload Files")
    
    uploaded_files = st.file_uploader(
        "Upload files (Spreadsheets, DOCX, PPT, PDFs etc.)",
        type=None,  # Allow all types of files
        accept_multiple_files=True,
        help="Upload attendance sheets (.xls/.xlsx), word templates (.docx), presentations (.ppt/.pptx), or other supporting documents."
    )
    
    st.markdown("---")
    st.markdown("### ⚙️ Rule Settings")
    use_pouring = st.toggle("Use Adjusted/Pouring Attendance", value=True, 
                            help="If checked, uses provisional/pouring attendance values which match Overall Attendance. Uncheck to run rules against raw attendance.")

# Ensure uploads folder exists
import os
os.makedirs('uploads', exist_ok=True)

# Main logic flow
if uploaded_files:
    excel_files = []
    docx_files = []
    ppt_files = []
    other_files = []
    
    # Save all uploaded files to uploads/ and categorize them
    for file in uploaded_files:
        file_bytes = file.getvalue()
        file_path = os.path.join('uploads', file.name)
        with open(file_path, 'wb') as f:
            f.write(file_bytes)
            
        ext = os.path.splitext(file.name)[1].lower()
        if ext in ['.xls', '.xlsx']:
            excel_files.append((file_path, file.name))
        elif ext in ['.docx']:
            docx_files.append((file_path, file.name))
        elif ext in ['.ppt', '.pptx']:
            ppt_files.append((file_path, file.name))
        else:
            other_files.append((file_path, file.name))
            
    # Sidebar Template Selector if any DOCX template uploaded
    selected_template_path = None
    if docx_files:
        with st.sidebar:
            st.markdown("---")
            st.markdown("### 📝 Template Settings")
            template_options = ["Default System Template"] + [f[1] for f in docx_files]
            sel_option = st.selectbox(
                "Select Word Template",
                options=template_options,
                help="Choose which uploaded DOCX file to use as the template for detention reports."
            )
            if sel_option != "Default System Template":
                selected_template_path = next(f[0] for f in docx_files if f[1] == sel_option)
                
    # Sidebar File Hub to show uploaded items
    if docx_files or ppt_files or other_files:
        with st.sidebar:
            st.markdown("---")
            with st.expander("📁 Uploaded Files Hub", expanded=True):
                if docx_files:
                    st.markdown("**Templates (.docx):**")
                    for _, name in docx_files:
                        st.markdown(f"- 📄 {name}")
                if ppt_files:
                    st.markdown("**Presentations (.ppt/.pptx):**")
                    for _, name in ppt_files:
                        st.markdown(f"- 📊 {name}")
                if other_files:
                    st.markdown("**Supporting Files:**")
                    for _, name in other_files:
                        st.markdown(f"- 📎 {name}")

    # Compile students across multiple sheets if uploaded
    all_students = []
    all_subjects = set()
    
    # Parse Excel files
    for file_path, file_name in excel_files:
        try:
            with open(file_path, 'rb') as f:
                content = f.read()
            students, subjects = parse_attendance_sheet(content, file_name)
            all_students.extend(students)
            all_subjects.update(subjects)
        except Exception as e:
            st.error(f"Error parsing attendance sheet '{file_name}': {e}")
            
    if all_students:
        st.session_state['students'] = all_students
        st.session_state['subject_cols'] = sorted(list(all_subjects))
        
        # Course Map Builder UI in sidebar/expander
        st.markdown("### 📚 Subject & Course Codes Mapping")
        with st.expander("Configure Course Codes & Names", expanded=True):
            st.info("Map Excel abbreviations (e.g. CLA) to official University Course Codes (e.g. UECS204T) and Names.")
            
            # Map subjects
            for sub in st.session_state['subject_cols']:
                # Set defaults
                default_code = st.session_state['course_map'].get(sub, {}).get('code', sub)
                default_name = st.session_state['course_map'].get(sub, {}).get('name', sub)
                
                col1, col2 = st.columns(2)
                with col1:
                    code_val = st.text_input(f"[{sub}] Code", value=default_code, key=f"code_{sub}")
                with col2:
                    name_val = st.text_input(f"[{sub}] Name", value=default_name, key=f"name_{sub}")
                
                st.session_state['course_map'][sub] = {'code': code_val, 'name': name_val}

        # Rules Engine - Calculate Tables
        table1_data = get_table_i(st.session_state['students'], use_pouring=use_pouring)
        table2_data = get_table_ii(st.session_state['students'], use_pouring=use_pouring)
        table3a_data = get_table_iii_a(st.session_state['students'], st.session_state['subject_cols'], st.session_state['course_map'], use_pouring=use_pouring)
        table3b_data = get_table_iii_b(st.session_state['students'], st.session_state['subject_cols'], st.session_state['course_map'], use_pouring=use_pouring)

        # Unique student counts for KPIs
        total_students = len(st.session_state['students'])
        detained_all = len(table1_data)
        condonation_list = len(table2_data)
        course_detained_unique = len(set(x['Exam Seat No'] for x in table3b_data))
        
        # KPI Cards HTML Display
        st.markdown(f"""
        <div class="card-container">
            <div class="kpi-card">
                <div class="kpi-value">{total_students}</div>
                <div class="kpi-label">Total Students Parsed</div>
            </div>
            <div class="kpi-card metric-detained">
                <div class="kpi-value">{detained_all}</div>
                <div class="kpi-label">Table I (Detained All Courses)</div>
            </div>
            <div class="kpi-card metric-condone">
                <div class="kpi-value">{condonation_list}</div>
                <div class="kpi-label">Table II (Condonation 60%-75%)</div>
            </div>
            <div class="kpi-card metric-passed">
                <div class="kpi-value">{course_detained_unique}</div>
                <div class="kpi-label">Unique Course-wise Detentions</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Action Bar / Export Panel
        st.markdown("### 📥 Download Reports")
        col_d1, col_d2, col_d3 = st.columns(3)
        
        date_str = report_date.strftime("%d/%m/%Y")
        
        # Generate files on demand
        try:
            docx_bytes = fill_docx_template(
                programme=programme,
                semester=semester,
                exam_name=exam_name,
                date_str=date_str,
                table1_data=table1_data,
                table2_data=table2_data,
                table3a_data=table3a_data,
                table3b_data=table3b_data,
                school=school,
                custom_template_path=selected_template_path if 'selected_template_path' in locals() else None
            )
            with col_d1:
                st.download_button(
                    label="📄 Download Official DOCX",
                    data=docx_bytes,
                    file_name=f"Detention_List_{programme}_{semester}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        except Exception as e:
            with col_d1:
                st.error(f"DOCX Generation Error: {e}")
                
        try:
            xls_bytes = generate_excel(
                table1_data=table1_data,
                table2_data=table2_data,
                table3a_data=table3a_data,
                table3b_data=table3b_data
            )
            with col_d2:
                st.download_button(
                    label="📊 Download Excel Summary",
                    data=xls_bytes,
                    file_name=f"Detention_Summary_{programme}_{semester}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            with col_d2:
                st.error(f"Excel Generation Error: {e}")
                
        try:
            pdf_bytes = generate_pdf(
                programme=programme,
                semester=semester,
                exam_name=exam_name,
                date_str=date_str,
                table1_data=table1_data,
                table2_data=table2_data,
                table3a_data=table3a_data,
                table3b_data=table3b_data,
                school=school
            )
            with col_d3:
                st.download_button(
                    label="📕 Download PDF Report",
                    data=pdf_bytes,
                    file_name=f"Detention_Report_{programme}_{semester}.pdf",
                    mime="application/pdf"
                )
        except Exception as e:
            with col_d3:
                st.error(f"PDF Generation Error: {e}")

        st.markdown("---")
        
        # Tabs for Tables View
        tab_t1, tab_t2, tab_t3a, tab_t3b, tab_full = st.tabs([
            "Table I (Detained Overall)", 
            "Table II (Condonation)", 
            "Table III(A) (Course-wise)", 
            "Table III(B) (Student-wise)",
            "Full Student View"
        ])
        
        with tab_t1:
            st.subheader("Table I: Aggregate attendance less than 75%")
            st.markdown("These students are detained in the examination in all courses.")
            df1_show = pd.DataFrame(table1_data)
            if not df1_show.empty:
                df1_show = df1_show.drop(columns=['raw_pct'], errors='ignore')
                st.dataframe(df1_show, use_container_width=True)
            else:
                st.info("No students found with aggregate attendance less than 75%.")
                
        with tab_t2:
            st.subheader("Table II: Aggregate attendance between 60% and 75%")
            st.markdown("These students have applied for condonation. If approved, they may be permitted to appear in the examinations.")
            df2_show = pd.DataFrame(table2_data)
            if not df2_show.empty:
                df2_show = df2_show.drop(columns=['raw_pct'], errors='ignore')
                st.dataframe(df2_show, use_container_width=True)
            else:
                st.info("No students found with aggregate attendance between 60% and 75%.")
                
        with tab_t3a:
            st.subheader("Table III(A): Course-wise Detention")
            st.markdown("List of courses along with students recommended to be detained due to subject-specific attendance less than 75%.")
            df3a_show = pd.DataFrame(table3a_data)
            if not df3a_show.empty:
                st.dataframe(df3a_show, use_container_width=True)
            else:
                st.info("No course-wise detentions found.")
                
        with tab_t3b:
            st.subheader("Table III(B): Student-wise Detention")
            st.markdown("List of students along with courses in which they are recommended to be detained.")
            df3b_show = pd.DataFrame(table3b_data)
            if not df3b_show.empty:
                st.dataframe(df3b_show, use_container_width=True)
            else:
                st.info("No student-wise course detentions found.")
                
        with tab_full:
            st.subheader("Parsed Student Directory")
            st.markdown("Complete database of all parsed student attendance records.")
            
            # Format student records as flat list for table display
            full_table = []
            for s in st.session_state['students']:
                row_item = {
                    'Roll No': s['roll_no'],
                    'Seat No': s['seat_no'],
                    'Student Name': s['student_name'],
                    'Overall Attendance': s['overall_attendance_raw'],
                    'Overall %': f"{s['overall_pct']}%" if s['overall_pct'] is not None else "N/A"
                }
                # Add subjects
                for sub in st.session_state['subject_cols']:
                    if sub in s['subjects']:
                        sub_info = s['subjects'][sub]
                        p_val = sub_info['pouring_pct']
                        n_val = sub_info['normal_pct']
                        if p_val != n_val:
                            row_item[sub] = f"{p_val}% (Adj) / {n_val}% (Raw)"
                        else:
                            row_item[sub] = f"{p_val}%"
                    else:
                        row_item[sub] = "N/A"
                full_table.append(row_item)
                
            st.dataframe(pd.DataFrame(full_table), use_container_width=True)

    else:
        st.warning("No students parsed from sheet. Please verify sheet formatting.")
else:
    # Landing / Upload Prompt
    st.info("👋 Welcome to EduShield! Please select configuration details in the sidebar and upload attendance sheets (.xls/.xlsx) to get started.")
    
    # Custom dashboard mock illustration using HTML
    st.markdown("""
    <div style="background-color: rgba(30, 41, 59, 0.4); border: 1px dashed #475569; padding: 3rem; border-radius: 16px; text-align: center; margin-top: 2rem;">
        <h3 style="color: #94a3b8; margin-bottom: 1rem;">No Data Uploaded</h3>
        <p style="color: #64748b; max-width: 500px; margin: 0 auto 1.5rem auto;">
            Upload your attendance Excel sheets via the sidebar to run the detention engine and generate official Word, PDF, and Excel reports.
        </p>
    </div>
    """, unsafe_allow_html=True)
