import os
import shutil
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd
import io

# ReportLab imports for PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfgen import canvas

# Ensure templates directory is configured and template is copied if present
def ensure_template():
    project_root = os.path.dirname(os.path.abspath(__file__))
    templates_dir = os.path.join(project_root, 'templates')
    os.makedirs(templates_dir, exist_ok=True)
    
    template_dest = os.path.join(templates_dir, 'detention_format.docx')
    template_src_root = os.path.join(project_root, 'RBU_DETENTION LIST FORMAT.docx')
    
    if os.path.exists(template_src_root) and not os.path.exists(template_dest):
        shutil.copy(template_src_root, template_dest)
        
    return template_dest

def set_cell_text(cell, text, bold=False, font_name="Calibri", font_size=10, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.text = text
    p.style.font.name = font_name
    p.style.font.size = Pt(font_size)
    if align:
        p.alignment = align
    for run in p.runs:
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.bold = bold

def fill_docx_template(programme, semester, exam_name, date_str, table1_data, table2_data, table3a_data, table3b_data, custom_template_path=None):
    """
    Loads templates/detention_format.docx or a custom template, replaces metadata, fills all 4 tables,
    and returns docx file as bytes.
    """
    if custom_template_path and os.path.exists(custom_template_path):
        template_path = custom_template_path
    else:
        template_path = ensure_template()
        
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found at {template_path}. Please place RBU_DETENTION LIST FORMAT.docx in root or templates/ folder.")
        
    doc = docx.Document(template_path)
    
    # 1. Replace placeholder text in paragraphs
    for p in doc.paragraphs:
        if "<Exam Name>" in p.text:
            p.text = p.text.replace("<Exam Name>", exam_name)
        if "Date: dd/mm/yyyy" in p.text:
            p.text = p.text.replace("dd/mm/yyyy", date_str)
        elif "Date:" in p.text and ("dd/mm/yyyy" in p.text or "___" in p.text):
            p.text = p.text.replace("dd/mm/yyyy", date_str)
            p.text = p.text.replace("___________________________", date_str)
            
        if "Programme:" in p.text and "Semester:" in p.text:
            # Format nicely
            p.text = f"Programme: {programme}               Semester: {semester}"
            
    # 2. Fill Table 0 (Table I - Aggregate < 75%)
    t0 = doc.tables[0]
    # Delete placeholders (from row index 1 onwards)
    while len(t0.rows) > 1:
        t0._tbl.remove(t0.rows[1]._tr)
        
    if not table1_data:
        # Nil row
        row = t0.add_row()
        set_cell_text(row.cells[0], "NIL", font_name="Calibri", font_size=10)
        set_cell_text(row.cells[1], "NIL", font_name="Calibri", font_size=10)
        set_cell_text(row.cells[2], "NIL", font_name="Calibri", font_size=10)
    else:
        for item in table1_data:
            row = t0.add_row()
            set_cell_text(row.cells[0], str(item.get('Exam Seat No', '')), font_name="Calibri")
            set_cell_text(row.cells[1], str(item.get('Name', '')), font_name="Calibri")
            set_cell_text(row.cells[2], str(item.get('Aggregate Attendance\n(%)', '')), font_name="Calibri")

    # 3. Fill Table 1 (Table II - Aggregate 60-75%)
    t1 = doc.tables[1]
    while len(t1.rows) > 1:
        t1._tbl.remove(t1.rows[1]._tr)
        
    if not table2_data:
        row = t1.add_row()
        set_cell_text(row.cells[0], "NIL", font_name="Calibri")
        set_cell_text(row.cells[1], "NIL", font_name="Calibri")
        set_cell_text(row.cells[2], "NIL", font_name="Calibri")
    else:
        for item in table2_data:
            row = t1.add_row()
            set_cell_text(row.cells[0], str(item.get('Exam Seat No', '')), font_name="Calibri")
            set_cell_text(row.cells[1], str(item.get('Student name', '')), font_name="Calibri")
            set_cell_text(row.cells[2], str(item.get('OverAll Attendance', '')), font_name="Calibri")

    # 4. Fill Table 2 (Table III(A) - Course-wise)
    t2 = doc.tables[2]
    # Header is rows 0 and 1, placeholders start at row 2
    while len(t2.rows) > 2:
        t2._tbl.remove(t2.rows[2]._tr)
        
    if not table3a_data:
        row = t2.add_row()
        for i in range(7):
            set_cell_text(row.cells[i], "NIL" if i < 3 else "", font_name="Calibri")
    else:
        # Keep track of courses for merging
        course_start_idx = 2
        current_course_code = None
        
        for idx, item in enumerate(table3a_data):
            row = t2.add_row()
            set_cell_text(row.cells[0], str(item.get('SN', '')), font_name="Calibri")
            set_cell_text(row.cells[1], str(item.get('Course Code', '')), font_name="Calibri")
            set_cell_text(row.cells[2], str(item.get('Course Name', '')), font_name="Calibri")
            set_cell_text(row.cells[3], str(item.get('Sec/Div.', '')), font_name="Calibri")
            set_cell_text(row.cells[4], str(item.get('Exam Seat No', '')), font_name="Calibri")
            set_cell_text(row.cells[5], str(item.get('Name', '')), font_name="Calibri")
            set_cell_text(row.cells[6], str(item.get('Attendance (%)', '')), font_name="Calibri")
            
            # Merging logic
            c_code = item.get('Course Code', '')
            if current_course_code is None:
                current_course_code = c_code
                course_start_idx = len(t2.rows) - 1
            elif current_course_code != c_code:
                # Merge previous course rows
                course_end_idx = len(t2.rows) - 2
                if course_end_idx > course_start_idx:
                    for col in [0, 1, 2]:
                        t2.cell(course_start_idx, col).merge(t2.cell(course_end_idx, col))
                current_course_code = c_code
                course_start_idx = len(t2.rows) - 1
                
        # Final merge
        course_end_idx = len(t2.rows) - 1
        if course_end_idx > course_start_idx:
            for col in [0, 1, 2]:
                t2.cell(course_start_idx, col).merge(t2.cell(course_end_idx, col))

    # 5. Fill Table 3 (Table III(B) - Student-wise)
    t3 = doc.tables[3]
    # Header is row 0, placeholders start at row 1
    while len(t3.rows) > 1:
        t3._tbl.remove(t3.rows[1]._tr)
        
    if not table3b_data:
        row = t3.add_row()
        for i in range(8):
            set_cell_text(row.cells[i], "NIL" if i < 4 else "", font_name="Calibri")
    else:
        student_start_idx = 1
        current_student_seat = None
        
        for idx, item in enumerate(table3b_data):
            row = t3.add_row()
            set_cell_text(row.cells[0], str(item.get('SN', '')), font_name="Calibri")
            set_cell_text(row.cells[1], str(item.get('Sec/\nDiv', '')), font_name="Calibri")
            set_cell_text(row.cells[2], str(item.get('Exam Seat No', '')), font_name="Calibri")
            set_cell_text(row.cells[3], str(item.get('Student Name', '')), font_name="Calibri")
            set_cell_text(row.cells[4], str(item.get('Overall Attendance (%)', '')), font_name="Calibri")
            set_cell_text(row.cells[5], str(item.get('Course Code', '')), font_name="Calibri")
            set_cell_text(row.cells[6], str(item.get('Course Name', '')), font_name="Calibri")
            set_cell_text(row.cells[7], str(item.get('Attendance (%)', '')), font_name="Calibri")
            
            # Merging logic
            seat = item.get('Exam Seat No', '')
            if current_student_seat is None:
                current_student_seat = seat
                student_start_idx = len(t3.rows) - 1
            elif current_student_seat != seat:
                # Merge previous student rows
                student_end_idx = len(t3.rows) - 2
                if student_end_idx > student_start_idx:
                    for col in [0, 1, 2, 3, 4]:
                        t3.cell(student_start_idx, col).merge(t3.cell(student_end_idx, col))
                current_student_seat = seat
                student_start_idx = len(t3.rows) - 1
                
        # Final merge
        student_end_idx = len(t3.rows) - 1
        if student_end_idx > student_start_idx:
            for col in [0, 1, 2, 3, 4]:
                t3.cell(student_start_idx, col).merge(t3.cell(student_end_idx, col))

    # Save to memory buffer
    docx_io = io.BytesIO()
    doc.save(docx_io)
    docx_io.seek(0)
    return docx_io.getvalue()

def generate_excel(table1_data, table2_data, table3a_data, table3b_data):
    """
    Creates a styled Excel file with 4 sheets representing the 4 tables.
    Returns excel file as bytes.
    """
    excel_io = io.BytesIO()
    
    with pd.ExcelWriter(excel_io, engine='openpyxl') as writer:
        # Table 1
        df1 = pd.DataFrame(table1_data)
        if 'raw_pct' in df1.columns:
            df1 = df1.drop(columns=['raw_pct'])
        if df1.empty:
            df1 = pd.DataFrame(columns=['Exam Seat No', 'Name', 'Aggregate Attendance (%)'])
        df1.to_excel(writer, sheet_name='Table I - Under 75%', index=False)
        
        # Table 2
        df2 = pd.DataFrame(table2_data)
        if 'raw_pct' in df2.columns:
            df2 = df2.drop(columns=['raw_pct'])
        if df2.empty:
            df2 = pd.DataFrame(columns=['Exam Seat No', 'Student name', 'OverAll Attendance'])
        df2.to_excel(writer, sheet_name='Table II - 60-75%', index=False)
        
        # Table 3A
        df3a = pd.DataFrame(table3a_data)
        if df3a.empty:
            df3a = pd.DataFrame(columns=['SN', 'Course Code', 'Course Name', 'Sec/Div.', 'Exam Seat No', 'Name', 'Attendance (%)'])
        df3a.to_excel(writer, sheet_name='Table IIIA - Course-wise', index=False)
        
        # Table 3B
        df3b = pd.DataFrame(table3b_data)
        if df3b.empty:
            df3b = pd.DataFrame(columns=['SN', 'Sec/Div', 'Exam Seat No', 'Student Name', 'Overall Attendance (%)', 'Course Code', 'Course Name', 'Attendance (%)'])
        df3b.to_excel(writer, sheet_name='Table IIIB - Student-wise', index=False)
        
    excel_io.seek(0)
    return excel_io.getvalue()

# NumberedCanvas for PDF pagination
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            super().showPage()
        super().save()

    def draw_page_number(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#4A5568"))
        page_text = f"Page {self._pageNumber} of {page_count}"
        self.drawRightString(612 - 54, 30, page_text)
        self.drawString(54, 30, "EduShield Attendance Automation System")
        self.setStrokeColor(colors.HexColor("#CBD5E0"))
        self.setLineWidth(0.5)
        self.line(54, 42, 612 - 54, 42)
        self.restoreState()

def generate_pdf(programme, semester, exam_name, date_str, table1_data, table2_data, table3a_data, table3b_data):
    """
    Generates a beautifully styled ReportLab PDF containing metadata and all 4 tables.
    Returns PDF file as bytes.
    """
    pdf_io = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_io,
        pagesize=letter,
        leftMargin=54,
        rightMargin=54,
        topMargin=54,
        bottomMargin=54
    )
    
    styles = getSampleStyleSheet()
    
    # Custom Styles for Premium Aesthetics
    title_style = ParagraphStyle(
        'UnivTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=15,
        textColor=colors.HexColor("#1A365D"), # Dark Blue
        alignment=1, # Center
        spaceAfter=6
    )
    
    subtitle_style = ParagraphStyle(
        'UnivSubtitle',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=12,
        textColor=colors.HexColor("#2C5282"),
        alignment=1, # Center
        spaceAfter=15
    )
    
    meta_style = ParagraphStyle(
        'MetaText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        textColor=colors.HexColor("#2D3748"),
        leading=14
    )
    
    section_title = ParagraphStyle(
        'SectionTitle',
        parent=styles['Heading3'],
        fontName='Helvetica-Bold',
        fontSize=11,
        textColor=colors.HexColor("#2B6CB0"),
        spaceBefore=14,
        spaceAfter=6
    )
    
    body_style = ParagraphStyle(
        'BodyTextCustom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        textColor=colors.HexColor("#2D3748"),
        leading=13,
        spaceAfter=8
    )
    
    th_style = ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        textColor=colors.white,
        alignment=1
    )
    
    td_style = ParagraphStyle(
        'TableCell',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        textColor=colors.HexColor("#1A202C"),
        alignment=1
    )

    t_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor("#F7FAFC"), colors.white]),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ])

    story = []
    
    # Title & Header
    story.append(Paragraph("RAMDEOBABA UNIVERSITY, NAGPUR", title_style))
    story.append(Paragraph("DETENTION LIST", subtitle_style))
    story.append(Paragraph(f"<b>Exam Name:</b> {exam_name}", meta_style))
    story.append(Paragraph(f"<b>Date:</b> {date_str}", meta_style))
    story.append(Paragraph(f"<b>Programme:</b> {programme} &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>Semester:</b> {semester}", meta_style))
    story.append(Spacer(1, 10))
    story.append(Paragraph("Submitted to Vice-Chancellor for Approval through Dean Academics", meta_style))
    story.append(Spacer(1, 10))
    
    # Table I
    story.append(Paragraph("Table I: Aggregate Attendance &lt; 75%", section_title))
    story.append(Paragraph("The following students have aggregate attendance less than 75%. As per the university ordinances, they are recommended to be detained in the examination in all courses.", body_style))
    
    t1_headers = [Paragraph("Exam Seat No", th_style), Paragraph("Student Name", th_style), Paragraph("Aggregate Attendance", th_style)]
    t1_rows = [t1_headers]
    if not table1_data:
        t1_rows.append([Paragraph("NIL", td_style), Paragraph("NIL", td_style), Paragraph("NIL", td_style)])
    else:
        for x in table1_data:
            t1_rows.append([
                Paragraph(str(x.get('Exam Seat No', '')), td_style),
                Paragraph(str(x.get('Name', '')), td_style),
                Paragraph(str(x.get('Aggregate Attendance\n(%)', '')), td_style)
            ])
    
    tab1 = Table(t1_rows, colWidths=[120, 264, 120])
    tab1.setStyle(t_style)
    story.append(tab1)
    story.append(Spacer(1, 15))
    
    # Table II
    story.append(Paragraph("Table II: Condonation list (Attendance 60% - 75%)", section_title))
    story.append(Paragraph("The following students have aggregate attendance between 60% and 75% and have applied for condonation of attendance. Their application forms and medical certificates/relevant documents have been verified.", body_style))
    
    t2_headers = [Paragraph("Exam Seat No", th_style), Paragraph("Student Name", th_style), Paragraph("Overall Attendance", th_style)]
    t2_rows = [t2_headers]
    if not table2_data:
        t2_rows.append([Paragraph("NIL", td_style), Paragraph("NIL", td_style), Paragraph("NIL", td_style)])
    else:
        for x in table2_data:
            t2_rows.append([
                Paragraph(str(x.get('Exam Seat No', '')), td_style),
                Paragraph(str(x.get('Student name', '')), td_style),
                Paragraph(str(x.get('OverAll Attendance', '')), td_style)
            ])
            
    tab2 = Table(t2_rows, colWidths=[120, 264, 120])
    tab2.setStyle(t_style)
    story.append(tab2)
    
    # Table IIIA on next page
    story.append(PageBreak())
    story.append(Paragraph("Table III(A): Course-wise Detention", section_title))
    story.append(Paragraph("The following is the course-wise list of students recommended to be detained in specific courses due to course-specific attendance being less than 75%.", body_style))
    
    t3a_headers = [
        Paragraph("SN", th_style),
        Paragraph("Course Code", th_style),
        Paragraph("Course Name", th_style),
        Paragraph("Sec/Div", th_style),
        Paragraph("Seat No", th_style),
        Paragraph("Student Name", th_style),
        Paragraph("Attendance", th_style)
    ]
    t3a_rows = [t3a_headers]
    if not table3a_data:
        t3a_rows.append([Paragraph("NIL", td_style), Paragraph("NIL", td_style), Paragraph("NIL", td_style), Paragraph("", td_style), Paragraph("", td_style), Paragraph("", td_style), Paragraph("", td_style)])
    else:
        for x in table3a_data:
            t3a_rows.append([
                Paragraph(str(x.get('SN', '')), td_style),
                Paragraph(str(x.get('Course Code', '')), td_style),
                Paragraph(str(x.get('Course Name', '')), td_style),
                Paragraph(str(x.get('Sec/Div.', '')), td_style),
                Paragraph(str(x.get('Exam Seat No', '')), td_style),
                Paragraph(str(x.get('Name', '')), td_style),
                Paragraph(str(x.get('Attendance (%)', '')), td_style)
            ])
            
    tab3a = Table(t3a_rows, colWidths=[24, 70, 110, 40, 70, 120, 70])
    tab3a.setStyle(t_style)
    story.append(tab3a)
    
    # Table IIIB on next page
    story.append(PageBreak())
    story.append(Paragraph("Table III(B): Student-wise Detention", section_title))
    story.append(Paragraph("The following is the student-wise list of courses in which they are recommended to be detained due to attendance in those courses being less than 75%.", body_style))
    
    t3b_headers = [
        Paragraph("SN", th_style),
        Paragraph("Sec/Div", th_style),
        Paragraph("Seat No", th_style),
        Paragraph("Student Name", th_style),
        Paragraph("Overall %", th_style),
        Paragraph("Course Code", th_style),
        Paragraph("Course Name", th_style),
        Paragraph("Attendance", th_style)
    ]
    t3b_rows = [t3b_headers]
    if not table3b_data:
        t3b_rows.append([Paragraph("NIL", td_style), Paragraph("NIL", td_style), Paragraph("NIL", td_style), Paragraph("NIL", td_style), Paragraph("", td_style), Paragraph("", td_style), Paragraph("", td_style), Paragraph("", td_style)])
    else:
        for x in table3b_data:
            t3b_rows.append([
                Paragraph(str(x.get('SN', '')), td_style),
                Paragraph(str(x.get('Sec/\nDiv', '')), td_style),
                Paragraph(str(x.get('Exam Seat No', '')), td_style),
                Paragraph(str(x.get('Student Name', '')), td_style),
                Paragraph(str(x.get('Overall Attendance (%)', '')), td_style),
                Paragraph(str(x.get('Course Code', '')), td_style),
                Paragraph(str(x.get('Course Name', '')), td_style),
                Paragraph(str(x.get('Attendance (%)', '')), td_style)
            ])
            
    tab3b = Table(t3b_rows, colWidths=[24, 40, 70, 110, 50, 70, 90, 50])
    tab3b.setStyle(t_style)
    story.append(tab3b)
    
    # Build Document using NumberedCanvas for dynamic total page count
    doc.build(story, canvasmaker=NumberedCanvas)
    
    pdf_io.seek(0)
    return pdf_io.getvalue()
