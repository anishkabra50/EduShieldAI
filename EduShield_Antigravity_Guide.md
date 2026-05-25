# EduShield Attendance Automation Tool - Build Guide for Antigravity

## Goal

Build a system that:

Input:
- Attendance Excel sheets
- Programme
- Semester
- Exam Name
- Number of attendance sheets

Output:
- Table I (Overall attendance <75%)
- Table II (Attendance 60%-75%)
- Table III(A) Course-wise detention
- Table III(B) Student-wise detention
- Export DOCX / Excel / PDF

---

## Tech Stack

Frontend:
- Streamlit

Backend:
- Python
- Pandas
- OpenPyXL
- python-docx

Optional:
- Plotly
- ReportLab

Install:

```bash
pip install streamlit pandas openpyxl python-docx plotly
```

---

## Project Structure

```text
EduShield/
│
├── app.py
├── parser.py
├── rules.py
├── generators.py
├── export.py
├── templates/
│   └── detention_format.docx
├── uploads/
├── outputs/
└── requirements.txt
```

---

## Module 1: Upload System

Inputs:
1. Programme
2. Semester
3. Exam Name
4. Number of Sheets
5. Upload XLS/XLSX

Streamlit:

```python
uploaded = st.file_uploader(
    "Upload Attendance",
    type=["xls","xlsx"]
)
```

---

## Module 2: Attendance Parser

Read Excel:

```python
import pandas as pd

df = pd.read_excel(uploaded)
```

Extract:

Example:

134/256 (52.34)

Need:

```python
attendance = 52.34
```

Regex:

```python
import re

def extract_percent(text):

    match = re.search(
        r"\(([\d\.]+)",
        str(text)
    )

    if match:
        return float(
            match.group(1)
        )

    return None
```

Apply to:
- Overall attendance
- Subject attendance
- Labs
- Practicals

---

## Module 3: Rule Engine

### Table I

Condition:

overall < 75

```python
table1 = df[
df["Overall"] < 75
]
```

### Table II

Condition:

60 <= overall < 75

```python
table2 = df[
(df["Overall"] >= 60)
&
(df["Overall"] < 75)
]
```

### Table III(A)

For each subject:

```python
if subject_attendance < 75:
    add_to_course_detention()
```

Store:
- Course code
- Course name
- Student
- Attendance

### Table III(B)

Store:

- Student
- Overall %
- Course
- Subject attendance

---

## Module 4: Report Generation

Generate:

- Table I
- Table II
- Table III(A)
- Table III(B)

Export:
- DOCX
- Excel
- PDF

Use:

```python
from docx import Document

doc = Document()

doc.add_heading(
"DETENTION LIST"
)

doc.save(
"output.docx"
)
```

---

## UI Flow

Upload Sheet

↓

Parser

↓

Attendance Extraction

↓

Rule Engine

↓

Generate Tables

↓

Export

---

## Tuesday Demo Flow

1. Upload O3 attendance file
2. Enter programme
3. Enter semester
4. Click Generate
5. Show Table I
6. Show Table II
7. Show Table III(A)
8. Show Table III(B)
9. Export report

---

## Antigravity Prompt

Paste into Antigravity:

Build a Streamlit application named EduShield.

Features:
1. Upload xls/xlsx attendance sheets
2. Input programme semester exam name number of sheets
3. Parse attendance values from text patterns
4. Generate Table I for attendance below 75
5. Generate Table II for attendance between 60 and 75
6. Generate Table III(A) course-wise detention
7. Generate Table III(B) student-wise detention
8. Export DOCX Excel PDF
9. Use modular architecture:
app.py parser.py rules.py generators.py export.py
10. Clean UI with dashboard cards and tables
11. Support multiple sheets
12. Use pandas and python-docx

Generate complete production-ready code.
