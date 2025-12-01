"""
Academic Results Dashboard Generator for Class 8
Generates Excel dashboard with student results, grades, and analytics
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import PieChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
import random
from datetime import datetime

# ============================================================================
# CONFIGURATION
# ============================================================================

SUBJECTS = {
    'Quran': 100,
    'English_I': 100,
    'English_II': 50,
    'Bangla_I': 100,
    'Bangla_II': 50,
    'Science': 100,
    'Mathematics': 100,
    'General_Knowledge': 100,
    'Arabic_I': 100,
    'Arabic_II': 100,
    'Aqaid': 100,
    'ICT': 50,
    'Bangladesh_and_Global_Studies': 100
}

TOTAL_MARKS = sum(SUBJECTS.values())  # 1200

# Pass thresholds (33% of full marks for each subject)
PASS_THRESHOLDS = {subject: marks * 0.33 for subject, marks in SUBJECTS.items()}

# Grade boundaries (percentage-based)
GRADE_BOUNDARIES = {
    'A+': 80,
    'A': 70,
    'A-': 60,
    'B': 50,
    'C': 40,
    'D': 33,
    'F': 0
}

# Grade Points
GRADE_POINTS = {
    'A+': 5.0,
    'A': 4.0,
    'A-': 3.5,
    'B': 3.0,
    'C': 2.0,
    'D': 1.0,
    'F': 0.0
}

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def calculate_grade(marks, full_marks):
    """Calculate grade based on percentage"""
    percentage = (marks / full_marks) * 100
    
    if percentage >= GRADE_BOUNDARIES['A+']:
        return 'A+'
    elif percentage >= GRADE_BOUNDARIES['A']:
        return 'A'
    elif percentage >= GRADE_BOUNDARIES['A-']:
        return 'A-'
    elif percentage >= GRADE_BOUNDARIES['B']:
        return 'B'
    elif percentage >= GRADE_BOUNDARIES['C']:
        return 'C'
    elif percentage >= GRADE_BOUNDARIES['D']:
        return 'D'
    else:
        return 'F'

def calculate_gp(grade):
    """Get grade point from grade"""
    return GRADE_POINTS[grade]

def is_subject_passed(marks, subject):
    """Check if student passed in a subject"""
    return marks >= PASS_THRESHOLDS[subject]

def calculate_overall_grade(gpa):
    """Calculate overall grade from GPA"""
    if gpa >= 5.0:
        return 'A+'
    elif gpa >= 4.0:
        return 'A'
    elif gpa >= 3.5:
        return 'A-'
    elif gpa >= 3.0:
        return 'B'
    elif gpa >= 2.0:
        return 'C'
    elif gpa >= 1.0:
        return 'D'
    else:
        return 'F'

# ============================================================================
# DATA GENERATION
# ============================================================================

def generate_sample_data(num_students=20):
    """Generate sample student data"""
    
    first_names = ['Ahmed', 'Fatima', 'Hassan', 'Aisha', 'Ali', 'Zainab', 'Omar', 'Maryam',
                   'Yusuf', 'Khadija', 'Ibrahim', 'Safiya', 'Bilal', 'Hafsa', 'Usman',
                   'Ruqayyah', 'Hamza', 'Asma', 'Khalid', 'Sumaya', 'Abdullah', 'Amina']
    
    last_names = ['Rahman', 'Khan', 'Ahmed', 'Ali', 'Hussain', 'Malik', 'Iqbal', 'Siddiqui',
                  'Hassan', 'Shah', 'Farooq', 'Aziz', 'Rashid', 'Nasir', 'Karim']
    
    data = []
    
    for i in range(num_students):
        student = {
            'Name': f"{random.choice(first_names)} {random.choice(last_names)}"
        }
        
        # Generate marks for each subject
        for subject, full_marks in SUBJECTS.items():
            # Generate realistic marks (mostly passing, some failing)
            if random.random() < 0.8:  # 80% likely to pass
                min_marks = PASS_THRESHOLDS[subject]
                marks = random.randint(int(min_marks), full_marks)
            else:  # 20% might fail
                marks = random.randint(0, full_marks)
            
            student[subject] = marks
        
        data.append(student)
    
    return pd.DataFrame(data)

# ============================================================================
# EXCEL SHEET CREATORS
# ============================================================================

def create_data_source_sheet(wb, df):
    """Create the main Data Source sheet with formulas"""
    ws = wb.create_sheet('Data Source', 0)
    
    # Title
    ws.merge_cells('A1:T1')
    title_cell = ws['A1']
    title_cell.value = 'CLASS 8 - ACADEMIC RESULTS DATA SOURCE'
    title_cell.font = Font(size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30
    
    # Headers (Row 2)
    headers = ['SL', 'Name'] + [name.replace('_', ' ') for name in SUBJECTS.keys()] + ['Total', 'Average', 'GPA', 'Grade', 'Position']
    
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # Add data and formulas for up to 1000 students
    for row in range(3, 1003):
        data_idx = row - 3  # Adjust for header rows
        
        # SL (A)
        if data_idx < len(df):
            ws[f'A{row}'].value = data_idx + 1
        else:
            ws[f'A{row}'] = f'=IF(B{row}="","",ROW()-2)'
        ws[f'A{row}'].alignment = Alignment(horizontal='center')
        
        # Name (B)
        if data_idx < len(df):
            ws[f'B{row}'].value = df.iloc[data_idx]['Name']
        
        # Subject marks (C to O) - 13 subjects
        if data_idx < len(df):
            for col_idx, subject in enumerate(SUBJECTS.keys(), 3):
                ws.cell(row=row, column=col_idx, value=df.iloc[data_idx][subject])
                ws.cell(row=row, column=col_idx).alignment = Alignment(horizontal='center')
        
        # Total marks (P = column 16)
        ws[f'P{row}'] = f'=IF(B{row}="","",SUM(C{row}:O{row}))'
        ws[f'P{row}'].alignment = Alignment(horizontal='center')
        ws[f'P{row}'].number_format = '0'
        
        # Average (Q = column 17)
        ws[f'Q{row}'] = f'=IF(B{row}="","",ROUND(P{row}/13,2))'
        ws[f'Q{row}'].alignment = Alignment(horizontal='center')
        ws[f'Q{row}'].number_format = '0.00'
        
        # Helper columns for individual subject GPs (hidden) - U to AG (21 to 33)
        subject_cols = list(range(3, 16))  # C to O (13 subjects)
        full_marks_list = list(SUBJECTS.values())
        
        for idx, (subj_col, subject, full_marks) in enumerate(zip(subject_cols, SUBJECTS.keys(), full_marks_list)):
            gp_col_num = 21 + idx  # Start from column U (21)
            if gp_col_num <= 26:
                gp_col = chr(64 + gp_col_num)
            else:
                gp_col = f'A{chr(64 + gp_col_num - 26)}'
            
            pass_threshold = PASS_THRESHOLDS[subject]
            
            # Subject GP formula
            ws[f'{gp_col}{row}'] = (
                f'=IF(B{row}="","",IF({chr(64+subj_col)}{row}<{pass_threshold},0,'
                f'IF({chr(64+subj_col)}{row}>={full_marks}*0.8,5,'
                f'IF({chr(64+subj_col)}{row}>={full_marks}*0.7,4,'
                f'IF({chr(64+subj_col)}{row}>={full_marks}*0.6,3.5,'
                f'IF({chr(64+subj_col)}{row}>={full_marks}*0.5,3,'
                f'IF({chr(64+subj_col)}{row}>={full_marks}*0.4,2,'
                f'IF({chr(64+subj_col)}{row}>={full_marks}*0.33,1,0))))))))'
            )
        
        # Check if failed in any subject (AH = column 34 - helper)
        ws[f'AH{row}'] = f'=IF(B{row}="","",IF(OR(U{row}=0,V{row}=0,W{row}=0,X{row}=0,Y{row}=0,Z{row}=0,AA{row}=0,AB{row}=0,AC{row}=0,AD{row}=0,AE{row}=0,AF{row}=0,AG{row}=0),1,0))'
        
        # GPA (R = column 18)
        ws[f'R{row}'] = (
            f'=IF(B{row}="","",IF(AH{row}=1,0,'
            f'MIN(5,ROUND((U{row}+V{row}+W{row}+X{row}+Y{row}+Z{row}+AA{row}+AB{row}+AC{row}+AD{row}+AE{row}+AF{row}+AG{row})/13,2))))'
        )
        ws[f'R{row}'].alignment = Alignment(horizontal='center')
        ws[f'R{row}'].number_format = '0.00'
        
        # Grade (S = column 19)
        ws[f'S{row}'] = (
            f'=IF(B{row}="","",IF(R{row}>=5,"A+",'
            f'IF(R{row}>=4,"A",IF(R{row}>=3.5,"A-",IF(R{row}>=3,"B",'
            f'IF(R{row}>=2,"C",IF(R{row}>=1,"D","F")))))))'
        )
        ws[f'S{row}'].alignment = Alignment(horizontal='center')
        
        # Position (T = column 20) - using COUNTIFS for consecutive ranking
        ws[f'T{row}'] = f'=IF(OR(B{row}="",S{row}="F"),"",COUNTIFS(P:P,">"&P{row},S:S,"<>F")+1)'
        ws[f'T{row}'].alignment = Alignment(horizontal='center')
    
    # Hide helper columns (U to AH for GPs and fail check)
    for col in range(21, 35):  # U to AH
        if col <= 26:
            col_letter = chr(64 + col)
        else:
            col_letter = f'A{chr(64 + col - 26)}'
        ws.column_dimensions[col_letter].hidden = True
    
    # Column widths
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 20
    for col in range(3, 16):  # C to O (subjects)
        ws.column_dimensions[chr(64 + col)].width = 10
    ws.column_dimensions['P'].width = 10
    ws.column_dimensions['Q'].width = 10
    ws.column_dimensions['R'].width = 10
    ws.column_dimensions['S'].width = 10
    ws.column_dimensions['T'].width = 10

def create_dashboard_sheet(wb):
    """Create Dashboard sheet with statistics and charts"""
    ws = wb.create_sheet('Dashboard')
    
    # Title
    ws.merge_cells('A1:H1')
    title_cell = ws['A1']
    title_cell.value = 'CLASS 8 - ACADEMIC DASHBOARD'
    title_cell.font = Font(size=18, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 35
    
    # Grade Distribution Section
    ws.merge_cells('A3:B3')
    ws['A3'] = 'GRADE DISTRIBUTION'
    ws['A3'].font = Font(size=12, bold=True)
    ws['A3'].fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    
    grades = ['A+', 'A', 'A-', 'B', 'C', 'D', 'F']
    for idx, grade in enumerate(grades, 4):
        ws[f'A{idx}'] = grade
        ws[f'A{idx}'].alignment = Alignment(horizontal='center')
        ws[f'B{idx}'] = f'=COUNTIF(\'Data Source\'!S:S,"{grade}")'
        ws[f'B{idx}'].alignment = Alignment(horizontal='center')
    
    # Top 5 Students Section
    ws.merge_cells('D3:F3')
    ws['D3'] = 'TOP 5 STUDENTS (BY TOTAL MARKS)'
    ws['D3'].font = Font(size=12, bold=True)
    ws['D3'].fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    
    ws['D4'] = 'Rank'
    ws['E4'] = 'Name'
    ws['F4'] = 'Total Marks'
    for cell in ['D4', 'E4', 'F4']:
        ws[cell].font = Font(bold=True)
        ws[cell].alignment = Alignment(horizontal='center')
    
    for rank in range(1, 6):
        row = rank + 4
        ws[f'D{row}'] = rank
        ws[f'D{row}'].alignment = Alignment(horizontal='center')
        
        # Get the nth highest total marks
        ws[f'F{row}'] = f'=IFERROR(LARGE(\'Data Source\'!$P:$P,{rank}),"")'
        ws[f'F{row}'].alignment = Alignment(horizontal='center')
        
        # Get name matching that total
        ws[f'E{row}'] = f'=IFERROR(INDEX(\'Data Source\'!$B:$B,MATCH(F{row},\'Data Source\'!$P:$P,0)),"")'
    
    # Statistics Section
    ws.merge_cells('A12:B12')
    ws['A12'] = 'STATISTICS'
    ws['A12'].font = Font(size=12, bold=True)
    ws['A12'].fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    
    stats = [
        ('Total Students', '=COUNTA(\'Data Source\'!B3:B1002)'),
        ('Pass Students', '=SUMPRODUCT(--(\'Data Source\'!S3:S1002<>"F"),--(\'Data Source\'!S3:S1002<>""))'),
        ('Fail Students', '=COUNTIF(\'Data Source\'!S3:S1002,"F")'),
        ('Pass Rate %', '=IF(B13=0,0,ROUND(B14/B13*100,2))'),
        ('Average GPA', '=ROUND(AVERAGEIF(\'Data Source\'!R3:R1002,">0"),2)'),
        ('Highest Total', '=MAX(\'Data Source\'!P3:P1002)'),
        ('Lowest Total', '=MIN(\'Data Source\'!P3:P1002)')
    ]
    
    for idx, (label, formula) in enumerate(stats, 13):
        ws[f'A{idx}'] = label
        ws[f'B{idx}'] = formula
        ws[f'B{idx}'].alignment = Alignment(horizontal='center')
    
    # Add Pie Chart
    pie = PieChart()
    labels = Reference(ws, min_col=1, min_row=4, max_row=10)
    data = Reference(ws, min_col=2, min_row=3, max_row=10)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.title = "Grade Distribution"
    ws.add_chart(pie, "D12")
    
    # Column widths
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['D'].width = 8
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 12

def create_subject_gpa_sheet(wb):
    """Create Subject-wise GPA sheet"""
    ws = wb.create_sheet('Subject-wise GPA')
    
    # Title
    ws.merge_cells('A1:P1')
    title_cell = ws['A1']
    title_cell.value = 'SUBJECT-WISE GPA'
    title_cell.font = Font(size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30
    
    # Headers
    headers = ['SL', 'Name'] + list(SUBJECTS.keys()) + ['Overall GPA', 'Grade']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=header.replace('_', ' '))
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # Data rows
    for row in range(3, 1003):
        # SL
        ws[f'A{row}'] = f'=IF(\'Data Source\'!B{row}="","",\'Data Source\'!A{row})'
        ws[f'A{row}'].alignment = Alignment(horizontal='center')
        
        # Name
        ws[f'B{row}'] = f'=IF(\'Data Source\'!B{row}="","",\'Data Source\'!B{row})'
        
        # Subject GPs (from hidden columns U to AG in Data Source - columns 21-33)
        gp_cols = ['U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG']
        for col_idx, gp_col in enumerate(gp_cols, 3):
            ws.cell(row=row, column=col_idx).value = f'=IF(\'Data Source\'!B{row}="","",\'Data Source\'!{gp_col}{row})'
            ws.cell(row=row, column=col_idx).alignment = Alignment(horizontal='center')
            ws.cell(row=row, column=col_idx).number_format = '0.0'
        
        # Overall GPA
        ws[f'P{row}'] = f'=IF(\'Data Source\'!B{row}="","",\'Data Source\'!R{row})'
        ws[f'P{row}'].alignment = Alignment(horizontal='center')
        ws[f'P{row}'].number_format = '0.00'
        
        # Grade
        ws[f'Q{row}'] = f'=IF(\'Data Source\'!B{row}="","",\'Data Source\'!S{row})'
        ws[f'Q{row}'].alignment = Alignment(horizontal='center')
    
    # Column widths
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 20
    for col in range(3, 16):
        ws.column_dimensions[chr(64 + col)].width = 10
    ws.column_dimensions['P'].width = 12
    ws.column_dimensions['Q'].width = 10

def create_subject_grade_sheet(wb):
    """Create Subject-wise Grade sheet"""
    ws = wb.create_sheet('Subject-wise Grade')
    
    # Title
    ws.merge_cells('A1:P1')
    title_cell = ws['A1']
    title_cell.value = 'SUBJECT-WISE GRADES'
    title_cell.font = Font(size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30
    
    # Headers
    headers = ['SL', 'Name'] + list(SUBJECTS.keys()) + ['Overall GPA', 'Grade']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=header.replace('_', ' '))
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # Data rows
    for row in range(3, 1003):
        # SL
        ws[f'A{row}'] = f'=IF(\'Data Source\'!B{row}="","",\'Data Source\'!A{row})'
        ws[f'A{row}'].alignment = Alignment(horizontal='center')
        
        # Name
        ws[f'B{row}'] = f'=IF(\'Data Source\'!B{row}="","",\'Data Source\'!B{row})'
        
        # Subject Grades (convert GP to grade from columns U to AG)
        gp_cols = ['U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG']
        for col_idx, gp_col in enumerate(gp_cols, 3):
            grade_formula = (
                f'=IF(\'Data Source\'!B{row}="","",IF(\'Data Source\'!{gp_col}{row}>=5,"A+",'
                f'IF(\'Data Source\'!{gp_col}{row}>=4,"A",IF(\'Data Source\'!{gp_col}{row}>=3.5,"A-",'
                f'IF(\'Data Source\'!{gp_col}{row}>=3,"B",IF(\'Data Source\'!{gp_col}{row}>=2,"C",'
                f'IF(\'Data Source\'!{gp_col}{row}>=1,"D","F")))))))'
            )
            ws.cell(row=row, column=col_idx).value = grade_formula
            ws.cell(row=row, column=col_idx).alignment = Alignment(horizontal='center')
        
        # Overall GPA
        ws[f'P{row}'] = f'=IF(\'Data Source\'!B{row}="","",\'Data Source\'!R{row})'
        ws[f'P{row}'].alignment = Alignment(horizontal='center')
        ws[f'P{row}'].number_format = '0.00'
        
        # Grade
        ws[f'Q{row}'] = f'=IF(\'Data Source\'!B{row}="","",\'Data Source\'!S{row})'
        ws[f'Q{row}'].alignment = Alignment(horizontal='center')
    
    # Column widths
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 20
    for col in range(3, 16):
        ws.column_dimensions[chr(64 + col)].width = 10
    ws.column_dimensions['P'].width = 12
    ws.column_dimensions['Q'].width = 10

def create_filtered_grade_sheet(wb, grade_filter, sheet_name):
    """Create sheet for specific grade (A+ or F students)"""
    ws = wb.create_sheet(sheet_name)
    
    # Title
    ws.merge_cells('A1:P1')
    title_cell = ws['A1']
    title_cell.value = f'STUDENTS WITH GRADE {grade_filter}'
    title_cell.font = Font(size=16, bold=True, color="FFFFFF")
    if grade_filter == 'A+':
        title_cell.fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
    else:
        title_cell.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30
    
    # Headers
    headers = ['SL', 'Name'] + list(SUBJECTS.keys()) + ['Overall GPA', 'Grade']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=header.replace('_', ' '))
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # Data rows - show subject grades
    display_row = 3
    for data_row in range(3, 1003):
        grade_check = f"'Data Source'!S{data_row}=\"{grade_filter}\""
        combined_check = f"AND('Data Source'!B{data_row}<>\"\",{grade_check})"
        
        # SL
        ws[f'A{display_row}'] = f'=IF({combined_check},COUNTIF(\'Data Source\'!$S$3:S{data_row},"{grade_filter}"),"")'
        ws[f'A{display_row}'].alignment = Alignment(horizontal='center')
        
        # Name
        ws[f'B{display_row}'] = f'=IF({combined_check},\'Data Source\'!B{data_row},"")'
        
        # Subject Grades (from columns U to AG)
        gp_cols = ['U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG']
        for col_idx, gp_col in enumerate(gp_cols, 3):
            grade_formula = (
                f'=IF({combined_check},IF(\'Data Source\'!{gp_col}{data_row}>=5,"A+",'
                f'IF(\'Data Source\'!{gp_col}{data_row}>=4,"A",IF(\'Data Source\'!{gp_col}{data_row}>=3.5,"A-",'
                f'IF(\'Data Source\'!{gp_col}{data_row}>=3,"B",IF(\'Data Source\'!{gp_col}{data_row}>=2,"C",'
                f'IF(\'Data Source\'!{gp_col}{data_row}>=1,"D","F")))))),"")'
            )
            ws.cell(row=display_row, column=col_idx).value = grade_formula
            ws.cell(row=display_row, column=col_idx).alignment = Alignment(horizontal='center')
        
        # Overall GPA
        ws[f'P{display_row}'] = f'=IF({combined_check},\'Data Source\'!R{data_row},"")'
        ws[f'P{display_row}'].alignment = Alignment(horizontal='center')
        ws[f'P{display_row}'].number_format = '0.00'
        
        # Grade
        ws[f'Q{display_row}'] = f'=IF({combined_check},\'Data Source\'!S{data_row},"")'
        ws[f'Q{display_row}'].alignment = Alignment(horizontal='center')
        
        display_row += 1
    
    # Column widths
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 20
    for col in range(3, 16):
        ws.column_dimensions[chr(64 + col)].width = 10
    ws.column_dimensions['P'].width = 12
    ws.column_dimensions['Q'].width = 10

def create_rank_position_sheet(wb):
    """Create Rank/Position sheet showing all passed students sorted by position"""
    ws = wb.create_sheet('Rank & Position')
    
    # Title
    ws.merge_cells('A1:F1')
    title_cell = ws['A1']
    title_cell.value = 'STUDENT RANK & POSITION (BY TOTAL MARKS)'
    title_cell.font = Font(size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30
    
    # Headers
    headers = ['Position', 'Name', 'Total Marks', 'Average', 'GPA', 'Grade']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Data rows - show students sorted by position (1, 2, 3, etc.)
    # For each display row, find the student with that position in Data Source
    for display_row in range(3, 1003):
        position = display_row - 2  # Position 1, 2, 3, etc.
        
        # Position
        ws[f'A{display_row}'] = f'=IFERROR(IF(COUNTIF(\'Data Source\'!$T:$T,{position})>0,{position},""),"")'
        ws[f'A{display_row}'].alignment = Alignment(horizontal='center')
        
        # Name - find student with this position
        ws[f'B{display_row}'] = f'=IFERROR(INDEX(\'Data Source\'!$B:$B,MATCH({position},\'Data Source\'!$T:$T,0)),"")'
        
        # Total Marks
        ws[f'C{display_row}'] = f'=IFERROR(INDEX(\'Data Source\'!$P:$P,MATCH({position},\'Data Source\'!$T:$T,0)),"")'
        ws[f'C{display_row}'].alignment = Alignment(horizontal='center')
        
        # Average
        ws[f'D{display_row}'] = f'=IFERROR(INDEX(\'Data Source\'!$Q:$Q,MATCH({position},\'Data Source\'!$T:$T,0)),"")'
        ws[f'D{display_row}'].alignment = Alignment(horizontal='center')
        ws[f'D{display_row}'].number_format = '0.00'
        
        # GPA
        ws[f'E{display_row}'] = f'=IFERROR(INDEX(\'Data Source\'!$R:$R,MATCH({position},\'Data Source\'!$T:$T,0)),"")'
        ws[f'E{display_row}'].alignment = Alignment(horizontal='center')
        ws[f'E{display_row}'].number_format = '0.00'
        
        # Grade
        ws[f'F{display_row}'] = f'=IFERROR(INDEX(\'Data Source\'!$S:$S,MATCH({position},\'Data Source\'!$T:$T,0)),"")'
        ws[f'F{display_row}'].alignment = Alignment(horizontal='center')
    
    # Column widths
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 8
    ws.column_dimensions['F'].width = 10

# ============================================================================
# MAIN GENERATION FUNCTION
# ============================================================================

def generate_excel_file():
    """Main function to generate the complete Excel dashboard"""
    
    print("🚀 Generating Class 8 Academic Results Dashboard...")
    
    # Generate sample data
    df = generate_sample_data(20)
    
    # Create workbook
    wb = Workbook()
    
    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    # Create all sheets
    print("📄 Creating Data Source sheet...")
    create_data_source_sheet(wb, df)
    
    print("📊 Creating Dashboard sheet...")
    create_dashboard_sheet(wb)
    
    print("📈 Creating Subject-wise GPA sheet...")
    create_subject_gpa_sheet(wb)
    
    print("📋 Creating Subject-wise Grade sheet...")
    create_subject_grade_sheet(wb)
    
    print("❌ Creating Fail Students sheet...")
    create_filtered_grade_sheet(wb, 'F', 'Fail Students')
    
    print("⭐ Creating A+ Students sheet...")
    create_filtered_grade_sheet(wb, 'A+', 'A+ Students')
    
    print("🏆 Creating Rank & Position sheet...")
    create_rank_position_sheet(wb)
    
    # Set calculation mode to automatic
    wb.calculation.calcMode = 'auto'
    
    # Save file
    filename = 'Class_8_Academic_Results.xlsx'
    wb.save(filename)
    
    print(f"\n✅ Excel file created successfully: {filename}")
    
    # Print summary statistics
    total_students = len(df)
    subject_cols = list(SUBJECTS.keys())
    
    # Calculate pass/fail for each student
    passed = 0
    for idx in range(len(df)):
        all_passed = True
        for subject in subject_cols:
            if df.iloc[idx][subject] < PASS_THRESHOLDS[subject]:
                all_passed = False
                break
        if all_passed:
            passed += 1
    
    print(f"\n📊 Summary:")
    print(f"   - Total Students: {total_students}")
    print(f"   - Pass Rate: {(passed/total_students*100):.1f}%")
    print(f"   - Total Marks: {TOTAL_MARKS}")
    print(f"\n📁 File saved as: {filename}")

if __name__ == "__main__":
    generate_excel_file()
