#!/usr/bin/env python3
"""
Academic Results Dashboard Generator
Creates an Excel file with Data Source, Pivot, and Dashboard sheets
"""

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
import warnings
warnings.filterwarnings('ignore')

def calculate_grade_point(marks, full_marks=100):
    """Calculate grade point for individual subject based on Dakhil curriculum
    
    Args:
        marks: Obtained marks
        full_marks: Total marks for the subject (100 or 200)
    """
    # Calculate percentage
    percentage = (marks / full_marks) * 100
    
    if percentage >= 80: return 5.0
    elif percentage >= 70: return 4.0
    elif percentage >= 60: return 3.5
    elif percentage >= 50: return 3.0
    elif percentage >= 40: return 2.0
    elif percentage >= 33: return 1.0
    else: return 0.0

def calculate_gpa_dakhil(row):
    """Calculate GPA based on Bangladeshi Dakhil curriculum
    
    Compulsory subjects:
    - 6 separate subjects (100 marks each)
    - 3 combined subjects (200 marks each)
    
    Additional subject (Mantiq):
    - Only counts if GPA >= 2.0
    - Adds bonus to final GPA
    
    Continuous Assessment (Career & Physical Education):
    - Pass/Fail only, doesn't affect GPA calculation
    """
    # Compulsory subjects for Dakhil curriculum (8 subjects, 4 combined = 12 columns)
    # Combined subjects are stored in two columns but calculated together
    compulsory_subjects = {
        'Quran_Hadith': ('Quran', 'Hadith', 200),  # (col1, col2, total_marks)
        'Arabic': ('Arabic_I', 'Arabic_II', 200),
        'Aqaid': ('Aqaid', None, 100),  # Single column subjects
        'English': ('English_I', 'English_II', 200),
        'Bangla': ('Bangla_I', 'Bangla_II', 200),
        'Mathematics': ('Mathematics', None, 100),
        'Islamic_History': ('Islamic_History', None, 100),
        'ICT': ('ICT', None, 100)
    }
    
    # Check if any compulsory subject failed (below 33%)
    for subject_name, subject_info in compulsory_subjects.items():
        col1, col2, full_marks = subject_info
        # If combined subject, add both columns
        if col2:
            marks = row[col1] + row[col2]
        else:
            marks = row[col1]
        min_passing = full_marks * 0.33
        if marks < min_passing:
            return 0.0
    
    # Check continuous assessment subjects (must pass)
    if row['Career_Education'] < 33 or row['Physical_Education'] < 33:
        return 0.0
    
    # Calculate grade points for all compulsory subjects
    grade_points = []
    for subject_name, subject_info in compulsory_subjects.items():
        col1, col2, full_marks = subject_info
        # If combined subject, add both columns
        if col2:
            marks = row[col1] + row[col2]
        else:
            marks = row[col1]
        gp = calculate_grade_point(marks, full_marks)
        grade_points.append(gp)
    
    # Calculate base GPA (average of compulsory subjects)
    base_gpa = sum(grade_points) / len(grade_points)
    
    # Handle additional subject (Mantiq) - Dakhil rule
    # If Mantiq GPA >= 2.0, add bonus points
    mantiq_gp = calculate_grade_point(row['Mantiq'], 100)
    if mantiq_gp >= 2.0:
        # Add (Mantiq GP - 2) / number of compulsory subjects
        bonus = (mantiq_gp - 2.0) / len(compulsory_subjects)
        base_gpa = base_gpa + bonus
    
    # Final GPA cannot exceed 5.0
    final_gpa = min(5.0, round(base_gpa, 2))
    return final_gpa

def calculate_letter_grade(gpa):
    """Convert GPA to letter grade"""
    if gpa >= 5.0: return 'A+'
    elif gpa >= 4.0: return 'A'
    elif gpa >= 3.5: return 'A-'
    elif gpa >= 3.0: return 'B'
    elif gpa >= 2.0: return 'C'
    elif gpa >= 1.0: return 'D'
    else: return 'F'

def create_data_source():
    """Create sample student data for Dakhil curriculum"""
    import random
    random.seed(42)  # For reproducible results
    
    data = {
        'SL': list(range(1, 21)),
        'Name': [
            'Ahmed Rahman', 'Fatima Khan', 'Karim Hassan', 'Ayesha Begum',
            'Rahim Uddin', 'Nadia Islam', 'Jahir Ahmed', 'Sadia Sultana',
            'Tarik Hasan', 'Ruhi Akter', 'Mehedi Hassan', 'Lubna Khatun',
            'Fahim Ahmed', 'Samira Begum', 'Imran Khan', 'Rafia Islam',
            'Shakib Rahman', 'Nusrat Jahan', 'Arif Hossain', 'Tasnia Akter'
        ],
        # Compulsory Subjects (8 subjects - 4 combined = 12 columns)
        'Quran': [random.randint(60, 95) for _ in range(20)],  # 100 marks
        'Hadith': [random.randint(65, 95) for _ in range(20)],  # 100 marks
        'Arabic_I': [random.randint(60, 95) for _ in range(20)],  # 100 marks
        'Arabic_II': [random.randint(65, 95) for _ in range(20)],  # 100 marks
        'Aqaid': [random.randint(60, 95) for _ in range(20)],  # 100 marks
        'English_I': [random.randint(55, 95) for _ in range(20)],  # 100 marks
        'English_II': [random.randint(60, 90) for _ in range(20)],  # 100 marks
        'Bangla_I': [random.randint(60, 95) for _ in range(20)],  # 100 marks
        'Bangla_II': [random.randint(65, 95) for _ in range(20)],  # 100 marks
        'Mathematics': [random.randint(55, 95) for _ in range(20)],  # 100 marks
        'Islamic_History': [random.randint(60, 90) for _ in range(20)],  # 100 marks
        'ICT': [random.randint(65, 95) for _ in range(20)],  # 100 marks
        # Additional Subject
        'Mantiq': [random.randint(50, 90) for _ in range(20)],  # 100 marks
        # Continuous Assessment Subjects
        'Career_Education': [random.randint(70, 95) for _ in range(20)],  # 100 marks (Pass/Fail)
        'Physical_Education': [random.randint(75, 95) for _ in range(20)],  # 100 marks (Pass/Fail)
    }
    
    # Return DataFrame with only raw marks. Total/Average/GPA/Grade will be
    # inserted into the workbook as Excel formulas so the workbook stays
    # dynamic when users edit marks directly in Excel.
    df = pd.DataFrame(data)
    subjects = ['Quran', 'Hadith', 'Arabic_I', 'Arabic_II', 'Aqaid', 'English_I', 'English_II', 
                'Bangla_I', 'Bangla_II', 'Mathematics', 'Islamic_History', 'ICT', 'Mantiq', 'Career_Education', 'Physical_Education']
    return df[(['SL', 'Name'] + subjects)]

def style_data_source_sheet(ws, df):
    """Apply styling to Data Source sheet"""
    
    # Header styling
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Apply conditional formatting colors for marks
    # All subjects are now 100 marks each (15 columns: C-Q)
    # C=Quran, D=Hadith, E=Arabic I, F=Arabic II, G=Aqaid, H=English I, I=English II, 
    # J=Bangla I, K=Bangla II, L=Math, M=Islamic History, N=ICT, O=Mantiq, P=Career, Q=Physical
    subject_cols_100_marks = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']

    for row in range(2, len(df) + 2):
        # All subjects are 100 marks
        for col in subject_cols_100_marks:
            cell = ws[f'{col}{row}']
            value = cell.value
            if value is None:
                continue
            try:
                num = float(value)
            except Exception:
                continue
            # Thresholds for 100 marks
            if num >= 80:
                cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
            elif num >= 60:
                cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            elif num >= 40:
                cell.fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
            else:
                cell.fill = PatternFill(start_color="FC9999", end_color="FC9999", fill_type="solid")
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 5   # SL
    ws.column_dimensions['B'].width = 18  # Name
    ws.column_dimensions['C'].width = 10  # Quran
    ws.column_dimensions['D'].width = 10  # Hadith
    ws.column_dimensions['E'].width = 10  # Arabic I
    ws.column_dimensions['F'].width = 10  # Arabic II
    ws.column_dimensions['G'].width = 10  # Aqaid
    ws.column_dimensions['H'].width = 10  # English I
    ws.column_dimensions['I'].width = 10  # English II
    ws.column_dimensions['J'].width = 10  # Bangla I
    ws.column_dimensions['K'].width = 10  # Bangla II
    ws.column_dimensions['L'].width = 10  # Mathematics
    ws.column_dimensions['M'].width = 13  # Islamic History
    ws.column_dimensions['N'].width = 10  # ICT
    ws.column_dimensions['O'].width = 10  # Mantiq
    ws.column_dimensions['P'].width = 12  # Career Education
    ws.column_dimensions['Q'].width = 12  # Physical Education
    ws.column_dimensions['R'].width = 10  # Total
    ws.column_dimensions['S'].width = 10  # Average
    ws.column_dimensions['T'].width = 8   # GPA
    ws.column_dimensions['U'].width = 12  # Overall Grade

def create_subject_grades_sheet(wb, df):
    """Create Subject Grades sheet with individual subject grades"""
    from openpyxl.utils.dataframe import dataframe_to_rows
    
    ws = wb.create_sheet('Subject Grades')
    
    # Title
    ws.merge_cells('A1:O1')
    title_cell = ws['A1']
    title_cell.value = 'SUBJECT-WISE GRADES'
    title_cell.font = Font(size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30
    
    # Headers
    headers = ['SL', 'Name', 'Quran+Hadith', 'Arabic', 'Aqaid', 'English', 'Bangla', 
               'Math', 'History', 'ICT', 'Mantiq', 'Career', 'Physical', 'Overall GPA', 'Overall Grade']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Add data and formulas
    for row_idx in range(len(df)):
        row = row_idx + 3  # Start from row 3 (after header)
        data_row = row_idx + 2  # Corresponding row in Data Source
        
        # SL and Name from Data Source
        ws[f'A{row}'] = f"='Data Source'!A{data_row}"
        ws[f'B{row}'] = f"='Data Source'!B{data_row}"
        
        # Subject grades with formulas (using new column layout)
        # Quran+Hadith (200 marks combined - C+D)
        ws[f'C{row}'] = f"=IF(('Data Source'!C{data_row}+'Data Source'!D{data_row})/200*100>=80,\"A+\",IF(('Data Source'!C{data_row}+'Data Source'!D{data_row})/200*100>=70,\"A\",IF(('Data Source'!C{data_row}+'Data Source'!D{data_row})/200*100>=60,\"A-\",IF(('Data Source'!C{data_row}+'Data Source'!D{data_row})/200*100>=50,\"B\",IF(('Data Source'!C{data_row}+'Data Source'!D{data_row})/200*100>=40,\"C\",IF(('Data Source'!C{data_row}+'Data Source'!D{data_row})/200*100>=33,\"D\",\"F\"))))))"
        
        # Arabic (200 marks combined - E+F)
        ws[f'D{row}'] = f"=IF(('Data Source'!E{data_row}+'Data Source'!F{data_row})/200*100>=80,\"A+\",IF(('Data Source'!E{data_row}+'Data Source'!F{data_row})/200*100>=70,\"A\",IF(('Data Source'!E{data_row}+'Data Source'!F{data_row})/200*100>=60,\"A-\",IF(('Data Source'!E{data_row}+'Data Source'!F{data_row})/200*100>=50,\"B\",IF(('Data Source'!E{data_row}+'Data Source'!F{data_row})/200*100>=40,\"C\",IF(('Data Source'!E{data_row}+'Data Source'!F{data_row})/200*100>=33,\"D\",\"F\"))))))"
        
        # Aqaid (100 marks - G)
        ws[f'E{row}'] = f"=IF('Data Source'!G{data_row}>=80,\"A+\",IF('Data Source'!G{data_row}>=70,\"A\",IF('Data Source'!G{data_row}>=60,\"A-\",IF('Data Source'!G{data_row}>=50,\"B\",IF('Data Source'!G{data_row}>=40,\"C\",IF('Data Source'!G{data_row}>=33,\"D\",\"F\"))))))"
        
        # English (200 marks combined - H+I)
        ws[f'F{row}'] = f"=IF(('Data Source'!H{data_row}+'Data Source'!I{data_row})/200*100>=80,\"A+\",IF(('Data Source'!H{data_row}+'Data Source'!I{data_row})/200*100>=70,\"A\",IF(('Data Source'!H{data_row}+'Data Source'!I{data_row})/200*100>=60,\"A-\",IF(('Data Source'!H{data_row}+'Data Source'!I{data_row})/200*100>=50,\"B\",IF(('Data Source'!H{data_row}+'Data Source'!I{data_row})/200*100>=40,\"C\",IF(('Data Source'!H{data_row}+'Data Source'!I{data_row})/200*100>=33,\"D\",\"F\"))))))"
        
        # Bangla (200 marks combined - J+K)
        ws[f'G{row}'] = f"=IF(('Data Source'!J{data_row}+'Data Source'!K{data_row})/200*100>=80,\"A+\",IF(('Data Source'!J{data_row}+'Data Source'!K{data_row})/200*100>=70,\"A\",IF(('Data Source'!J{data_row}+'Data Source'!K{data_row})/200*100>=60,\"A-\",IF(('Data Source'!J{data_row}+'Data Source'!K{data_row})/200*100>=50,\"B\",IF(('Data Source'!J{data_row}+'Data Source'!K{data_row})/200*100>=40,\"C\",IF(('Data Source'!J{data_row}+'Data Source'!K{data_row})/200*100>=33,\"D\",\"F\"))))))"
        
        # Mathematics (100 marks - L)
        ws[f'H{row}'] = f"=IF('Data Source'!L{data_row}>=80,\"A+\",IF('Data Source'!L{data_row}>=70,\"A\",IF('Data Source'!L{data_row}>=60,\"A-\",IF('Data Source'!L{data_row}>=50,\"B\",IF('Data Source'!L{data_row}>=40,\"C\",IF('Data Source'!L{data_row}>=33,\"D\",\"F\"))))))"
        
        # Islamic History (100 marks - M)
        ws[f'I{row}'] = f"=IF('Data Source'!M{data_row}>=80,\"A+\",IF('Data Source'!M{data_row}>=70,\"A\",IF('Data Source'!M{data_row}>=60,\"A-\",IF('Data Source'!M{data_row}>=50,\"B\",IF('Data Source'!M{data_row}>=40,\"C\",IF('Data Source'!M{data_row}>=33,\"D\",\"F\"))))))"
        
        # ICT (100 marks - N)
        ws[f'J{row}'] = f"=IF('Data Source'!N{data_row}>=80,\"A+\",IF('Data Source'!N{data_row}>=70,\"A\",IF('Data Source'!N{data_row}>=60,\"A-\",IF('Data Source'!N{data_row}>=50,\"B\",IF('Data Source'!N{data_row}>=40,\"C\",IF('Data Source'!N{data_row}>=33,\"D\",\"F\"))))))"
        
        # Mantiq (100 marks - O)
        ws[f'K{row}'] = f"=IF('Data Source'!O{data_row}>=80,\"A+\",IF('Data Source'!O{data_row}>=70,\"A\",IF('Data Source'!O{data_row}>=60,\"A-\",IF('Data Source'!O{data_row}>=50,\"B\",IF('Data Source'!O{data_row}>=40,\"C\",IF('Data Source'!O{data_row}>=33,\"D\",\"F\"))))))"
        
        # Career Education (Pass/Fail - P)
        ws[f'L{row}'] = f"=IF('Data Source'!P{data_row}>=33,\"Pass\",\"Fail\")"
        
        # Physical Education (Pass/Fail - Q)
        ws[f'M{row}'] = f"=IF('Data Source'!Q{data_row}>=33,\"Pass\",\"Fail\")"
        
        # Overall GPA from Data Source (column T)
        ws[f'N{row}'] = f"='Data Source'!T{data_row}"
        ws[f'N{row}'].number_format = '0.00'
        
        # Overall Grade from Data Source (column U)
        ws[f'O{row}'] = f"='Data Source'!U{data_row}"
        
        # Alignment
        for col in ['A', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']:
            ws[f'{col}{row}'].alignment = Alignment(horizontal='center')
    
    # Column widths
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 18
    for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']:
        ws.column_dimensions[col].width = 10
    ws.column_dimensions['N'].width = 10
    ws.column_dimensions['O'].width = 12

def create_dashboard_sheet(wb, df):
    """Create Dashboard sheet with charts and summary"""
    
    ws = wb.create_sheet('Dashboard')
    
    # Title
    ws.merge_cells('A1:L2')
    title_cell = ws['A1']
    title_cell.value = '📊 ACADEMIC RESULTS DASHBOARD'
    title_cell.font = Font(size=24, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Summary Statistics (use Excel formulas so dashboard updates when Data Source changes)
    ws['A4'] = 'SUMMARY STATISTICS'
    ws['A4'].font = Font(size=14, bold=True)

    # Total Students (count names), Average GPA, Highest Total, Pass Rate
    ws['B5'] = 'Total Students:'
    ws['C5'] = "=COUNTA('Data Source'!B:B)-1"
    ws['E5'] = 'Average GPA:'
    ws['F5'] = "=IFERROR(ROUND(AVERAGE('Data Source'!T:T),2), 0)"  # Column T is GPA
    ws['H5'] = 'Highest Total:'
    ws['I5'] = "=MAX('Data Source'!R:R)"  # Column R is Total
    ws['K5'] = 'Pass Rate:'
    # Count GPA > 0 divided by total - format as percentage
    ws['L5'] = '=IF(C5>0,COUNTIF(\'Data Source\'!T:T,">0")/C5,0)'  # Column T is GPA
    ws['L5'].number_format = '0.0%'

    # Style the label cells
    for label_cell in ['B5','E5','H5','K5']:
        ws[label_cell].font = Font(bold=True)
    for value_cell in ['C5','F5','I5','L5']:
        ws[value_cell].font = Font(size=12, bold=True, color="1F4E78")
    
    # Format number cells
    ws['F5'].number_format = '0.00'  # GPA with 2 decimals
    
    # Grade Distribution Table
    ws['A7'] = 'GRADE DISTRIBUTION'
    ws['A7'].font = Font(size=12, bold=True)
    
    # Grade distribution using COUNTIF for common grade buckets
    ws['A8'] = 'Grade'
    ws['B8'] = 'Count'
    ws['A8'].font = Font(bold=True)
    ws['B8'].font = Font(bold=True)
    ws['A8'].alignment = Alignment(horizontal='center')
    ws['B8'].alignment = Alignment(horizontal='center')
    
    grade_list = ['A+','A','A-','B','C','D','F']
    row = 9
    for grade in grade_list:
        ws[f'A{row}'] = grade
        ws[f'A{row}'].alignment = Alignment(horizontal='center')
        # COUNTIF on Data Source Overall Grade column Q
        ws[f'B{row}'] = f"=COUNTIF('Data Source'!Q:Q,\"{grade}\")"
        ws[f'B{row}'].alignment = Alignment(horizontal='center')
        row += 1
    
    # Subject-wise Average
    ws['D7'] = 'SUBJECT-WISE AVERAGE'
    ws['D7'].font = Font(size=12, bold=True)
    
    # Dakhil curriculum subjects with their columns and full marks
    subjects_info = [
        ('Quran Mazid', 'C', 100),
        ('Arabic (Comb.)', 'D', 200),
        ('Aqaid', 'E', 100),
        ('English (Comb.)', 'F', 200),
        ('Bangla (Comb.)', 'G', 200),
        ('Mathematics', 'H', 100),
        ('Islamic History', 'I', 100),
        ('ICT', 'J', 100),
        ('Mantiq', 'K', 100)
    ]
    
    ws['D8'] = 'Subject'
    ws['E8'] = 'Avg'
    ws['F8'] = '%'
    ws['D8'].font = Font(bold=True)
    ws['E8'].font = Font(bold=True)
    ws['F8'].font = Font(bold=True)
    ws['E8'].alignment = Alignment(horizontal='center')
    ws['F8'].alignment = Alignment(horizontal='center')
    
    row = 9
    for subject_name, col_letter, full_marks in subjects_info:
        ws[f'D{row}'] = subject_name
        ws[f'E{row}'] = f"=IFERROR(ROUND(AVERAGE('Data Source'!{col_letter}:{col_letter}),2),0)"
        ws[f'E{row}'].number_format = '0.00'
        ws[f'E{row}'].alignment = Alignment(horizontal='center')
        # Percentage average
        ws[f'F{row}'] = f"=IFERROR(ROUND(E{row}/{full_marks}*100,1),0)"
        ws[f'F{row}'].number_format = '0.0"%"'
        ws[f'F{row}'].alignment = Alignment(horizontal='center')
        row += 1
    
    # Top 5 Students - Simplified approach using Python to get initial top 5
    # Then use formulas that reference Data Source directly
    ws['G7'] = 'TOP 5 STUDENTS'
    ws['G7'].font = Font(size=12, bold=True)
    
    ws['G8'] = 'Rank'
    ws['H8'] = 'Name'
    ws['I8'] = 'GPA'
    ws['G8'].font = Font(bold=True)
    ws['H8'].font = Font(bold=True)
    ws['I8'].font = Font(bold=True)
    ws['G8'].alignment = Alignment(horizontal='center')
    ws['I8'].alignment = Alignment(horizontal='center')
    
    # Calculate top 5 students in Python to get their row numbers using Dakhil GPA calculation
    df_with_gpa = df.copy()
    df_with_gpa['GPA_calc'] = df_with_gpa.apply(calculate_gpa_dakhil, axis=1)
    df_sorted = df_with_gpa.nlargest(5, 'GPA_calc')
    top_5_rows = [df_with_gpa.index.get_loc(idx) + 2 for idx in df_sorted.index]  # +2 for header and 1-based indexing
    
    # Create Top 5 list with formulas that reference specific rows from Data Source
    row = 9
    for i, data_source_row in enumerate(top_5_rows, 1):
        ws[f'G{row}'] = i
        ws[f'G{row}'].alignment = Alignment(horizontal='center')
        
        # Reference the specific student's name and GPA from Data Source (column T)
        # This way if their marks change, their GPA updates
        ws[f'H{row}'] = f"='Data Source'!B{data_source_row}"
        ws[f'I{row}'] = f"='Data Source'!T{data_source_row}"
        ws[f'I{row}'].number_format = '0.00'
        ws[f'I{row}'].alignment = Alignment(horizontal='center')
        
        row += 1
    
    # Create Pie Chart - Grade Distribution
    pie = PieChart()
    labels = Reference(ws, min_col=1, min_row=9, max_row=15)
    data = Reference(ws, min_col=2, min_row=8, max_row=15)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.title = "Grade Distribution"
    ws.add_chart(pie, "A17")
    
    # Create Bar Chart - Subject-wise Performance (9 Dakhil subjects)
    bar = BarChart()
    bar.type = "col"
    bar.title = "Subject-wise Average Scores"
    bar.y_axis.title = 'Average Marks'
    bar.x_axis.title = 'Subjects'
    
    data = Reference(ws, min_col=5, min_row=8, max_row=17)  # E8:E17 (9 subjects + header)
    cats = Reference(ws, min_col=4, min_row=9, max_row=17)  # D9:D17 (9 subject names)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    ws.add_chart(bar, "G17")
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['H'].width = 18

def generate_excel_file(filename='Academic_Results_Dashboard.xlsx'):
    """Main function to generate the Excel file"""
    
    print("🚀 Generating Academic Results Dashboard...")
    
    # Create data
    df = create_data_source()
    
    # Create Excel file
    wb = Workbook()
    ws = wb.active
    ws.title = 'Data Source'
    
    # Write data to Data Source sheet
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    
    # Add Total, Average, GPA, Overall Grade columns (columns R-U)
    # Subject columns: C-Q (15 subjects - 4 combined subjects split into pairs)
    # Summary columns: R=Total, S=Average, T=GPA, U=Overall Grade
    
    # Summary columns
    ws['R1'] = 'Total'
    ws['S1'] = 'Average'
    ws['T1'] = 'GPA'
    ws['U1'] = 'Overall Grade'
    
    # Add formulas for each student row (rows 2 to len(df)+1)
    for row in range(2, len(df) + 2):
        # New column structure:
        # C=Quran(100), D=Hadith(100), E=Arabic I(100), F=Arabic II(100), G=Aqaid(100), 
        # H=English I(100), I=English II(100), J=Bangla I(100), K=Bangla II(100), 
        # L=Math(100), M=Islamic History(100), N=ICT(100), O=Mantiq(100), P=Career(100), Q=Physical(100)
        
        # Total: Sum of compulsory subjects (8 subjects but 12 columns due to 4 combined subjects)
        # Compulsory: Quran+Hadith, Arabic I+II, Aqaid, English I+II, Bangla I+II, Math, Islamic History, ICT
        # That's columns C through N (excluding O=Mantiq, P=Career, Q=Physical)
        ws[f'R{row}'] = f'=SUM(C{row}:N{row})'
        
        # Average: Average of compulsory subjects (total marks = 1200, so divide by 12 columns then multiply by 100/100)
        # Actually simpler: Total/12 to get average per 100-mark subject
        ws[f'S{row}'] = f'=ROUND(R{row}/12,2)'
        ws[f'S{row}'].number_format = '0.00'
        
        # GPA: Dakhil curriculum calculation
        # For combined subjects, we need to add the two columns and treat as 200 marks
        # GP formula for 100-mark subject: IF(marks>=80,5,IF(marks>=70,4,IF(marks>=60,3.5,IF(marks>=50,3,IF(marks>=40,2,IF(marks>=33,1,0))))))
        # GP formula for 200-mark subject: IF(marks/200*100>=80,5,IF(marks/200*100>=70,4,...
        
        # Combined subjects (add two columns, treat as 200 marks)
        gp_quran_hadith = f'IF((C{row}+D{row})/200*100>=80,5,IF((C{row}+D{row})/200*100>=70,4,IF((C{row}+D{row})/200*100>=60,3.5,IF((C{row}+D{row})/200*100>=50,3,IF((C{row}+D{row})/200*100>=40,2,IF((C{row}+D{row})/200*100>=33,1,0))))))'
        gp_arabic = f'IF((E{row}+F{row})/200*100>=80,5,IF((E{row}+F{row})/200*100>=70,4,IF((E{row}+F{row})/200*100>=60,3.5,IF((E{row}+F{row})/200*100>=50,3,IF((E{row}+F{row})/200*100>=40,2,IF((E{row}+F{row})/200*100>=33,1,0))))))'
        gp_aqaid = f'IF(G{row}>=80,5,IF(G{row}>=70,4,IF(G{row}>=60,3.5,IF(G{row}>=50,3,IF(G{row}>=40,2,IF(G{row}>=33,1,0))))))'
        gp_english = f'IF((H{row}+I{row})/200*100>=80,5,IF((H{row}+I{row})/200*100>=70,4,IF((H{row}+I{row})/200*100>=60,3.5,IF((H{row}+I{row})/200*100>=50,3,IF((H{row}+I{row})/200*100>=40,2,IF((H{row}+I{row})/200*100>=33,1,0))))))'
        gp_bangla = f'IF((J{row}+K{row})/200*100>=80,5,IF((J{row}+K{row})/200*100>=70,4,IF((J{row}+K{row})/200*100>=60,3.5,IF((J{row}+K{row})/200*100>=50,3,IF((J{row}+K{row})/200*100>=40,2,IF((J{row}+K{row})/200*100>=33,1,0))))))'
        gp_math = f'IF(L{row}>=80,5,IF(L{row}>=70,4,IF(L{row}>=60,3.5,IF(L{row}>=50,3,IF(L{row}>=40,2,IF(L{row}>=33,1,0))))))'
        gp_history = f'IF(M{row}>=80,5,IF(M{row}>=70,4,IF(M{row}>=60,3.5,IF(M{row}>=50,3,IF(M{row}>=40,2,IF(M{row}>=33,1,0))))))'
        gp_ict = f'IF(N{row}>=80,5,IF(N{row}>=70,4,IF(N{row}>=60,3.5,IF(N{row}>=50,3,IF(N{row}>=40,2,IF(N{row}>=33,1,0))))))'
        gp_mantiq = f'IF(O{row}>=80,5,IF(O{row}>=70,4,IF(O{row}>=60,3.5,IF(O{row}>=50,3,IF(O{row}>=40,2,IF(O{row}>=33,1,0))))))'
        
        # Check if any compulsory subject or continuous assessment failed
        # Combined subjects need at least 66 marks total (33% of 200)
        # Single 100-mark subjects need at least 33 marks (33% of 100)
        fail_check = f'OR((C{row}+D{row})<66,(E{row}+F{row})<66,G{row}<33,(H{row}+I{row})<66,(J{row}+K{row})<66,L{row}<33,M{row}<33,N{row}<33,P{row}<33,Q{row}<33)'
        
        # Base GPA = average of 8 compulsory subjects
        base_gpa = f'({gp_quran_hadith}+{gp_arabic}+{gp_aqaid}+{gp_english}+{gp_bangla}+{gp_math}+{gp_history}+{gp_ict})/8'
        
        # Additional subject bonus: If Mantiq GP >= 2, add (Mantiq GP - 2) / 8
        mantiq_bonus = f'IF({gp_mantiq}>=2,({gp_mantiq}-2)/8,0)'
        
        # Final GPA formula
        gpa_formula = f'=IF({fail_check},0,MIN(5,ROUND({base_gpa}+{mantiq_bonus},2)))'
        
        ws[f'T{row}'] = gpa_formula
        ws[f'T{row}'].number_format = '0.00'
        
        # Overall Grade: Based on GPA value
        grade_formula = f'=IF(T{row}>=5,"A+",IF(T{row}>=4,"A",IF(T{row}>=3.5,"A-",IF(T{row}>=3,"B",IF(T{row}>=2,"C",IF(T{row}>=1,"D","F"))))))'
        ws[f'U{row}'] = grade_formula
        
        # Center align GPA and Overall Grade
        ws[f'T{row}'].alignment = Alignment(horizontal='center')
        ws[f'U{row}'].alignment = Alignment(horizontal='center')
    
    # Style Data Source sheet
    style_data_source_sheet(ws, df)
    
    # Create Dashboard sheet
    create_dashboard_sheet(wb, df)
    
    # Create Subject Grades sheet
    create_subject_grades_sheet(wb, df)
    
    # Note: Pivot sheet would require manual creation in Excel or additional library
    wb.create_sheet('Pivot')
    pivot_ws = wb['Pivot']
    pivot_ws['A1'] = 'Pivot tables can be created manually in Excel'
    pivot_ws['A2'] = 'Use Insert > PivotTable from the Data Source sheet'
    
    # Save file
    wb.save(filename)
    print(f"✅ Excel file created successfully: {filename}")
    # For console summary we can compute GPA locally (this does not alter the
    # workbook which contains formulas). This provides quick feedback to user.
    df_console = df.copy()
    try:
        df_console['GPA'] = df_console.apply(calculate_gpa_dakhil, axis=1)
        # Grade based on GPA
        def get_grade(gpa):
            if gpa >= 5: return 'A+'
            elif gpa >= 4: return 'A'
            elif gpa >= 3.5: return 'A-'
            elif gpa >= 3: return 'B'
            elif gpa >= 2: return 'C'
            elif gpa >= 1: return 'D'
            else: return 'F'
        df_console['Grade'] = df_console['GPA'].apply(get_grade)
        avg_gpa = df_console['GPA'].mean()
        num_a_plus = (df_console['Grade'] == 'A+').sum()
        pass_rate = (df_console['GPA'] > 0).sum() / len(df_console) * 100
    except Exception:
        avg_gpa = 0
        num_a_plus = 0
        pass_rate = 0

    print(f"\n📊 Summary:")
    print(f"   - Total Students: {len(df)}")
    print(f"   - Average Class GPA: {avg_gpa:.2f}")
    print(f"   - Students with A+: {num_a_plus}")
    print(f"   - Pass Rate: {pass_rate:.1f}%")
    print(f"\n📁 File saved as: {filename}")

if __name__ == "__main__":
    generate_excel_file()