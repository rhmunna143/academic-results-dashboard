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
    # Compulsory subjects with their full marks
    compulsory_subjects = {
        'Quran_Mazid': 200,
        'Arabic_Combined': 200,
        'Aqaid': 100,
        'English_Combined': 200,
        'Bangla_Combined': 200,
        'Mathematics': 100,
        'Islamic_History': 100,
        'ICT': 100
    }
    
    # Check if any compulsory subject failed (below 33%)
    for subject, full_marks in compulsory_subjects.items():
        min_passing = full_marks * 0.33
        if row[subject] < min_passing:
            return 0.0
    
    # Check continuous assessment subjects (must pass)
    if row['Career_Education'] < 33 or row['Physical_Education'] < 33:
        return 0.0
    
    # Calculate grade points for all compulsory subjects
    grade_points = []
    for subject, full_marks in compulsory_subjects.items():
        gp = calculate_grade_point(row[subject], full_marks)
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
        # Compulsory Subjects (8 subjects)
        'Quran_Mazid': [random.randint(130, 190) for _ in range(20)],  # 200 marks (Combined)
        'Arabic_Combined': [random.randint(130, 190) for _ in range(20)],  # 200 marks
        'Aqaid': [random.randint(60, 95) for _ in range(20)],  # 100 marks
        'English_Combined': [random.randint(120, 185) for _ in range(20)],  # 200 marks
        'Bangla_Combined': [random.randint(125, 190) for _ in range(20)],  # 200 marks
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
    subjects = ['Quran_Mazid', 'Arabic_Combined', 'Aqaid', 'English_Combined', 'Bangla_Combined', 
                'Mathematics', 'Islamic_History', 'ICT', 'Mantiq', 'Career_Education', 'Physical_Education']
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
    # Columns: C-M for all subjects
    # Different thresholds for 100-mark vs 200-mark subjects
    subject_cols_100 = ['C', 'E', 'F', 'H', 'I', 'J', 'K', 'L', 'M']  # Quran, Aqaid, Math, Islamic History, ICT, Mantiq, Career, Physical
    subject_cols_200 = ['D', 'G']  # Arabic Combined, English Combined, Bangla Combined (actually G is Bangla)
    
    # Update: Correct column mapping
    # C=Quran(200), D=Arabic(200), E=Aqaid(100), F=English(200), G=Bangla(200), H=Math(100), I=Islamic History(100), J=ICT(100), K=Mantiq(100), L=Career(100), M=Physical(100)
    subject_cols_100_marks = ['E', 'H', 'I', 'J', 'K', 'L', 'M']  
    subject_cols_200_marks = ['C', 'D', 'F', 'G']

    for row in range(2, len(df) + 2):
        # 100-mark subjects
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
        
        # 200-mark subjects
        for col in subject_cols_200_marks:
            cell = ws[f'{col}{row}']
            value = cell.value
            if value is None:
                continue
            try:
                num = float(value)
            except Exception:
                continue
            # Thresholds for 200 marks (double the 100-mark thresholds)
            if num >= 160:
                cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
            elif num >= 120:
                cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            elif num >= 80:
                cell.fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
            else:
                cell.fill = PatternFill(start_color="FC9999", end_color="FC9999", fill_type="solid")
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 5   # SL
    ws.column_dimensions['B'].width = 18  # Name
    ws.column_dimensions['C'].width = 12  # Quran Mazid
    ws.column_dimensions['D'].width = 14  # Arabic Combined
    ws.column_dimensions['E'].width = 10  # Aqaid
    ws.column_dimensions['F'].width = 14  # English Combined
    ws.column_dimensions['G'].width = 14  # Bangla Combined
    ws.column_dimensions['H'].width = 12  # Mathematics
    ws.column_dimensions['I'].width = 13  # Islamic History
    ws.column_dimensions['J'].width = 10  # ICT
    ws.column_dimensions['K'].width = 10  # Mantiq
    ws.column_dimensions['L'].width = 12  # Career Education
    ws.column_dimensions['M'].width = 12  # Physical Education
    ws.column_dimensions['N'].width = 11  # Quran Grade
    ws.column_dimensions['O'].width = 11  # Arabic Grade
    ws.column_dimensions['P'].width = 11  # Aqaid Grade
    ws.column_dimensions['Q'].width = 11  # English Grade
    ws.column_dimensions['R'].width = 11  # Bangla Grade
    ws.column_dimensions['S'].width = 11  # Math Grade
    ws.column_dimensions['T'].width = 11  # History Grade
    ws.column_dimensions['U'].width = 10  # ICT Grade
    ws.column_dimensions['V'].width = 11  # Mantiq Grade
    ws.column_dimensions['W'].width = 11  # Career Grade
    ws.column_dimensions['X'].width = 12  # Physical Grade
    ws.column_dimensions['Y'].width = 10  # Total
    ws.column_dimensions['Z'].width = 10  # Average
    ws.column_dimensions['AA'].width = 8  # GPA
    ws.column_dimensions['AB'].width = 8  # Overall Grade

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
    ws['F5'] = "=IFERROR(ROUND(AVERAGE('Data Source'!AA:AA),2), 0)"  # Column AA is GPA
    ws['H5'] = 'Highest Total:'
    ws['I5'] = "=MAX('Data Source'!Y:Y)"  # Column Y is Total
    ws['K5'] = 'Pass Rate:'
    # Count GPA > 0 divided by total - format as percentage
    ws['L5'] = '=IF(C5>0,COUNTIF(\'Data Source\'!AA:AA,">0")/C5,0)'  # Column AA is GPA
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
        # COUNTIF on Data Source Overall Grade column AB
        ws[f'B{row}'] = f"=COUNTIF('Data Source'!AB:AB,\"{grade}\")"
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
        
        # Reference the specific student's name and GPA from Data Source (column AA)
        # This way if their marks change, their GPA updates
        ws[f'H{row}'] = f"='Data Source'!B{data_source_row}"
        ws[f'I{row}'] = f"='Data Source'!AA{data_source_row}"
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
    
    # Add subject grade columns after each subject (columns N-X for 11 subject grades)
    # Then add Total, Average, GPA, Overall Grade columns
    # Subject columns: C-M (11 subjects)
    # Grade columns: N-X (11 grade columns, one after each subject)
    # Summary columns: Y=Total, Z=Average, AA=GPA, AB=Overall Grade
    
    # Subject grade headers
    ws['N1'] = 'Quran Grade'
    ws['O1'] = 'Arabic Grade'
    ws['P1'] = 'Aqaid Grade'
    ws['Q1'] = 'English Grade'
    ws['R1'] = 'Bangla Grade'
    ws['S1'] = 'Math Grade'
    ws['T1'] = 'History Grade'
    ws['U1'] = 'ICT Grade'
    ws['V1'] = 'Mantiq Grade'
    ws['W1'] = 'Career Grade'
    ws['X1'] = 'Physical Grade'
    
    # Summary columns
    ws['Y1'] = 'Total'
    ws['Z1'] = 'Average'
    ws['AA1'] = 'GPA'
    ws['AB1'] = 'Overall Grade'
    
    # Add formulas for each student row (rows 2 to len(df)+1)
    for row in range(2, len(df) + 2):
        # Grade formulas for each subject
        # For 100-mark subjects: grade based on actual marks
        # For 200-mark subjects: grade based on percentage
        
        # Quran Mazid (200 marks) - Column C, Grade in Column N
        ws[f'N{row}'] = f'=IF(C{row}/200*100>=80,"A+",IF(C{row}/200*100>=70,"A",IF(C{row}/200*100>=60,"A-",IF(C{row}/200*100>=50,"B",IF(C{row}/200*100>=40,"C",IF(C{row}/200*100>=33,"D","F"))))))'
        ws[f'N{row}'].alignment = Alignment(horizontal='center')
        
        # Arabic Combined (200 marks) - Column D, Grade in Column O
        ws[f'O{row}'] = f'=IF(D{row}/200*100>=80,"A+",IF(D{row}/200*100>=70,"A",IF(D{row}/200*100>=60,"A-",IF(D{row}/200*100>=50,"B",IF(D{row}/200*100>=40,"C",IF(D{row}/200*100>=33,"D","F"))))))'
        ws[f'O{row}'].alignment = Alignment(horizontal='center')
        
        # Aqaid (100 marks) - Column E, Grade in Column P
        ws[f'P{row}'] = f'=IF(E{row}>=80,"A+",IF(E{row}>=70,"A",IF(E{row}>=60,"A-",IF(E{row}>=50,"B",IF(E{row}>=40,"C",IF(E{row}>=33,"D","F"))))))'
        ws[f'P{row}'].alignment = Alignment(horizontal='center')
        
        # English Combined (200 marks) - Column F, Grade in Column Q
        ws[f'Q{row}'] = f'=IF(F{row}/200*100>=80,"A+",IF(F{row}/200*100>=70,"A",IF(F{row}/200*100>=60,"A-",IF(F{row}/200*100>=50,"B",IF(F{row}/200*100>=40,"C",IF(F{row}/200*100>=33,"D","F"))))))'
        ws[f'Q{row}'].alignment = Alignment(horizontal='center')
        
        # Bangla Combined (200 marks) - Column G, Grade in Column R
        ws[f'R{row}'] = f'=IF(G{row}/200*100>=80,"A+",IF(G{row}/200*100>=70,"A",IF(G{row}/200*100>=60,"A-",IF(G{row}/200*100>=50,"B",IF(G{row}/200*100>=40,"C",IF(G{row}/200*100>=33,"D","F"))))))'
        ws[f'R{row}'].alignment = Alignment(horizontal='center')
        
        # Mathematics (100 marks) - Column H, Grade in Column S
        ws[f'S{row}'] = f'=IF(H{row}>=80,"A+",IF(H{row}>=70,"A",IF(H{row}>=60,"A-",IF(H{row}>=50,"B",IF(H{row}>=40,"C",IF(H{row}>=33,"D","F"))))))'
        ws[f'S{row}'].alignment = Alignment(horizontal='center')
        
        # Islamic History (100 marks) - Column I, Grade in Column T
        ws[f'T{row}'] = f'=IF(I{row}>=80,"A+",IF(I{row}>=70,"A",IF(I{row}>=60,"A-",IF(I{row}>=50,"B",IF(I{row}>=40,"C",IF(I{row}>=33,"D","F"))))))'
        ws[f'T{row}'].alignment = Alignment(horizontal='center')
        
        # ICT (100 marks) - Column J, Grade in Column U
        ws[f'U{row}'] = f'=IF(J{row}>=80,"A+",IF(J{row}>=70,"A",IF(J{row}>=60,"A-",IF(J{row}>=50,"B",IF(J{row}>=40,"C",IF(J{row}>=33,"D","F"))))))'
        ws[f'U{row}'].alignment = Alignment(horizontal='center')
        
        # Mantiq (100 marks) - Column K, Grade in Column V
        ws[f'V{row}'] = f'=IF(K{row}>=80,"A+",IF(K{row}>=70,"A",IF(K{row}>=60,"A-",IF(K{row}>=50,"B",IF(K{row}>=40,"C",IF(K{row}>=33,"D","F"))))))'
        ws[f'V{row}'].alignment = Alignment(horizontal='center')
        
        # Career Education (100 marks) - Column L, Grade in Column W
        ws[f'W{row}'] = f'=IF(L{row}>=33,"Pass","Fail")'
        ws[f'W{row}'].alignment = Alignment(horizontal='center')
        
        # Physical Education (100 marks) - Column M, Grade in Column X
        ws[f'X{row}'] = f'=IF(M{row}>=33,"Pass","Fail")'
        ws[f'X{row}'].alignment = Alignment(horizontal='center')
        
        # Total: Sum of compulsory subjects only (C:J, excluding Mantiq K and continuous assessment L,M)
        # Compulsory: Quran(C), Arabic(D), Aqaid(E), English(F), Bangla(G), Math(H), Islamic History(I), ICT(J)
        ws[f'Y{row}'] = f'=SUM(C{row}:J{row})'
        
        # Average: Average of compulsory subjects
        ws[f'Z{row}'] = f'=ROUND(Y{row}/8,2)'
        ws[f'Z{row}'].number_format = '0.00'
        
        # GPA: Dakhil curriculum calculation
        # Helper function to calculate grade point for each subject
        # For 100-mark subjects: percentage-based
        # For 200-mark subjects: percentage-based  
        # C=Quran(200), D=Arabic(200), E=Aqaid(100), F=English(200), G=Bangla(200), H=Math(100), I=Islamic History(100), J=ICT(100)
        
        # Create grade point calculation for each subject
        # GP formula for 100-mark subject: IF(marks>=80,5,IF(marks>=70,4,IF(marks>=60,3.5,IF(marks>=50,3,IF(marks>=40,2,IF(marks>=33,1,0))))))
        # GP formula for 200-mark subject: IF(marks/200*100>=80,5,IF(marks/200*100>=70,4,IF(marks/200*100>=60,3.5,IF(marks/200*100>=50,3,IF(marks/200*100>=40,2,IF(marks/200*100>=33,1,0))))))
        
        gp_quran = f'IF(C{row}/200*100>=80,5,IF(C{row}/200*100>=70,4,IF(C{row}/200*100>=60,3.5,IF(C{row}/200*100>=50,3,IF(C{row}/200*100>=40,2,IF(C{row}/200*100>=33,1,0))))))'
        gp_arabic = f'IF(D{row}/200*100>=80,5,IF(D{row}/200*100>=70,4,IF(D{row}/200*100>=60,3.5,IF(D{row}/200*100>=50,3,IF(D{row}/200*100>=40,2,IF(D{row}/200*100>=33,1,0))))))'
        gp_aqaid = f'IF(E{row}>=80,5,IF(E{row}>=70,4,IF(E{row}>=60,3.5,IF(E{row}>=50,3,IF(E{row}>=40,2,IF(E{row}>=33,1,0))))))'
        gp_english = f'IF(F{row}/200*100>=80,5,IF(F{row}/200*100>=70,4,IF(F{row}/200*100>=60,3.5,IF(F{row}/200*100>=50,3,IF(F{row}/200*100>=40,2,IF(F{row}/200*100>=33,1,0))))))'
        gp_bangla = f'IF(G{row}/200*100>=80,5,IF(G{row}/200*100>=70,4,IF(G{row}/200*100>=60,3.5,IF(G{row}/200*100>=50,3,IF(G{row}/200*100>=40,2,IF(G{row}/200*100>=33,1,0))))))'
        gp_math = f'IF(H{row}>=80,5,IF(H{row}>=70,4,IF(H{row}>=60,3.5,IF(H{row}>=50,3,IF(H{row}>=40,2,IF(H{row}>=33,1,0))))))'
        gp_history = f'IF(I{row}>=80,5,IF(I{row}>=70,4,IF(I{row}>=60,3.5,IF(I{row}>=50,3,IF(I{row}>=40,2,IF(I{row}>=33,1,0))))))'
        gp_ict = f'IF(J{row}>=80,5,IF(J{row}>=70,4,IF(J{row}>=60,3.5,IF(J{row}>=50,3,IF(J{row}>=40,2,IF(J{row}>=33,1,0))))))'
        gp_mantiq = f'IF(K{row}>=80,5,IF(K{row}>=70,4,IF(K{row}>=60,3.5,IF(K{row}>=50,3,IF(K{row}>=40,2,IF(K{row}>=33,1,0))))))'
        
        # Check if any compulsory subject or continuous assessment failed
        fail_check = f'OR(C{row}<66,D{row}<66,E{row}<33,F{row}<66,G{row}<66,H{row}<33,I{row}<33,J{row}<33,L{row}<33,M{row}<33)'
        
        # Base GPA = average of 8 compulsory subjects
        base_gpa = f'({gp_quran}+{gp_arabic}+{gp_aqaid}+{gp_english}+{gp_bangla}+{gp_math}+{gp_history}+{gp_ict})/8'
        
        # Additional subject bonus: If Mantiq GP >= 2, add (Mantiq GP - 2) / 8
        mantiq_bonus = f'IF({gp_mantiq}>=2,({gp_mantiq}-2)/8,0)'
        
        # Final GPA formula
        gpa_formula = f'=IF({fail_check},0,MIN(5,ROUND({base_gpa}+{mantiq_bonus},2)))'
        
        ws[f'AA{row}'] = gpa_formula
        ws[f'AA{row}'].number_format = '0.00'
        
        # Overall Grade: Based on GPA value
        grade_formula = f'=IF(AA{row}>=5,"A+",IF(AA{row}>=4,"A",IF(AA{row}>=3.5,"A-",IF(AA{row}>=3,"B",IF(AA{row}>=2,"C",IF(AA{row}>=1,"D","F"))))))'
        ws[f'AB{row}'] = grade_formula
        
        # Center align GPA and Overall Grade
        ws[f'AA{row}'].alignment = Alignment(horizontal='center')
        ws[f'AB{row}'].alignment = Alignment(horizontal='center')
    
    # Style Data Source sheet
    style_data_source_sheet(ws, df)
    
    # Create Dashboard sheet
    create_dashboard_sheet(wb, df)
    
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