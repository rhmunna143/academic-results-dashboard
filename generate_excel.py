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

def calculate_grade_point(marks):
    """Calculate grade point for individual subject"""
    if marks >= 80: return 5.0
    elif marks >= 70: return 4.0
    elif marks >= 60: return 3.5
    elif marks >= 50: return 3.0
    elif marks >= 40: return 2.0
    elif marks >= 33: return 1.0
    else: return 0.0

def calculate_gpa(row):
    """Calculate GPA based on all subjects"""
    subjects = ['Bangla', 'English', 'Mathematics', 'ICT', 'Physics', 'Chemistry', 'Biology']
    
    # Check if any subject failed (below 33)
    for subject in subjects:
        if row[subject] < 33:
            return 0.0
    
    # Calculate average of grade points
    grade_points = [calculate_grade_point(row[subject]) for subject in subjects]
    return min(5.0, round(sum(grade_points) / len(grade_points), 2))

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
    """Create sample student data"""
    data = {
        'SL': list(range(1, 21)),
        'Name': [
            'Ahmed Rahman', 'Fatima Khan', 'Karim Hassan', 'Ayesha Begum',
            'Rahim Uddin', 'Nadia Islam', 'Jahir Ahmed', 'Sadia Sultana',
            'Tarik Hasan', 'Ruhi Akter', 'Mehedi Hassan', 'Lubna Khatun',
            'Fahim Ahmed', 'Samira Begum', 'Imran Khan', 'Rafia Islam',
            'Shakib Rahman', 'Nusrat Jahan', 'Arif Hossain', 'Tasnia Akter'
        ],
        'Bangla': [85, 90, 72, 95, 65, 82, 55, 78, 45, 88, 92, 68, 75, 80, 58, 87, 70, 84, 62, 91],
        'English': [78, 88, 68, 92, 62, 85, 58, 75, 48, 90, 86, 65, 72, 82, 55, 85, 68, 88, 60, 89],
        'Mathematics': [92, 85, 75, 98, 68, 80, 52, 82, 42, 85, 90, 70, 78, 85, 60, 88, 72, 86, 65, 93],
        'ICT': [88, 92, 70, 94, 70, 88, 60, 80, 50, 87, 88, 72, 80, 86, 62, 90, 74, 89, 68, 92],
        'Physics': [75, 87, 65, 90, 58, 83, 50, 76, 46, 89, 84, 68, 74, 81, 54, 86, 70, 85, 63, 88],
        'Chemistry': [82, 89, 71, 93, 64, 81, 54, 79, 44, 86, 87, 69, 76, 83, 56, 87, 71, 87, 64, 90],
        'Biology': [80, 91, 69, 96, 66, 84, 57, 77, 47, 91, 89, 71, 77, 84, 59, 89, 73, 88, 66, 92]
    }
    
    # Return DataFrame with only raw marks. Total/Average/GPA/Grade will be
    # inserted into the workbook as Excel formulas so the workbook stays
    # dynamic when users edit marks directly in Excel.
    df = pd.DataFrame(data)
    subjects = ['Bangla', 'English', 'Mathematics', 'ICT', 'Physics', 'Chemistry', 'Biology']
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
    
    # Apply conditional formatting colors for marks (subject columns C:I)
    # Values in these columns are raw marks so fills can be applied directly.
    subjects_cols = ['C', 'D', 'E', 'F', 'G', 'H', 'I']  # Subject columns

    for row in range(2, len(df) + 2):
        for col in subjects_cols:
            cell = ws[f'{col}{row}']
            value = cell.value
            if value is None:
                continue

            try:
                num = float(value)
            except Exception:
                continue

            if num >= 80:
                cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
            elif num >= 60:
                cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            elif num >= 40:
                cell.fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
            else:
                cell.fill = PatternFill(start_color="FC9999", end_color="FC9999", fill_type="solid")
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 18
    for col in subjects_cols:
        ws.column_dimensions[col].width = 12
    ws.column_dimensions['J'].width = 10
    ws.column_dimensions['K'].width = 10
    ws.column_dimensions['L'].width = 8
    ws.column_dimensions['M'].width = 8

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
    ws['F5'] = "=IFERROR(ROUND(AVERAGE('Data Source'!L:L),2), 0)"
    ws['H5'] = 'Highest Total:'
    ws['I5'] = "=MAX('Data Source'!J:J)"
    ws['K5'] = 'Pass Rate:'
    # Count GPA > 0 divided by total - format as percentage
    ws['L5'] = '=IF(C5>0,COUNTIF(\'Data Source\'!L:L,">0")/C5,0)'
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
        # COUNTIF on Data Source Grade column
        ws[f'B{row}'] = f"=COUNTIF('Data Source'!M:M,\"{grade}\")"
        ws[f'B{row}'].alignment = Alignment(horizontal='center')
        row += 1
    
    # Subject-wise Average
    ws['D7'] = 'SUBJECT-WISE AVERAGE'
    ws['D7'].font = Font(size=12, bold=True)
    
    subjects = ['Bangla', 'English', 'Mathematics', 'ICT', 'Physics', 'Chemistry', 'Biology']
    ws['D8'] = 'Subject'
    ws['E8'] = 'Average'
    ws['D8'].font = Font(bold=True)
    ws['E8'].font = Font(bold=True)
    ws['E8'].alignment = Alignment(horizontal='center')
    
    row = 9
    for subject in subjects:
        ws[f'D{row}'] = subject
        col_letter = {'Bangla':'C','English':'D','Mathematics':'E','ICT':'F','Physics':'G','Chemistry':'H','Biology':'I'}[subject]
        ws[f'E{row}'] = f"=IFERROR(ROUND(AVERAGE('Data Source'!{col_letter}:{col_letter}),2),0)"
        ws[f'E{row}'].number_format = '0.00'
        ws[f'E{row}'].alignment = Alignment(horizontal='center')
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
    
    # Calculate top 5 students in Python to get their row numbers
    df_with_gpa = df.copy()
    df_with_gpa['GPA_calc'] = df_with_gpa.apply(calculate_gpa, axis=1)
    df_sorted = df_with_gpa.nlargest(5, 'GPA_calc')
    top_5_rows = [df_with_gpa.index.get_loc(idx) + 2 for idx in df_sorted.index]  # +2 for header and 1-based indexing
    
    # Create Top 5 list with formulas that reference specific rows from Data Source
    row = 9
    for i, data_source_row in enumerate(top_5_rows, 1):
        ws[f'G{row}'] = i
        ws[f'G{row}'].alignment = Alignment(horizontal='center')
        
        # Reference the specific student's name and GPA from Data Source
        # This way if their marks change, their GPA updates
        ws[f'H{row}'] = f"='Data Source'!B{data_source_row}"
        ws[f'I{row}'] = f"='Data Source'!L{data_source_row}"
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
    
    # Create Bar Chart - Subject-wise Performance
    bar = BarChart()
    bar.type = "col"
    bar.title = "Subject-wise Average Scores"
    bar.y_axis.title = 'Average Marks'
    bar.x_axis.title = 'Subjects'
    
    data = Reference(ws, min_col=5, min_row=8, max_row=15)
    cats = Reference(ws, min_col=4, min_row=9, max_row=15)
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
    
    # Add Total, Average, GPA, and Grade columns with Excel formulas
    # Add headers first
    ws['J1'] = 'Total'
    ws['K1'] = 'Average'
    ws['L1'] = 'GPA'
    ws['M1'] = 'Grade'
    
    # Add formulas for each student row (rows 2 to len(df)+1)
    for row in range(2, len(df) + 2):
        # Total: Sum of all subject marks (columns C to I)
        ws[f'J{row}'] = f'=SUM(C{row}:I{row})'
        
        # Average: Average of all subject marks
        ws[f'K{row}'] = f'=ROUND(AVERAGE(C{row}:I{row}),2)'
        ws[f'K{row}'].number_format = '0.00'
        
        # GPA: Complex nested IF formula based on grade point calculation
        # First check if any subject failed (< 33), then calculate GPA
        gpa_formula = f'''=IF(OR(C{row}<33,D{row}<33,E{row}<33,F{row}<33,G{row}<33,H{row}<33,I{row}<33),0,
MIN(5,ROUND((
IF(C{row}>=80,5,IF(C{row}>=70,4,IF(C{row}>=60,3.5,IF(C{row}>=50,3,IF(C{row}>=40,2,IF(C{row}>=33,1,0))))))+
IF(D{row}>=80,5,IF(D{row}>=70,4,IF(D{row}>=60,3.5,IF(D{row}>=50,3,IF(D{row}>=40,2,IF(D{row}>=33,1,0))))))+
IF(E{row}>=80,5,IF(E{row}>=70,4,IF(E{row}>=60,3.5,IF(E{row}>=50,3,IF(E{row}>=40,2,IF(E{row}>=33,1,0))))))+
IF(F{row}>=80,5,IF(F{row}>=70,4,IF(F{row}>=60,3.5,IF(F{row}>=50,3,IF(F{row}>=40,2,IF(F{row}>=33,1,0))))))+
IF(G{row}>=80,5,IF(G{row}>=70,4,IF(G{row}>=60,3.5,IF(G{row}>=50,3,IF(G{row}>=40,2,IF(G{row}>=33,1,0))))))+
IF(H{row}>=80,5,IF(H{row}>=70,4,IF(H{row}>=60,3.5,IF(H{row}>=50,3,IF(H{row}>=40,2,IF(H{row}>=33,1,0))))))+
IF(I{row}>=80,5,IF(I{row}>=70,4,IF(I{row}>=60,3.5,IF(I{row}>=50,3,IF(I{row}>=40,2,IF(I{row}>=33,1,0))))))
)/7,2)))'''
        ws[f'L{row}'] = gpa_formula.replace('\n', '')
        ws[f'L{row}'].number_format = '0.00'
        
        # Grade: Based on GPA value
        grade_formula = f'=IF(L{row}>=5,"A+",IF(L{row}>=4,"A",IF(L{row}>=3.5,"A-",IF(L{row}>=3,"B",IF(L{row}>=2,"C",IF(L{row}>=1,"D","F"))))))'
        ws[f'M{row}'] = grade_formula
        
        # Center align GPA and Grade
        ws[f'L{row}'].alignment = Alignment(horizontal='center')
        ws[f'M{row}'].alignment = Alignment(horizontal='center')
    
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
        df_console['GPA'] = df_console.apply(calculate_gpa, axis=1)
        df_console['Grade'] = df_console['GPA'].apply(calculate_letter_grade)
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