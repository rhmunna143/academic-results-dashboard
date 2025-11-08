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
    
    df = pd.DataFrame(data)
    
    # Calculate Total and Average
    subjects = ['Bangla', 'English', 'Mathematics', 'ICT', 'Physics', 'Chemistry', 'Biology']
    df['Total'] = df[subjects].sum(axis=1)
    df['Average'] = df[subjects].mean(axis=1).round(2)
    
    # Calculate GPA and Grade
    df['GPA'] = df.apply(calculate_gpa, axis=1)
    df['Grade'] = df['GPA'].apply(calculate_letter_grade)
    
    return df

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
    subjects_cols = ['C', 'D', 'E', 'F', 'G', 'H', 'I']  # Subject columns
    
    for row in range(2, len(df) + 2):
        for col in subjects_cols:
            cell = ws[f'{col}{row}']
            value = cell.value
            
            if value >= 80:
                cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
            elif value >= 60:
                cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            elif value >= 40:
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
    
    # Summary Statistics
    ws['A4'] = 'SUMMARY STATISTICS'
    ws['A4'].font = Font(size=14, bold=True)
    
    summaries = [
        ('B5', 'Total Students:', len(df)),
        ('E5', 'Average GPA:', f"{df['GPA'].mean():.2f}"),
        ('H5', 'Highest Total:', df['Total'].max()),
        ('K5', 'Pass Rate:', f"{(df['GPA'] > 0).sum() / len(df) * 100:.1f}%")
    ]
    
    for cell, label, value in summaries:
        ws[cell] = label
        ws[cell].font = Font(bold=True)
        next_cell = ws.cell(row=ws[cell].row, column=ws[cell].column + 1)
        next_cell.value = value
        next_cell.font = Font(size=12, bold=True, color="1F4E78")
    
    # Grade Distribution Table
    ws['A7'] = 'GRADE DISTRIBUTION'
    ws['A7'].font = Font(size=12, bold=True)
    
    grade_dist = df['Grade'].value_counts().sort_index()
    ws['A8'] = 'Grade'
    ws['B8'] = 'Count'
    
    row = 9
    for grade, count in grade_dist.items():
        ws[f'A{row}'] = grade
        ws[f'B{row}'] = count
        row += 1
    
    # Subject-wise Average
    ws['D7'] = 'SUBJECT-WISE AVERAGE'
    ws['D7'].font = Font(size=12, bold=True)
    
    subjects = ['Bangla', 'English', 'Mathematics', 'ICT', 'Physics', 'Chemistry', 'Biology']
    ws['D8'] = 'Subject'
    ws['E8'] = 'Average'
    
    row = 9
    for subject in subjects:
        ws[f'D{row}'] = subject
        ws[f'E{row}'] = round(df[subject].mean(), 2)
        row += 1
    
    # Top 5 Students
    ws['G7'] = 'TOP 5 STUDENTS'
    ws['G7'].font = Font(size=12, bold=True)
    
    top_students = df.nlargest(5, 'GPA')[['Name', 'GPA']]
    ws['G8'] = 'Rank'
    ws['H8'] = 'Name'
    ws['I8'] = 'GPA'
    
    row = 9
    rank = 1
    for _, student in top_students.iterrows():
        ws[f'G{row}'] = rank
        ws[f'H{row}'] = student['Name']
        ws[f'I{row}'] = student['GPA']
        rank += 1
        row += 1
    
    # Create Pie Chart - Grade Distribution
    pie = PieChart()
    labels = Reference(ws, min_col=1, min_row=9, max_row=8 + len(grade_dist))
    data = Reference(ws, min_col=2, min_row=8, max_row=8 + len(grade_dist))
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
    print(f"\n📊 Summary:")
    print(f"   - Total Students: {len(df)}")
    print(f"   - Average Class GPA: {df['GPA'].mean():.2f}")
    print(f"   - Students with A+: {(df['Grade'] == 'A+').sum()}")
    print(f"   - Pass Rate: {(df['GPA'] > 0).sum() / len(df) * 100:.1f}%")
    print(f"\n📁 File saved as: {filename}")

if __name__ == "__main__":
    generate_excel_file()