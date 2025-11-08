# 📊 Academic Results Dashboard Generator

An automated Excel dashboard generator for academic results with dynamic visualizations, GPA calculations, and performance analytics.

## 🎯 Features

- **Automated Excel Generation**: Creates a complete `.xlsx` file with formulas and formatting
- **100% Dynamic Dashboard**: All calculations use Excel formulas - edit data directly in Excel and everything updates automatically!
- **Three Sheets**:
  - 📋 **Data Source**: Student records with Excel formula-based calculations (Total, Average, GPA, Grade)
  - 📊 **Dashboard**: Live charts and statistics that update when data changes
  - 🔄 **Pivot**: Placeholder for manual pivot table creation
- **Dynamic Calculations** (via Excel formulas):
  - Total marks per student (SUM)
  - Subject-wise averages (AVERAGE)
  - GPA calculation with automatic fail detection (complex IF/OR formulas)
  - Letter grade assignment (IF formulas)
  - Pass rate percentage (COUNTIF)
  - Top 5 students ranking (LARGE + INDEX/MATCH)
- **Visual Analytics** (auto-updating):
  - Grade distribution pie chart
  - Subject-wise performance bar chart
  - Top 5 students ranking table
  - Conditional color coding for marks
- **Professional Formatting**:
  - Number formats (2 decimal places for GPA/averages)
  - Percentage format for pass rate
  - Center-aligned data
  - Bold headers with styling

## 📚 Grading Scale

| Grade | Marks Range | GPA |
|-------|-------------|-----|
| A+ | 80-100 | 5.00 |
| A | 70-79 | 4.00 |
| A- | 60-69 | 3.50 |
| B | 50-59 | 3.00 |
| C | 40-49 | 2.00 |
| D | 33-39 | 1.00 |
| F | Below 33 | 0.00 |

**Note**: If any subject score is below 33, the student receives GPA 0.00 (Fail)

## 🚀 Quick Start

### Prerequisites

- Python 3.7 or higher
- pip package manager

### Installation

1. **Clone the repository**:
```bash
git clone https://github.com/rhmunna143/academic-results-dashboard.git
cd academic-results-dashboard
```

2. **Install dependencies**:
```bash
pip install -r requirements.txt
```

### Usage

**Generate the Excel file**:
```bash
python generate_excel.py
```

This will create `Academic_Results_Dashboard.xlsx` in the current directory.

### ✨ Dynamic Dashboard Features

The generated Excel workbook is **100% DYNAMIC** with zero Python dependencies after generation:

#### 🔄 What Updates Automatically:

**Data Source Sheet:**
- ✅ Total (SUM formula)
- ✅ Average (ROUND + AVERAGE formulas)
- ✅ GPA (Complex nested IF with fail detection)
- ✅ Grade (IF formula based on GPA thresholds)

**Dashboard Sheet:**
- ✅ Total Students count (COUNTA)
- ✅ Average Class GPA (AVERAGE)
- ✅ Highest Total score (MAX)
- ✅ Pass Rate % (COUNTIF with percentage formatting)
- ✅ Grade Distribution for all 7 grades (COUNTIF)
- ✅ Subject-wise Averages for all 7 subjects (AVERAGE)
- ✅ Top 5 Students ranking (LARGE + INDEX/MATCH)
- ✅ Grade Distribution Pie Chart (linked to formulas)
- ✅ Subject Performance Bar Chart (linked to formulas)

#### 🎯 How to Use Dynamic Features:

**Test 1 - Edit Marks:**
1. Open `Academic_Results_Dashboard.xlsx`
2. Go to `Data Source` sheet
3. Change any student's subject mark
4. Watch Total, Average, GPA, Grade update instantly
5. Go to `Dashboard` sheet - all stats and charts update!

**Test 2 - Make Student Fail:**
1. Change any subject mark below 33
2. GPA automatically becomes 0.00
3. Grade automatically becomes "F"
4. Dashboard pass rate updates

**Test 3 - Create Top Student:**
1. Change all marks to 90+
2. GPA becomes 5.00, Grade becomes "A+"
3. Student appears in Top 5 list
4. Dashboard statistics update

**No Python re-run needed!** All updates happen in Excel using formulas. 🚀

**Pro Tip:** Press `F9` in Excel if formulas don't recalculate immediately.

## 📋 Subjects Included

1. Bangla
2. English
3. Mathematics
4. ICT (Information & Communication Technology)
5. Physics
6. Chemistry
7. Biology

Maximum marks per subject: **100**

## 📊 Dashboard Components

### Summary Statistics
- Total number of students
- Average class GPA
- Highest total score
- Pass rate percentage

### Visualizations
1. **Grade Distribution** (Pie Chart)
2. **Subject-wise Average Scores** (Bar Chart)
3. **Top 5 Students** (Ranked Table)

### Color Coding
- 🟢 **Green** (80-100): Excellent
- 🟡 **Yellow** (60-79): Good
- 🟠 **Orange** (40-59): Average
- 🔴 **Red** (Below 40): Needs Improvement

## 🛠️ Customization

### Modify Student Data

Edit the `create_data_source()` function in `generate_excel.py`:

```python
data = {
    'SL': list(range(1, 21)),
    'Name': ['Your', 'Student', 'Names', ...],
    'Bangla': [85, 90, 72, ...],
    # ... add your data
}
```

### Change Grading Scale

Modify the `calculate_grade_point()` and `calculate_letter_grade()` functions:

```python
def calculate_grade_point(marks):
    if marks >= 90: return 5.0  # Your custom scale
    elif marks >= 80: return 4.5
    # ... customize as needed
```

### Add More Subjects

1. Add column to the data dictionary
2. Update the `subjects` list in calculations
3. Adjust column widths in styling functions

## 📦 File Structure

```
academic-results-dashboard/
│
├── generate_excel.py       # Main script
├── requirements.txt        # Python dependencies
├── README.md              # Documentation
└── Academic_Results_Dashboard.xlsx  # Generated output
```

## 🔧 Technical Details

### Libraries Used
- **pandas**: Data manipulation and calculations
- **openpyxl**: Excel file creation and styling
  - Chart creation (Pie, Bar charts)
  - Conditional formatting
  - Cell styling and formatting

### GPA Calculation Logic
```python
1. Check if any subject < 33 → GPA = 0.00 (Fail)
2. Calculate grade point for each subject
3. Average all grade points
4. Round to 2 decimal places
5. Cap at maximum 5.00
```

## 🎨 Excel Features

✅ Auto-calculated formulas  
✅ Conditional formatting  
✅ Interactive charts  
✅ Professional styling  
✅ Merged cells for headers  
✅ Custom column widths  
✅ Color-coded performance indicators  

## 📝 Sample Output

The generated Excel file contains **20 sample students** with realistic data across all subjects.

### Example Student Record:
| SL | Name | Bangla | English | Math | ICT | Physics | Chemistry | Biology | Total | Average | GPA | Grade |
|----|------|--------|---------|------|-----|---------|-----------|---------|-------|---------|-----|-------|
| 1 | Ahmed Rahman | 85 | 78 | 92 | 88 | 75 | 82 | 80 | 580 | 82.86 | 4.71 | A+ |

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📄 License

This project is open source and available under the [MIT License](LICENSE).

## 👨‍💻 Author

**Md Rabbiul Hassan Munna**  
GitHub: [@rhmunna143](https://github.com/rhmunna143)

## 🙏 Acknowledgments

- Built for educational institutions
- Designed for easy customization
- Automated grading system following standard academic practices

## 📞 Support

For issues or questions, please open an issue on GitHub or contact the author.

---

**⭐ If you find this useful, please give it a star!**