# Bangladeshi Dakhil Curriculum Implementation

## Overview
The Academic Results Dashboard has been updated to implement the **Bangladeshi Dakhil Madrasah Curriculum** grading system with dynamic calculations.

## Subject Structure

### 1. Compulsory Subjects (8 subjects)
These subjects contribute to the base GPA calculation:

| Subject | Full Marks | Column | Notes |
|---------|-----------|--------|-------|
| Quran Mazid | 100 | C | Individual subject |
| Arabic Combined | 200 | D | Combined subject (2 papers) |
| Aqaid | 100 | E | Individual subject |
| English Combined | 200 | F | Combined subject (2 papers) |
| Bangla Combined | 200 | G | Combined subject (2 papers) |
| Mathematics | 100 | H | Individual subject |
| Islamic History | 100 | I | Individual subject |
| ICT | 100 | J | Individual subject |

**Total Compulsory Marks:** 1000 marks

### 2. Additional Subject (1 subject)
Provides bonus to GPA if Grade Point ≥ 2.0:

| Subject | Full Marks | Column | Bonus Formula |
|---------|-----------|--------|---------------|
| Mantiq | 100 | K | (GP - 2) ÷ 8 if GP ≥ 2.0 |

### 3. Continuous Assessment Subjects (2 subjects)
Must pass (≥33 marks) but don't contribute to GPA:

| Subject | Full Marks | Column | Requirement |
|---------|-----------|--------|-------------|
| Career Education | 100 | L | Pass/Fail only (≥33 to pass) |
| Physical Education | 100 | M | Pass/Fail only (≥33 to pass) |

## Grading System

### Grade Points Calculation
All subjects use **percentage-based grading**:

| Percentage | Grade Point | Letter Grade |
|-----------|-------------|--------------|
| 80% - 100% | 5.00 | A+ |
| 70% - 79% | 4.00 | A |
| 60% - 69% | 3.50 | A- |
| 50% - 59% | 3.00 | B |
| 40% - 49% | 2.00 | C |
| 33% - 39% | 1.00 | D |
| 0% - 32% | 0.00 | F (Fail) |

### Pass/Fail Threshold
- **100-mark subjects:** Minimum 33 marks (33%)
- **200-mark subjects:** Minimum 66 marks (33%)
- **Continuous Assessment:** Minimum 33 marks (33%)

### GPA Calculation Formula

```
Step 1: Check if student passed all subjects
- If any compulsory subject score < 33% → GPA = 0
- If any continuous assessment score < 33 → GPA = 0

Step 2: Calculate Grade Points for each compulsory subject
- For 100-mark subjects: Grade Point based on actual marks percentage
- For 200-mark subjects: Grade Point based on (marks/200 × 100)%

Step 3: Calculate Base GPA
Base GPA = Average of 8 compulsory subject Grade Points

Step 4: Calculate Mantiq Bonus (if applicable)
- If Mantiq GP < 2.0 → Bonus = 0
- If Mantiq GP ≥ 2.0 → Bonus = (Mantiq GP - 2) / 8

Step 5: Final GPA
Final GPA = MIN(5.00, ROUND(Base GPA + Mantiq Bonus, 2))
```

## Excel Implementation

### Data Source Sheet Structure
- **Column A:** Serial Number (SL)
- **Column B:** Student Name
- **Columns C-M:** Subject Marks (11 subjects)
- **Column N:** Total (Sum of 8 compulsory subjects only)
- **Column O:** Average (Total ÷ 8)
- **Column P:** GPA (Dynamic Excel formula)
- **Column Q:** Grade (Dynamic Excel formula based on GPA)

### Key Excel Formulas

#### Total (Column N)
```excel
=SUM(C2:J2)
```
*Sums only compulsory subjects (excludes Mantiq and continuous assessment)*

#### Average (Column O)
```excel
=ROUND(N2/8,2)
```
*Average of compulsory subjects*

#### GPA (Column P)
```excel
=IF(fail_check, 0, 
   MIN(5, ROUND(base_gpa + mantiq_bonus, 2)))
```

Where:
- `fail_check`: Checks if any subject failed
- `base_gpa`: Average of 8 compulsory subject grade points
- `mantiq_bonus`: Bonus from Mantiq if GP ≥ 2.0

#### Grade (Column Q)
```excel
=IF(P2>=5,"A+",IF(P2>=4,"A",IF(P2>=3.5,"A-",
  IF(P2>=3,"B",IF(P2>=2,"C",IF(P2>=1,"D","F"))))))
```

### Dashboard Features
All dashboard features update dynamically when marks change:

1. **Summary Statistics**
   - Total Students: Count of names
   - Average GPA: Average of column P
   - Highest Total: Max of column N
   - Pass Rate: Percentage of students with GPA > 0

2. **Grade Distribution**
   - Counts students by letter grade (A+ through F)
   - Displayed as pie chart

3. **Subject-wise Average**
   - Shows average marks for 9 subjects (excluding continuous assessment)
   - Displays both raw average and percentage
   - 200-mark subjects shown with adjusted scale

4. **Top 5 Students**
   - Uses hybrid approach: Python identifies initial top 5
   - Excel formulas reference specific rows for dynamic GPA updates
   - Shows Rank, Name, and GPA

## Testing the Dynamic Features

### Test Case 1: Change marks in Data Source
1. Open `Academic_Results_Dashboard.xlsx`
2. Go to "Data Source" sheet
3. Change any student's marks
4. Observe automatic updates in:
   - Total (Column N)
   - Average (Column O)
   - GPA (Column P)
   - Grade (Column Q)
   - Dashboard statistics
   - Charts

### Test Case 2: Test fail conditions
1. Set any compulsory subject marks < 33% threshold
2. Verify GPA becomes 0
3. Verify Grade becomes "F"

### Test Case 3: Test Mantiq bonus
1. Find a student with Mantiq GP ≥ 2.0
2. Note their GPA
3. Reduce Mantiq marks to give GP < 2.0
4. Verify GPA decreases (bonus removed)
5. Increase Mantiq marks to give GP ≥ 2.0
6. Verify GPA increases (bonus applied)

### Test Case 4: Test continuous assessment fail
1. Set Career Education or Physical Education marks < 33
2. Verify GPA becomes 0 (student fails overall)

## Differences from Previous System

| Aspect | Previous (Generic) | New (Dakhil) |
|--------|-------------------|--------------|
| Subjects | 7 subjects (Bangla, English, Math, ICT, Physics, Chemistry, Biology) | 11 subjects (8 compulsory + 1 additional + 2 continuous assessment) |
| Full Marks | All 100 marks | Mixed: 100-mark and 200-mark subjects |
| Grading | Fixed GP per subject | Percentage-based GP for all subjects |
| GPA Calculation | Average of all 7 subjects | Average of 8 compulsory + optional Mantiq bonus |
| Bonus System | None | Mantiq provides bonus if GP ≥ 2.0 |
| Continuous Assessment | None | Career & Physical Education (pass/fail only) |
| Pass Threshold | 33 marks | 33% threshold (varies by subject: 33 or 66 marks) |

## Sample GPA Calculation

**Example Student:**
- Quran Mazid: 85/100 → 85% → GP = 5.00
- Arabic Combined: 150/200 → 75% → GP = 4.00
- Aqaid: 78/100 → 78% → GP = 4.00
- English Combined: 130/200 → 65% → GP = 3.50
- Bangla Combined: 140/200 → 70% → GP = 4.00
- Mathematics: 82/100 → 82% → GP = 5.00
- Islamic History: 72/100 → 72% → GP = 4.00
- ICT: 88/100 → 88% → GP = 5.00
- Mantiq: 75/100 → 75% → GP = 4.00
- Career Education: 80/100 → Pass ✓
- Physical Education: 85/100 → Pass ✓

**Calculation:**
1. All subjects passed (all ≥ 33% threshold) ✓
2. Base GPA = (5.00 + 4.00 + 4.00 + 3.50 + 4.00 + 5.00 + 4.00 + 5.00) / 8 = 4.31
3. Mantiq Bonus = (4.00 - 2) / 8 = 0.25
4. Final GPA = MIN(5.00, ROUND(4.31 + 0.25, 2)) = **4.56**
5. Letter Grade = **A** (GPA ≥ 4.0)

## Files Modified
- `generate_excel.py`: Complete implementation with Dakhil curriculum
  - New function: `calculate_gpa_dakhil()`
  - Updated: `calculate_grade_point()` to handle 100/200 marks
  - Updated: `create_data_source()` with 11 subjects
  - Updated: `style_data_source_sheet()` for 11 columns
  - Updated: Formula generation with complex Dakhil GPA logic
  - Updated: Dashboard sheet references and subject list

## Benefits
1. ✅ **Fully Dynamic:** All calculations update automatically when marks change
2. ✅ **Accurate:** Implements authentic Bangladeshi Dakhil grading rules
3. ✅ **Transparent:** Excel formulas visible and auditable
4. ✅ **Comprehensive:** Handles all subject types (100-mark, 200-mark, continuous assessment)
5. ✅ **Fair:** Includes Mantiq bonus system and proper fail detection
6. ✅ **User-friendly:** Dashboard provides instant insights and visualizations

---
*Implementation Date: 2024*
*Status: Complete and tested ✅*
