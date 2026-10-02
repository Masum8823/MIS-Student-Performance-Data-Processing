# MIS Project — Student Performance Data Processing

A complete step-by-step guide for processing student performance data using **Microsoft Excel**.

This README is designed as a practical guide. You can keep it beside your Excel file and follow each step from **data entry to final charts**.

---

# Table of Contents

* [1. Project Overview](#1-project-overview)
* [2. Complete Project Workflow](#2-complete-project-workflow)
* [3. Stage 1 — Data Collection](#3-stage-1--data-collection)
* [4. Stage 2 — Creating the Excel File](#4-stage-2--creating-the-excel-file)
* [5. Stage 3 — Data Entry](#5-stage-3--data-entry)
* [6. Stage 4 — Formatting the Dataset](#6-stage-4--formatting-the-dataset)
* [7. Stage 5 — Data Validation](#7-stage-5--data-validation)
* [8. Stage 6 — Data Cleaning](#8-stage-6--data-cleaning)
* [9. Stage 7 — Total Score Calculation](#9-stage-7--total-score-calculation)
* [10. Stage 8 — Grade Calculation](#10-stage-8--grade-calculation)
* [11. Stage 9 — Data Analysis](#11-stage-9--data-analysis)
* [12. Stage 10 — Department-wise Analysis](#12-stage-10--department-wise-analysis)
* [13. Stage 11 — Attendance Analysis](#13-stage-11--attendance-analysis)
* [14. Stage 12 — A+ and A Analysis](#14-stage-12--a-and-a-analysis)
* [15. Stage 13 — Correlation Analysis](#15-stage-13--correlation-analysis)
* [16. Stage 14 — Data Visualization](#16-stage-14--data-visualization)
* [17. Stage 15 — Grade Distribution](#17-stage-15--grade-distribution)
* [18. Stage 16 — Study Hours vs Total Score](#18-stage-16--study-hours-vs-total-score)
* [19. Final Analysis Summary](#19-final-analysis-summary)
* [20. Final Excel Structure](#20-final-excel-structure)
* [21. Excel Functions Used](#21-excel-functions-used)
* [22. Final Checklist](#22-final-checklist)

---

# 1. Project Overview

## Project Title

**Student Performance Data Processing**

## Project Objective

The objective of this project is to process raw student information and convert it into useful academic information using Microsoft Excel.

The project covers:

* Data Collection
* Data Entry
* Data Validation
* Data Cleaning
* Data Processing
* Total Score Calculation
* Grade Calculation
* Data Analysis
* Data Visualization

---

# 2. Complete Project Workflow

```text
Data Collection
       ↓
Create Excel File
       ↓
Enter Student Data
       ↓
Format Dataset
       ↓
Data Validation
       ↓
Data Cleaning
       ↓
Calculate Total Score
       ↓
Calculate Grade
       ↓
Data Analysis
       ↓
Create Charts
       ↓
Final Report
```

---

# 3. Stage 1 — Data Collection

Student performance data can come from different university systems.

## Source 1 — Student Information System (SIS)

Possible data:

* Student ID
* Student Name
* Department
* Gender

## Source 2 — Attendance Management System

Possible data:

* Attendance Percentage

## Source 3 — LMS / Academic Records

Possible data:

* Assignment Mark
* Midterm Mark
* Final Mark

## Source 4 — Student Survey

Possible data:

* Study Hours per Week

---

## Source Mapping

| Data            | Source                       |
| --------------- | ---------------------------- |
| Student ID      | Student Information System   |
| Name            | Student Information System   |
| Department      | Student Information System   |
| Gender          | Student Information System   |
| Attendance      | Attendance Management System |
| Assignment Mark | LMS / Academic Records       |
| Midterm Mark    | Exam Records                 |
| Final Mark      | Exam Records                 |
| Study Hours     | Student Survey               |

---

# 4. Stage 2 — Creating the Excel File

## Step 1 — Open Excel

Open:

```text
Microsoft Excel
```

Then select:

```text
Blank Workbook
```

---

## Step 2 — Save the File

Go to:

```text
File
→ Save As
```

Give the file a name such as:

```text
Student_Performance_Data_Processing.xlsx
```

---

# 5. Stage 3 — Data Entry

We will use **Row 1** for headers.

Enter the following:

| Cell | Header               |
| ---- | -------------------- |
| A1   | Student ID           |
| B1   | Name                 |
| C1   | Department           |
| D1   | Gender               |
| E1   | Attendance           |
| F1   | Assignment Mark      |
| G1   | Midterm Mark         |
| H1   | Final Mark           |
| I1   | Study Hours per Week |

---

## Final Header Row

```text
A1 = Student ID
B1 = Name
C1 = Department
D1 = Gender
E1 = Attendance
F1 = Assignment Mark
G1 = Midterm Mark
H1 = Final Mark
I1 = Study Hours per Week
```

---

## Student Data

Enter the 15 students from:

```text
Row 2 → Row 16
```

Therefore:

```text
First student = Row 2
Last student  = Row 16
```

---

## Important Data Entry Rule

### Attendance

Enter:

```text
88
```

Do not enter:

```text
88%
```

### Study Hours

Enter:

```text
12
```

Do not enter:

```text
12 hours
```

The reason is that Excel needs numeric values for calculations.

---

# 6. Stage 4 — Formatting the Dataset

After entering all data:

Select:

```text
A1:I16
```

Then press:

```text
Ctrl + T
```

A dialog box will appear.

Check:

```text
☑ My table has headers
```

Then click:

```text
OK
```

---

## Alternative Method

You can also use:

```text
Insert
→ Table
```

Then select the data range:

```text
A1:I16
```

and enable:

```text
My table has headers
```

---

## Recommended Header Formatting

Select:

```text
A1:I1
```

Then:

```text
Home
→ Bold
```

You can also use:

```text
Home
→ Center
```

for better alignment.

---