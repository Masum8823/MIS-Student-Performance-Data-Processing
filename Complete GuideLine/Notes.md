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

# 7. Stage 5 — Data Validation

Now we check whether the dataset contains incorrect data.

---

# 7.1 Missing Data Check

Select the entire data range:

```text
A2:I16
```

Then go to:

```text
Home
→ Find & Select
→ Go To Special
```

A dialog box will open.

Select:

```text
Blanks
```

Then click:

```text
OK
```

### Expected Result

If no cell is selected:

```text
No missing data found.
```

If Excel selects any cells, those cells contain missing values.

---

# 7.2 Duplicate Student ID Check

Select:

```text
A2:A16
```

Then:

```text
Home
→ Conditional Formatting
→ Highlight Cells Rules
→ Duplicate Values
```

A dialog box will appear.

Keep:

```text
Duplicate
```

Then click:

```text
OK
```

Excel will highlight duplicate Student IDs.

### Expected Result

If nothing is highlighted:

```text
No duplicate Student IDs found.
```

---

# 7.3 Marks Validation

Marks are stored in:

```text
F2:H16
```

because:

```text
F = Assignment
G = Midterm
H = Final
```

Valid range:

```text
0–100
```

---

## Check Values Below 0

Select:

```text
F2:H16
```

Then:

```text
Home
→ Conditional Formatting
→ Highlight Cells Rules
→ Less Than
```

Enter:

```text
0
```

Click:

```text
OK
```

Any value below 0 will be highlighted.

---

## Check Values Above 100

Again select:

```text
F2:H16
```

Then:

```text
Home
→ Conditional Formatting
→ Highlight Cells Rules
→ Greater Than
```

Enter:

```text
100
```

Click:

```text
OK
```

Any value above 100 will be highlighted.

---

## Expected Result

If no invalid marks are found:

```text
All marks are within the valid range of 0–100.
```

---

# 7.4 Attendance Validation

Attendance is stored in:

```text
E2:E16
```

Valid range:

```text
0–100
```

Select:

```text
E2:E16
```

Then:

```text
Home
→ Conditional Formatting
→ Highlight Cells Rules
→ Less Than
```

Enter:

```text
0
```

Then repeat:

```text
Home
→ Conditional Formatting
→ Highlight Cells Rules
→ Greater Than
```

Enter:

```text
100
```

If no values are highlighted:

```text
All attendance values are within the valid range.
```

---

# 7.5 Study Hours Validation

Study Hours are stored in:

```text
I2:I16
```

Expected values:

```text
5–16 hours/week
```

---

## Check Negative Study Hours

Select:

```text
I2:I16
```

Then:

```text
Home
→ Conditional Formatting
→ Highlight Cells Rules
→ Less Than
```

Enter:

```text
0
```

Click:

```text
OK
```

---

## Check Text Values

If values such as:

```text
12 hours
10 hours
```

were entered, Excel may treat them as text.

A quick way to check is:

```text
Select I2:I16
```

Then look at the values/formulas and make sure the cells contain numeric values.

Correct:

```text
12
```

Incorrect:

```text
12 hours
```

---


# 8. Stage 6 — Data Cleaning

After validation, clean the data.

## Cleaning Rule 1 — Attendance

Incorrect:

```text
88%
```

Correct:

```text
88
```

---

## Cleaning Rule 2 — Study Hours

Incorrect:

```text
12 hours
```

Correct:

```text
12
```

---

## Cleaning Rule 3 — Duplicate IDs

Expected:

```text
No duplicate Student IDs.
```

---

## Cleaning Rule 4 — Missing Values

Expected:

```text
No missing values.
```

---

## Final Cleaning Statement

> The dataset was cleaned by standardizing attendance and study-hour values into numeric format. No missing values or duplicate student IDs were found.

---

# 9. Stage 7 — Total Score Calculation

Now we calculate the weighted total score.

Add two new columns.

## Add Total Score

Enter in:

```text
J1
```

```text
Total Score
```

## Add Grade

Enter in:

```text
K1
```

```text
Grade
```

---

# 9.1 Weight Distribution

The marks are weighted as follows:

| Component  | Weight |
| ---------- | -----: |
| Assignment |    20% |
| Midterm    |    30% |
| Final      |    50% |

Therefore:

```text
Total Score
=
Assignment × 20%
+
Midterm × 30%
+
Final × 50%
```

---

# 9.2 Excel Formula for Total Score

Click:

```text
J2
```

Enter:

```excel
=F2*20%+G2*30%+H2*50%
```

Press:

```text
Enter
```

---

## Formula Explanation

```text
F2*20%
```

means:

```text
Assignment × 20%
```

```text
G2*30%
```

means:

```text
Midterm × 30%
```

```text
H2*50%
```

means:

```text
Final × 50%
```

The three values are then added.

---

# 9.3 Copy Formula to All Students

After entering the formula in:

```text
J2
```

you will see a small square at the bottom-right corner of the cell.

This is called the **Fill Handle**.

Double-click the Fill Handle.

Excel will automatically copy the formula down:

```text
J2
J3
J4
...
J16
```

---

## Alternative Method

You can also:

1. Select J2.
2. Move the mouse to the bottom-right corner.
3. Drag down to J16.

---

# 9.4 Example Calculation

Suppose S001 has:

```text
Assignment = 82
Midterm = 78
Final = 85
```

Then:

```text
82 × 20% = 16.4

78 × 30% = 23.4

85 × 50% = 42.5
```

Therefore:

```text
Total Score
= 16.4 + 23.4 + 42.5
= 82.3
```

Excel will show:

```text
82.3
```

---

# 10. Stage 8 — Grade Calculation

The grading scale used for this project is:

| Total Score | Grade |
| ----------: | :---: |
|      80–100 |   A+  |
|       75–79 |   A   |
|       70–74 |   A-  |
|       65–69 |   B+  |
|       60–64 |   B   |
|       55–59 |   B-  |
|       50–54 |   C+  |
|       45–49 |   C   |
|       40–44 |   D   |
|    Below 40 |   F   |

> Always follow the teacher's grading scale if a different scale is provided.

---

# 10.1 Grade Formula

Click:

```text
K2
```

Enter:

```excel
=IF(J2>=80,"A+",IF(J2>=75,"A",IF(J2>=70,"A-",IF(J2>=65,"B+",IF(J2>=60,"B",IF(J2>=55,"B-",IF(J2>=50,"C+",IF(J2>=45,"C",IF(J2>=40,"D","F")))))))))
```

Press:

```text
Enter
```

---

# 10.2 How the Formula Works

Excel checks the conditions from left to right.

First:

```text
J2 >= 80
```

If true:

```text
A+
```

Otherwise:

```text
J2 >= 75
```

If true:

```text
A
```

Then:

```text
J2 >= 70
```

means:

```text
A-
```

The process continues until F.

---

## Grade Formula Logic

```text
Score >= 80 → A+
Score >= 75 → A
Score >= 70 → A-
Score >= 65 → B+
Score >= 60 → B
Score >= 55 → B-
Score >= 50 → C+
Score >= 45 → C
Score >= 40 → D
Otherwise    → F
```

---

# 10.3 Copy Grade Formula

After entering the formula in:

```text
K2
```

Double-click the Fill Handle.

The formula will automatically fill:

```text
K2:K16
```

---


# 11. Stage 9 — Data Analysis

Now we analyze the processed data.

---

# 11.1 Average Total Score

Click any empty cell.

For example:

```text
M2
```

Enter:

```excel
=AVERAGE(J2:J16)
```

Press:

```text
Enter
```

Expected result:

```text
78.01
```

---

## What Does AVERAGE Do?

```text
AVERAGE()
```

calculates the arithmetic mean.

For example:

```text
10 + 20 + 30 = 60
```

Number of values:

```text
3
```

Average:

```text
60 / 3 = 20
```

---

# 11.2 Highest Score

Click:

```text
M3
```

Enter:

```excel
=MAX(J2:J16)
```

Expected result:

```text
92.60
```

`MAX()` returns the largest value from the selected range.

---

# 11.3 Lowest Score

Click:

```text
M4
```

Enter:

```excel
=MIN(J2:J16)
```

Expected result:

```text
55.70
```

`MIN()` returns the smallest value from the selected range.

---

# 12. Stage 10 — Department-wise Analysis

Create a separate table.

For example:

| M          | N             |
| ---------- | ------------- |
| Department | Average Score |
| CSE        |               |
| EEE        |               |
| BBA        |               |

---

# 12.1 CSE Average

Click:

```text
N2
```

Enter:

```excel
=AVERAGEIF(C2:C16,"CSE",J2:J16)
```

Press:

```text
Enter
```

Expected result:

```text
70.52
```

---

# 12.2 EEE Average

Click:

```text
N3
```

Enter:

```excel
=AVERAGEIF(C2:C16,"EEE",J2:J16)
```

Expected result:

```text
82.18
```

---

# 12.3 BBA Average

Click:

```text
N4
```

Enter:

```excel
=AVERAGEIF(C2:C16,"BBA",J2:J16)
```

Expected result:

```text
84.03
```

---

# 12.4 AVERAGEIF Formula Explanation

The formula:

```excel
=AVERAGEIF(C2:C16,"CSE",J2:J16)
```

contains three parts:

```text
C2:C16
```

This is the **criteria range**.

Excel checks the department names here.

```text
"CSE"
```

This is the condition.

```text
J2:J16
```

This is the range from which the average is calculated.

Therefore:

```text
Find CSE students
        ↓
Take their Total Scores
        ↓
Calculate Average
```

---

# 13. Stage 11 — Attendance Analysis

## 13.1 Average Attendance

Click an empty cell.

For example:

```text
M6
```

Enter:

```excel
=AVERAGE(E2:E16)
```

Expected result:

```text
80.60
```

If you want to display it as:

```text
80.60%
```

you can format the cell appropriately based on how your attendance values are stored.

> Since this project stores attendance as `88`, `75`, etc., not as Excel percentage values like `88% = 0.88`, the numerical average is `80.60` percentage points.

---

# 13.2 Students Below 75% Attendance

Click:

```text
M7
```

Enter:

```excel
=COUNTIF(E2:E16,"<75")
```

Expected result:

```text
4
```

You can label it:

```text
Students Below 75% Attendance
```

---

# 13.3 COUNTIF Formula Explanation

The formula:

```excel
=COUNTIF(E2:E16,"<75")
```

means:

```text
Look at E2:E16
       ↓
Find values less than 75
       ↓
Count them
```

Therefore, if four students have attendance below 75:

```text
Result = 4
```

---

# 14. Stage 12 — A+ and A Analysis

According to the grading scale:

```text
A+ = 80+
A  = 75–79
```

Therefore:

```text
A+ or A
=
Total Score >= 75
```

---

## Formula

Click an empty cell.

Enter:

```excel
=COUNTIF(J2:J16,">=75")
```

Expected result:

```text
9
```

Therefore:

```text
9 students received A+ or A.
```

---

# 15. Stage 13 — Correlation Analysis

We now analyze the relationship between:

```text
Study Hours per Week
```

and:

```text
Total Score
```

Study Hours:

```text
I2:I16
```

Total Score:

```text
J2:J16
```

---

# 15.1 CORREL Formula

Click an empty cell.

Enter:

```excel
=CORREL(I2:I16,J2:J16)
```

Press:

```text
Enter
```

Expected result:

```text
Approximately 0.965
```

---

# 15.2 Understanding Correlation

Correlation coefficient `r` generally ranges from:

```text
-1 to +1
```

### Positive Correlation

```text
r > 0
```

As one variable increases, the other tends to increase.

### Negative Correlation

```text
r < 0
```

As one variable increases, the other tends to decrease.

### Near Zero

```text
r ≈ 0
```

There is little linear relationship.

---

## This Dataset

The result is approximately:

```text
r = 0.965
```

This indicates a strong positive relationship in this dataset.

### Report Statement

> There is a strong positive relationship between study hours and total score. In this dataset, students who study more hours per week generally have higher total scores.

---

# 16. Stage 14 — Data Visualization

We will create three charts:

```text
Chart 1 → Department-wise Average Score
Chart 2 → Grade Distribution
Chart 3 → Study Hours vs. Total Score
```

---

# 16.1 Chart 1 — Department-wise Average Score

First create this table:

| Department | Average Score |
| ---------- | ------------: |
| CSE        |         70.52 |
| EEE        |         82.18 |
| BBA        |         84.03 |

---

## Step-by-Step

### Step 1

Select:

```text
M1:N4
```

### Step 2

Go to:

```text
Insert
→ Charts
→ Column or Bar Chart
→ 2-D Column
→ Clustered Column
```

### Step 3

Click the chart title.

Change it to:

```text
Department-wise Average Score
```

---

## Result

The X-axis should show:

```text
CSE
EEE
BBA
```

The Y-axis should show:

```text
Average Score
```

---

# 17. Stage 15 — Grade Distribution

First create a grade count table.

For example:

| M     |                  N |
| ----- | -----------------: |
| Grade | Number of Students |
| A+    |                    |
| A     |                    |
| A-    |                    |
| B+    |                    |
| B     |                    |
| B-    |                    |
| C+    |                    |
| C     |                    |
| D     |                    |
| F     |                    |

---

# 17.1 Enter Grades

Enter:

```text
M8 = A+
M9 = A
M10 = A-
M11 = B+
M12 = B
M13 = B-
M14 = C+
M15 = C
M16 = D
M17 = F
```

---

# 17.2 Count Each Grade

Click:

```text
N8
```

Enter:

```excel
=COUNTIF($K$2:$K$16,M8)
```

Press:

```text
Enter
```

---

# 17.3 Why `$` Is Used

The formula contains:

```text
$K$2:$K$16
```

The `$` makes the range **absolute**.

When the formula is copied down, Excel will keep:

```text
K2:K16
```

instead of changing it.

Meanwhile:

```text
M8
```

will become:

```text
M9
M10
M11
...
```

This allows Excel to count each grade separately.

---

# 17.4 Copy the Formula

Select:

```text
N8
```

Then drag the Fill Handle down to:

```text
N17
```

Now each grade count will be calculated automatically.

---

# 17.5 Create Grade Distribution Chart

Select:

```text
M8:N17
```

Then:

```text
Insert
→ Column or Bar Chart
→ 2-D Column
→ Clustered Column
```

Change the chart title to:

```text
Grade Distribution
```

---

# 18. Stage 16 — Study Hours vs Total Score

This chart is different from a normal column chart.

We need a **Scatter Chart**.

---

# 18.1 Required Data

We need:

```text
Study Hours
```

and:

```text
Total Score
```

The columns are:

```text
I = Study Hours
J = Total Score
```

---

# 18.2 Select Data

Select:

```text
I1:J16
```

---

# 18.3 Insert Scatter Chart

Go to:

```text
Insert
→ Scatter (X, Y)
→ Scatter with only Markers
```

---

# 18.4 Chart Title

Change the title to:

```text
Study Hours vs. Total Score
```

---

# 18.5 Axis Meaning

The horizontal axis should represent:

```text
Study Hours per Week
```

The vertical axis should represent:

```text
Total Score
```

Therefore:

```text
X-axis = Study Hours
Y-axis = Total Score
```

---

# 18.6 Add Trendline

To make the relationship easier to see:

1. Click the scatter chart.
2. Click the `+` icon beside the chart.
3. Enable:

```text
Trendline
```

4. Select:

```text
Linear
```

The trendline will show the general direction of the relationship.

---