# Calendar Generator (Excel → Color-Coded Weekly Matrix)

A small Python + Streamlit web tool that takes **multiple student timetable Excel files** and generates a **single consolidated Excel** in a “matrix calendar” format:

- **Column A:** Day of the week  
- **Column B:** Time of the day (slot grid)  
- **Column C onward:** One column per student  
- Each student’s occupied slot is **filled with that student’s color** and contains the course name (and time range).
- Includes a **Legend** sheet mapping **student → color**.

This is useful for instructors/program coordinators to quickly visualize multiple students’ schedules in one place.

---

## Input Format

Upload one or more `.xlsx` files (one per student).  
Each file must contain a sheet with at least **4 columns**:

1. **Course name**
2. **Day of the week**
3. **Start time**
4. **End time**

Header names can vary (e.g., `Course`, `Subject`, `Day`, `Start time`, `Time From`, `End time`, etc.).  
If headers are missing or unclear, the tool falls back to using the **first 4 columns**.

### Supported day formats
- `Mon`, `Monday`, `Tue`, `Tues`, … etc.

### Supported time formats
- Excel time cells
- Strings like `08:30`, `8.30`, `08:30:00`

---

## Output

The app produces an Excel file with:

### 1) `Matrix` sheet
A weekly slot grid:

| Day | Time | student_1 | student_2 | ... |
|-----|------|-----------|-----------|-----|

Cells are:
- **Filled** with the student’s color if they have a class in that slot
- **Empty** if no class
- If a student has overlapping items in the same slot, the cell will stack text (multiple lines)

### 2) `Legend` sheet
A simple mapping of each student to their assigned color.

---

## Features

- Upload multiple student files at once
- Automatically selects the best sheet from each workbook (if multiple sheets exist)
- Skips invalid/empty sheets without crashing
- Deterministic color assignment per student (same student → same color every run)
- Adjustable time-slot resolution (15/30/60 minutes)

---

## Project Structure

