# Detailed Program and Function Flow

This document provides a detailed step-by-step flow of the Python scripts used to generate the final `activity_tracker_formatted_2.xlsx` file.

## Step 1: Initial Excel File Creation

**File:** `create_excel.py`

**Purpose:** To create the initial `activity_tracker.xlsx` file with basic structure and the "Assigned By" dropdown.

**Execution Flow:**

1.  **`openpyxl.Workbook()`**: A new, empty Excel workbook is created in memory.
2.  **`wb.active`**: The active worksheet is selected.
3.  **`ws.title = "Activities"`**: The worksheet is renamed to "Activities".
4.  **`ws.append(headers)`**: The header row with ["Serial No", "Start Date", "End Date", "Activity", "Assigned By"] is added to the worksheet.
5.  **`assignee_list = [...]`**: A Python list `assignee_list` is created containing the names for the dropdown.
6.  **`DataValidation(...)`**: A `DataValidation` object is created with the `assignee_list` as a list-based validation rule.
7.  **`ws.add_data_validation(dv)`**: The data validation rule is attached to the worksheet.
8.  **`dv.add("E2:E1048576")`**: The data validation is applied to all cells in the "Assigned By" column (column E) from row 2 to the end of the sheet.
9.  **`ws[f'{col}{i}'].number_format = "YYYY-MM-DD"`**: The "Start Date" and "End Date" columns are formatted to accept dates in the "YYYY-MM-DD" format.
10. **`wb.save("activity_tracker.xlsx")`**: The workbook is saved to the disk as `activity_tracker.xlsx`.

---

## Step 2: Applying Advanced Formatting

**File:** `update_excel_formatting.py`

**Purpose:** To read the `activity_tracker.xlsx` file and apply all the advanced conditional formatting and data validation rules, saving the result as `activity_tracker_formatted_2.xlsx`.

**Execution Flow:**

1.  **`openpyxl.load_workbook("activity_tracker.xlsx")`**: The `activity_tracker.xlsx` file is loaded into memory.
2.  **`PatternFill(...)` and `Color(...)`**: Several `PatternFill` objects are created to define the background colors (green, red, yellow, and theme-based accent colors) for the conditional formatting.
3.  **Column Letter Discovery Loop (`for cell in ws[1]...`)**: The script iterates through the header row of the worksheet to find the column letters for "Status", "Start Date", "End Date", and "Assigned By". This makes the script more robust if the column order changes.
4.  **`ws.conditional_formatting.add(...)` for "Status" column**:
    -   Three `FormulaRule` objects are created and added to the worksheet's conditional formatting rules.
    -   These rules apply the green, red, and yellow fills to the "Status" column cells based on whether their value is "COMPLETE", "INCOMPLETE", or "WIP".
5.  **`DataValidation(...)` for "Status" column**: A `DataValidation` object is created and applied to the "Status" column to create a dropdown with the "COMPLETE", "INCOMPLETE", and "WIP" options.
6.  **`ws.conditional_formatting.add(...)` for other columns**:
    -   Three more `FormulaRule` objects are created and added.
    -   These rules apply the theme-based accent colors to the "Start Date", "End Date", and "Assigned By" columns if the cells in those columns are not empty (`NOT(ISBLANK(...))`).
7.  **`wb.save("activity_tracker_formatted_2.xlsx")`**: The modified workbook is saved to the disk as `activity_tracker_formatted_2.xlsx`.