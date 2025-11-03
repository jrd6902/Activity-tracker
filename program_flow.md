# Detailed Program and Function Flow (v1.1)

This document provides a detailed step-by-step flow of the Python scripts used to generate the final `activity_tracker_formatted.xlsx` file.

## Step 1: Initial Excel File Creation

**File:** `create_excel_v1.1.py`

**Purpose:** To create the initial `activity_tracker.xlsx` file with basic structure and the "Assigned By" dropdown. The script has been refactored to use variables for better maintainability.

**Execution Flow:**

1.  **Configuration Variables**: Global variables like `OUTPUT_FILENAME`, `WORKSHEET_TITLE`, `HEADERS`, `ASSIGNEE_LIST`, etc., are defined at the top of the script.
2.  **`create_base_excel_file()` function**:
    -   **`openpyxl.Workbook()`**: A new, empty Excel workbook is created.
    -   The worksheet is created and titled using the `WORKSHEET_TITLE` variable.
    -   Headers are appended from the `HEADERS` list.
    -   A `DataValidation` object is created for the "Assigned By" dropdown using the `ASSIGNEE_LIST`.
    -   The data validation is applied to the "Assigned By" column.
    -   The "Start Date" and "End Date" columns are formatted using the `DATE_FORMAT` variable.
    -   The workbook is saved using the `OUTPUT_FILENAME` variable.
3.  **`if __name__ == "__main__":`**: This block ensures that the `create_base_excel_file()` function is called only when the script is executed directly.

---

## Step 2: Applying Advanced Formatting

**File:** `update_excel_formatting_v1.1.py`

**Purpose:** To read the `activity_tracker.xlsx` file and apply all the advanced conditional formatting and data validation rules. This script has been refactored for better structure and maintainability.

**Execution Flow:**

1.  **Configuration Variables**: Global variables for `INPUT_FILENAME`, `OUTPUT_FILENAME`, and `STATUS_OPTIONS` are defined.
2.  **Color Definitions**: `PatternFill` objects are created for all the colors used in conditional formatting.
3.  **`apply_formatting(ws)` function**:
    -   A dictionary `column_letters` is created to map column names to their letters for easy access.
    -   **Status Column Formatting**:
        -   Conditional formatting rules are added for each option in `STATUS_OPTIONS` ("COMPLETE", "INCOMPLETE", "WIP").
        -   A `DataValidation` dropdown is created for the "Status" column using the `STATUS_OPTIONS`.
    -   **Other Column Formatting**: Conditional formatting is applied to the "Start Date", "End Date", and "Assigned By" columns to highlight cells that are not empty.
4.  **`main()` function**:
    -   This function wraps the main logic of the script.
    -   It uses a `try...except FileNotFoundError` block to handle errors if the input file is not found.
    -   It calls `apply_formatting()` to apply the formatting to the worksheet.
    -   It saves the final workbook using the `OUTPUT_FILENAME` variable.
5.  **`if __name__ == "__main__":`**: This block calls the `main()` function when the script is executed.
