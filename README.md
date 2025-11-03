# Excel Activity Tracker Generator v1.1

This project contains a set of Python scripts to generate a formatted Excel file for tracking activities. This version (v1.1) has been refactored for better maintainability and clarity.

## Features

-   **Dynamic Excel File Creation:** The project starts by creating a base Excel file with a predefined structure.
-   **Advanced Conditional Formatting:** The final Excel sheet includes:
    -   **Status Highlighting:** The "Status" column cells are colored based on their value ("COMPLETE", "INCOMPLETE", "WIP").
    -   **Data Presence Highlighting:** The "Start Date", "End Date", and "Assigned By" columns are highlighted if they contain data.
-   **Data Validation Dropdowns:**
    -   The "Assigned By" column has a dropdown with a predefined list of names.
    -   The "Status" column has a dropdown with "COMPLETE", "INCOMPLETE", and "WIP" options.

## Usage

The process involves running two scripts in sequence:

### Step 1: Create the Base Excel File

First, run `create_excel_v1.1.py` to generate the initial `activity_tracker.xlsx` file.

```bash
python create_excel_v1.1.py
```

### Step 2: Apply Advanced Formatting

Next, run `update_excel_formatting_v1.1.py`. This script will load the `activity_tracker.xlsx` file and apply all the conditional formatting and the "Status" dropdown, saving the result as `activity_tracker_formatted.xlsx`.

```bash
python update_excel_formatting_v1.1.py
```

## Prerequisites

-   Python 3.x
-   `openpyxl` library

You can install the required library using pip:
```bash
pip install openpyxl
```

## Project Files (v1.1)

-   `create_excel_v1.1.py`: Refactored script to create the initial Excel file.
-   `update_excel_formatting_v1.1.py`: Refactored script to apply advanced formatting.
-   `README.md`: This file.
-   `flowchart.md`: A flowchart of the program flow.
-   `program_flow.md`: A detailed explanation of the scripts and their functions.
