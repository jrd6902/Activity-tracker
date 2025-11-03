# Excel Activity Tracker Generator

This project contains a set of Python scripts to generate a formatted Excel file for tracking activities. The process involves two main steps to create the final, feature-rich spreadsheet.

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

First, run `create_excel.py` to generate the initial `activity_tracker.xlsx` file. This file will have the basic headers and the "Assigned By" dropdown.

```bash
python create_excel.py
```

### Step 2: Apply Advanced Formatting

Next, run `update_excel_formatting.py`. This script will load the `activity_tracker.xlsx` file and apply all the conditional formatting and the "Status" dropdown, saving the result as `activity_tracker_formatted_2.xlsx`.

```bash
python update_excel_formatting.py
```

## Prerequisites

-   Python 3.x
-   `openpyxl` library

You can install the required library using pip:
```bash
pip install openpyxl
```

## Project Files

-   `create_excel.py`: Script to create the initial Excel file.
-   `update_excel_formatting.py`: Script to apply advanced formatting.
-   `flowchart.md`: A textual flowchart of the program flow.
-   `program_flow.md`: A detailed explanation of the scripts and their functions.