import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

# --- Configuration Variables ---
OUTPUT_FILENAME = "activity_tracker.xlsx"
WORKSHEET_TITLE = "Activities"
HEADERS = ["Serial No", "Start Date", "End Date", "Activity", "Assigned By", "Status"]
ASSIGNEE_LIST = ["IT team", "Testing Team", "Prod Team", "Manager 1", "Manager 2", "Stres Test Team", "Client", "Dev team"]
DATE_FORMAT = "DD-MM-YYYY"
ASSIGNED_BY_COLUMN = "E"
START_DATE_COLUMN = "B"
END_DATE_COLUMN = "C"

def create_base_excel_file():
    """
    Creates the initial Excel file with headers, dropdowns, and basic formatting.
    """
    # --- Create a new workbook and worksheet ---
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = WORKSHEET_TITLE

    # --- Add headers to the worksheet ---
    ws.append(HEADERS)

    # --- Create a dropdown list for the "Assigned By" column ---
    assignee_string = ",".join(ASSIGNEE_LIST)
    dv_assignee = DataValidation(type="list", formula1=f'"{assignee_string}"', allow_blank=True)
    ws.add_data_validation(dv_assignee)
    dv_assignee.add(f"{ASSIGNED_BY_COLUMN}2:{ASSIGNED_BY_COLUMN}1048576")

    # --- Set the number format for date columns ---
    for col in [START_DATE_COLUMN, END_DATE_COLUMN]:
        for i in range(2, 1048577):
            ws[f'{col}{i}'].number_format = DATE_FORMAT

    # --- Save the workbook to a file ---
    wb.save(OUTPUT_FILENAME)
    print(f"Successfully created {OUTPUT_FILENAME}.")

if __name__ == "__main__":
    create_base_excel_file()
