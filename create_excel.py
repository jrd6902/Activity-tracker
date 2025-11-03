import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

# --- Create a new workbook and worksheet ---
# This sets up the initial Excel file in memory.
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Activities"

# --- Add headers to the worksheet ---
# Defines the column titles for the activity tracker.
headers = ["Serial No", "Start Date", "End Date", "Activity", "Assigned By"]
ws.append(headers)

# --- Create a dropdown list for the "Assigned By" column ---
# This list contains the names that will appear in the dropdown.
assignee_list = ["IT team", "Testing Team", "Prod Team", "Manager 1", "Manager 2", "Stres Test Team", "Client", "Dev team"]
# The list is converted to a comma-separated string for the data validation formula.
assignee_string = ",".join(assignee_list)

# --- Create and apply the data validation for the dropdown ---
# A DataValidation object is configured to use the list of assignees.
dv = DataValidation(type="list", formula1=f'"{assignee_string}"', allow_blank=True)
# The validation is added to the worksheet.
ws.add_data_validation(dv)
# The validation is applied to all cells in column E from row 2 downwards.
dv.add("E2:E1048576")

# --- Set the number format for date columns ---
# This ensures that the "Start Date" and "End Date" columns are formatted to display dates correctly.
for col in ["B", "C"]:
    for i in range(2, 1048577):
        ws[f'{col}{i}'].number_format = "DD-MM-YYYY"


# --- Save the workbook to a file ---
# The final workbook is saved to the disk with the specified filename.
wb.save("activity_tracker.xlsx")

print("Successfully created activity_tracker.xlsx.")