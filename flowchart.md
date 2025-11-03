# Program Flowchart (v1.1)

This flowchart illustrates the process of generating the final formatted Excel file, `activity_tracker_formatted.xlsx`.

```
[Start]
   |
   v
[Run `create_excel_v1.1.py`]
   |
   +--> [Input: None]
   |
   +--> [Process: Creates a new Excel workbook, adds headers, and a dropdown for "Assigned By"]
   |
   v
[Output: `activity_tracker.xlsx`]
   |
   |
   v
[Run `update_excel_formatting_v1.1.py`]
   |
   +--> [Input: `activity_tracker.xlsx`]
   |
   +--> [Process: Loads the workbook and applies all conditional formatting and data validation]
   |
   v
[Output: `activity_tracker_formatted.xlsx`]
   |
   v
[End]
```
