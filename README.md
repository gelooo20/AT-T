**📁 Expected Excel Structure**
____________________________________________________________________________
Report Workbook
Must contain a worksheet named exactly:
Raw
Pivot tables should reference data from the Raw sheet
Workbook must not be opened in Protected View
CSV File
UTF-8 encoded (with or without BOM)
Comma-delimited
First row treated as headers
Must fit within Excel row limits (~1,048,576 rows)

**🔒 Safety & Controls**
____________________________________________________________________________
CSV is read by Python (never opened in Excel)
Excel alerts, events, and link updates are disabled during execution
Only the selected report workbook is opened
Formatting in the Raw sheet is preserved

**⚠️ Common Issues & Fixes**
____________________________________________________________________________
Raw sheet not found
Ensure the sheet name is exactly Raw
Sheet names are case-sensitive in Excel COM
File opens as Read-Only
Right-click file → Properties → Unblock
Open once manually, enable editing, then save
Pivot not updating
Ensure pivot source points to the Raw sheet
If using Excel Tables, confirm table auto-expands
CSV delimiter is not comma
Script assumes comma-delimited CSV
Adjust delimiter if needed

**🔧 Customization Options**
____________________________________________________________________________
This tool can be extended to:
Preserve header row only
Append data instead of replacing
Add a progress bar for large CSVs
Validate column count before loading
Refresh specific pivot tables only
Add logging or audit trails

**🧠 Maintainer Notes**
____________________________________________________________________________
Uses Excel COM (win32com) for native Excel behavior

Designed for reporting, finance, and ops workflows

Safe to distribute internally as an EXE

No external services or network access required

**📌 Versioning**
____________________________________________________________________________
v1.0
CSV → Raw load
Pivot refresh
UI status updates
EXE-ready
