# VBA-automation

// If your organization is using an Excel-based control tracker, list, or matrix and 
// manually updates leadsheets by copying the information over into each workpaper,
// here's a minimum-viable macro solution built using VBA code

// I have highlighted the lines of code in the txt file that need to be reconfigured depending
// on how you want to implement this. The ranges are hardcoded at the moment but those
// can be replaced with FileDialog and InputBox options that provide a safe interface for a user to
// interact with instead of fudging with the code

// FYI
// workbook refers to the actual excel file (.xlsx)
// worksheet refers to a sheet in the workbook

// The Macro requires the following files:
// 1. List/Matrix/Table of controls (Tracker, RCM, DRL, etc) - This is an .xlsm file where the macro is stored
// 2. Workpaper template that is formatted - This is an .xlsx file

// The code logic is as follows:

Macro 1
1. Loop through every record in the list of controls
  A. Copy template workbook 
  B. Rename workbook as "Control ID" + "Short Description"
  C. Save workbook in Output folder

If Macro 1 successfully finishes, prompt message box that says "Your files have been created"
User clicking "OK" triggers Macro 2

Macro 2
1. Loop through every workbook in Output folder until
  A. Delimit workbook name/title to retrieve the Control ID
  B. Define a "row" variable that returns the row index of where the Control ID is in the table workbook
  C. Loop through every field in the Control Summary section in the active workpaper
    i) If field is found in the table header row of the table workbook (if field is found as a column in the table)
      a- Define a "column" variable that returns the column index of where the field is in the Excel table
      b- Populate the cell to the right of the field cell by indexing the row column
        // this is basically saying, if the field can be referenced in the table workbook,
        // return the value of the position of the cell in the table that corresponds to 
        // the current record (control) and field in the adjacent cell of the field in the workpaper
    ii) If field is not found, leave cell blank

When Macro 2 successfully finishes, prompt message box that says "Your leadsheets have been updated"
