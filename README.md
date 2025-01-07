## The PhoneFix.ps1 script is designed to format phone numbers in an Excel file. Here's a brief description of its functionality:

- Input: It takes a file path as a mandatory parameter ($FilePath).
File Check: It first checks if the specified file exists, exiting if it doesn't.
- Module Installation: It checks for the presence of the ImportExcel PowerShell module. If not found, it attempts to install it, importing it otherwise.
- Backup: Before making changes, it creates a backup of the original Excel file with a timestamp in the filename.
- Phone Number Formatting: It defines a function Format-PhoneNumber that formats a 10-digit phone number into the format (XXX) XXX-XXXX, leaving other numbers unchanged.
- Sheet Processing: 
Opens the Excel file to get all sheet names.
For each sheet, it reads the data, searches for columns named 'PhoneNumber' or 'phonenumber', formats the numbers if they match the 10-digit criteria, and writes changes back to the file.
- Output: 
It informs the user whether changes were made or not, and where the backup is stored.
If changes were made, it displays a sample row from the first sheet after updating to confirm the changes.

This script ensures data integrity by backing up the original file before modification and provides feedback on the operations performed.

## This PowerShell script, UserFix.ps1, is designed to modify Excel spreadsheets by appending the AgencyID to the UserID column where the UserID is numeric. Here's a breakdown of its functionality:

- Parameters:
It takes one mandatory parameter, $FilePath, which specifies the path to the Excel file to be processed.
- File Handling:
The script uses the Resolve-Path cmdlet to get the full path of the input file, ensuring the file exists before processing.
- Excel Manipulation:
Utilizes the Excel COM Object to interact with Excel files. This allows for more complex manipulation of Excel data without requiring external libraries like Excel Interop.
- Processing Logic:
The script opens the Excel workbook and iterates through each worksheet. For each worksheet:
It searches for columns named "AgencyID" and "UserID" (with some flexibility in naming due to case-insensitive and space-insensitive matching).
If these columns are found, it processes each row starting from the second row (assuming the first row is headers):
It concatenates AgencyID and UserID only if UserID is purely numeric, creating a new UserID in the format AgencyIDUserID.
- Logging and Diagnostics:
A diagnostic log file is created or cleared, logging actions like file processing details, column headers, and row modifications for debugging purposes.
- Error Handling and Cleanup:
The script includes error handling to catch and report issues during execution. It also ensures proper cleanup of COM objects to avoid memory leaks.
- Backup Creation:
Before saving changes, the script creates a timestamped backup of the original Excel file, ensuring data integrity.
- Output:
Messages are written to the console indicating success, backup creation, or errors encountered during the process.

Overall, this script is useful for data normalization or standardization tasks within Excel spreadsheets in an organizational context where user IDs need to be prefixed with an agency identifier for consistency or tracking purposes.

## The description was created using Grok to provide an insightful overview of the script's functionality.
