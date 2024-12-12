# PhoneFix.ps1

param (
    [Parameter(Mandatory=$true)]
    [string]$FilePath
)

# Check if the file exists
if (-not (Test-Path $FilePath)) {
    Write-Host "The specified file does not exist."
    exit
}

# Check if ImportExcel module is installed, if not, attempt to install it
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    try {
        Write-Host "ImportExcel module not found. Attempting to install..."
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
        Import-Module ImportExcel
    } catch {
        Write-Host "Failed to install ImportExcel module. Error: $_"
        exit
    }
} else {
    Import-Module ImportExcel
}

# Get current date for backup
$date = (Get-Date).ToString("yyyyMMdd")

# Backup the original file
$backupFilePath = [System.IO.Path]::ChangeExtension($FilePath, "_$date.xlsx")
Copy-Item -Path $FilePath -Destination $backupFilePath

# Function to format phone number
function Format-PhoneNumber {
    param(
        [string]$number
    )
    
    # Remove all non-digit characters
    $number = $number -replace '[^0-9]'
    
    # Check if the number has 10 digits
    if ($number.Length -eq 10) {
        return "({0}) {1}-{2}" -f $number.Substring(0,3), $number.Substring(3,3), $number.Substring(6,4)
    } else {
        return $number  # Return original if not 10 digits
    }
}

# Open the Excel file to get all sheet names
$excel = Open-ExcelPackage -Path $FilePath
$sheetNames = $excel.Workbook.Worksheets | Select-Object -ExpandProperty Name

# Process each sheet
$changesMade = $false
foreach ($sheetName in $sheetNames) {
    Write-Host "Processing sheet: $sheetName"
    $sheetData = Import-Excel -Path $FilePath -WorksheetName $sheetName
    
    foreach ($row in $sheetData) {
        if ($row.PSObject.Properties.Name -contains 'PhoneNumber' -or $row.PSObject.Properties.Name -contains 'phonenumber') {
            $originalNumber = $row.PhoneNumber
            $formattedNumber = Format-PhoneNumber $originalNumber
            if ($formattedNumber -ne $originalNumber) {
                $row.PhoneNumber = $formattedNumber
                $changesMade = $true
                Write-Host "Updated in sheet ${sheetName}: $originalNumber -> $formattedNumber"
            }
        }
    }
    # Export the modified data back to Excel for each sheet
    $sheetData | Export-Excel -Path $FilePath -WorksheetName $sheetName -AutoSize -AutoFilter -FreezeTopRow
}

if ($changesMade) {
    Write-Host "Phone numbers have been formatted and the file has been updated. A backup was created at $backupFilePath"
} else {
    Write-Host "No changes were made to the phone numbers."
}

# Check if file was successfully updated
if (Test-Path $FilePath) {
    Write-Host "Sample rows after update from the first sheet:"
    $updatedData = Import-Excel -Path $FilePath -WorksheetName $sheetNames[0] | Select-Object -First 1
    Write-Host ($updatedData | ConvertTo-Json)
}