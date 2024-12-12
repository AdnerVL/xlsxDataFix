# UserFix.ps1
# Script to modify UserID columns in Excel spreadsheets by prefixing with AgencyID
param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$FilePath
)
# Resolve the full path of the input file
$FullFilePath = Resolve-Path $FilePath -ErrorAction Stop
# Fallback method using Excel COM object
function Process-ExcelWithCOMObject {
    param([string]$FilePath)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        $workbook = $excel.Workbooks.Open($FilePath)
        # Create a log file for diagnostics
        $logPath = Join-Path (Split-Path $FilePath -Parent) "UserFix_Diagnostic_Log.txt"
        # Clear previous log if exists
        if (Test-Path $logPath) { Clear-Content $logPath }
        Add-Content -Path $logPath -Value "Diagnostic Log for UserFix Script"
        Add-Content -Path $logPath -Value "Processed File: $FilePath"
        Add-Content -Path $logPath -Value "Processing Date: $(Get-Date)`n"
        # Iterate through each worksheet
        foreach ($worksheet in $workbook.Worksheets) {
            $worksheet.Select() | Out-Null
            Add-Content -Path $logPath -Value "Worksheet Name: $($worksheet.Name)"
            # Find column indices
            $headerRow = $worksheet.UsedRange.Rows(1)
            $agencyIDCol = -1
            $userIDCol = -1
            # Log all column headers for debugging
            Add-Content -Path $logPath -Value "Column Headers:"
            for ($col = 1; $col -le $worksheet.UsedRange.Columns.Count; $col++) {
                $colHeader = $headerRow.Cells($col).Text
                Add-Content -Path $logPath -Value "Column ${col}: '$colHeader'"
                # More flexible column matching
                if ($colHeader -replace '\s','' -like "*AgencyID*") { $agencyIDCol = $col }
                if ($colHeader -replace '\s','' -like "*UserID*") { $userIDCol = $col }
            }
            Add-Content -Path $logPath -Value "`nFound Columns:"
            Add-Content -Path $logPath -Value "AgencyID Column: $agencyIDCol"
            Add-Content -Path $logPath -Value "UserID Column: $userIDCol"
            if ($agencyIDCol -eq -1 -or $userIDCol -eq -1) {
                Add-Content -Path $logPath -Value "ERROR: Required columns not found"
                Write-Host "Required columns not found in worksheet: $($worksheet.Name)" -ForegroundColor Yellow
                continue
            }
            # Process rows
            Add-Content -Path $logPath -Value "`nRow Processing:"
            for ($row = 2; $row -le $worksheet.UsedRange.Rows.Count; $row++) {
                $agencyID = $worksheet.Cells($row, $agencyIDCol).Text.Trim()
                $userID = $worksheet.Cells($row, $userIDCol).Text.Trim()
                Add-Content -Path $logPath -Value "Row $row - AgencyID: '$agencyID', UserID: '$userID'"
                # Check if UserID is numeric
                if ($userID -match '^\d+$') {
                    $newUserID = "$agencyID$userID"
                    $worksheet.Cells($row, $userIDCol).Value2 = $newUserID
                    Add-Content -Path $logPath -Value "Modified Row ${row}: New UserID = '$newUserID'"
                }
            }
        }
        # Create backup
        $backupDate = Get-Date -Format "yyyyMMdd_HHmmss"
        $backupFileName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
        $backupExtension = [System.IO.Path]::GetExtension($FilePath)
        $backupPath = Join-Path (Split-Path $FilePath -Parent) "$backupFileName`_$backupDate$backupExtension"
        $workbook.SaveAs($backupPath)
        $workbook.Save()
        Write-Host "Backup created: $backupPath" -ForegroundColor Green
        Write-Host "File processed successfully: $FilePath" -ForegroundColor Green
        Add-Content -Path $logPath -Value "`nBackup created: $backupPath"
    }
    catch {
        Write-Host "Error processing file: $_" -ForegroundColor Red
    }
    finally {
        if ($workbook) {
            $workbook.Close($false)
        }
        if ($excel) {
            $excel.Quit()
        }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
# Main script execution
try {
    # Verify file exists using full path
    if (-not (Test-Path $FullFilePath)) {
        Write-Host "File not found: $FullFilePath" -ForegroundColor Red
        exit 1
    }
    # Process the file using COM object method
    Process-ExcelWithCOMObject -FilePath $FullFilePath
}
catch {
    Write-Host "An unexpected error occurred: $_" -ForegroundColor Red
    exit 1
}