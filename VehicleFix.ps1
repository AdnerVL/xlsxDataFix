param(
    [Parameter(Position=0)]
    [string]$FilePath
)

function List-ExcelFiles {
    $scriptDir = (Get-Location).Path
    $excelFiles = Get-ChildItem -Path $scriptDir -Filter *.xlsx
    if ($excelFiles.Count -eq 0) {
        Write-Host "No Excel files found in the script directory." -ForegroundColor Red
        exit
    }
    Write-Host "Available Excel files in the script directory:" -ForegroundColor Green
    for ($i = 0; $i -lt $excelFiles.Count; $i++) {
        Write-Host "[$i] $($excelFiles[$i].Name)"
    }
    return $excelFiles
}

function Process-ExcelWithCOMObject {
    param([string]$FilePath)
    
    $excel = $null
    $workbook = $null
    $worksheet = $null
    
    try {
        # Create backup before any modifications
        $backupDate = Get-Date -Format "yyyyMMdd_HHmmss"
        $backupFileName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
        $backupExtension = [System.IO.Path]::GetExtension($FilePath)
        $backupPath = Join-Path (Split-Path $FilePath -Parent) "$backupFileName`_$backupDate$backupExtension"
        Copy-Item -Path $FilePath -Destination $backupPath -Force
        Write-Host "Backup created: $backupPath" -ForegroundColor Green

        # Initialize Excel
        Write-Host "Initializing Excel..." -ForegroundColor Yellow
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        # Open workbook with retry mechanism
        $maxRetries = 3
        $retryCount = 0
        while ($true) {
            try {
                $workbook = $excel.Workbooks.Open($FilePath)
                break
            } catch {
                $retryCount++
                if ($retryCount -ge $maxRetries) {
                    throw "Failed to open workbook after $maxRetries attempts: ${_}"
                }
                Start-Sleep -Seconds 2
            }
        }

        # Process each worksheet
        foreach ($worksheet in $workbook.Worksheets) {
            Write-Host "Processing worksheet: $($worksheet.Name)" -ForegroundColor Yellow
            
            # Find columns
            $headerRow = $worksheet.UsedRange.Rows(1)
            $agencyIDCol = -1
            $VehicleIDCol = -1
            
            for ($col = 1; $col -le $headerRow.Columns.Count; $col++) {
                $header = $headerRow.Cells(1, $col).Text.Trim()
                if ($header -replace '\s','' -like "*AgencyID*") { $agencyIDCol = $col }
                if ($header -replace '\s','' -like "*VehicleID*") { $VehicleIDCol = $col }
            }
            
            if ($agencyIDCol -eq -1 -or $VehicleIDCol -eq -1) {
                Write-Host "Skipping worksheet $($worksheet.Name): Required columns not found" -ForegroundColor Yellow
                continue
            }

            # Process rows
            $rowCount = $worksheet.UsedRange.Rows.Count
            for ($row = 2; $row -le $rowCount; $row++) {
                try {
                    $agencyID = $worksheet.Cells($row, $agencyIDCol).Text.Trim()
                    $VehicleID = $worksheet.Cells($row, $VehicleIDCol).Text.Trim()
                    
                    if ([string]::IsNullOrWhiteSpace($VehicleID)) { continue }
                    
                    if ($VehicleID -match '[a-zA-Z]') {
                        # Replace letters with AgencyID while keeping numbers after
                        $newVehicleID = $agencyID + ($VehicleID -replace '[a-zA-Z]', '')
                    } else {
                        # Append AgencyID before numeric VehicleID
                        $newVehicleID = "$agencyID$VehicleID"
                    }
                    
                    $worksheet.Cells($row, $VehicleIDCol).Value2 = $newVehicleID
                } catch {
                    Write-Host "Error processing row ${row}: ${_}" -ForegroundColor Red
                }
            }
        }

        $workbook.Save()
        Write-Host "Changes saved successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "Excel processing error: ${_}" -ForegroundColor Red
        throw
    }
    finally {
        if ($worksheet) {
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) } catch {}
        }
        if ($workbook) {
            try {
                $workbook.Close($false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
            } catch {}
        }
        if ($excel) {
            try {
                $excel.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
            } catch {}
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# Main execution
try {
    if (-not $FilePath) {
        $excelFiles = List-ExcelFiles
        $index = Read-Host "Please enter the index of the Excel file to process"
        if ($index -match '^\d+$' -and $index -ge 0 -and $index -lt $excelFiles.Count) {
            $FilePath = $excelFiles[$index].FullName
        } else {
            Write-Host "Invalid selection." -ForegroundColor Red
            exit 1
        }
    }

    if (-not (Test-Path $FilePath)) {
        Write-Host "File not found: $FilePath" -ForegroundColor Red
        exit 1
    }

    Process-ExcelWithCOMObject -FilePath $FilePath
}
catch {
    Write-Host "Fatal error: ${_}" -ForegroundColor Red
    exit 1
}