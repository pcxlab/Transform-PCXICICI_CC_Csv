# Define the path to the current directory
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Create an Excel application object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Convert all CSV files in the current directory to XLSX
$csvFiles = Get-ChildItem -Path $scriptRoot -Filter ICICI_CC*.csv
foreach ($csvFile in $csvFiles) {

    # Extract MOP from the file name
    $fileNameParts = (Get-Item $csvFile).BaseName -split '_'
    $MOP = "$($fileNameParts[0])_$($fileNameParts[1])_$($fileNameParts[2])"

    
    # Define the input and output file paths
    $inputCsv = $csvFile.FullName
    $convertedXlsxFilePath = "$($scriptRoot)\$($csvFile.BaseName)_ConvertedFromCsv.xlsx"
    
    # Open the CSV file
    $workbook = $excel.Workbooks.Open($inputCsv)
    
    # Save as Excel file
    $workbook.SaveAs($convertedXlsxFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault)
    
    # Close the workbook
    $workbook.Close($false)
    
    # Process the newly created XLSX file
    $newExcelFilePath = "$($scriptRoot)\$($csvFile.BaseName)_Transformed.xlsx"

    # Open the newly created Excel file
    try {
        $workbook = $excel.Workbooks.Open($convertedXlsxFilePath)
        Write-Output "Opened Excel file: $convertedXlsxFilePath"
    } catch {
        Write-Output "Failed to open Excel file: $convertedXlsxFilePath"
        continue
    }

    $worksheet = $workbook.Worksheets.Item(1)
    Write-Output "Accessed the first worksheet."

    # Define the headers to be found and their new names
    $requiredHeaders = @{
        "Transaction Date" = "Date"
        "Details" = "Narration"
        "Amount (INR)" = "Amount (INR)"  # Handle both Amt (Dr) and Amt (Cr) later
        "Reference Number" = "Chq./Ref.No."
    }

    # Define the new headers and their order
    $newHeaders = @("Date", "Narration", "Item", "Catogery", "Place", "Freq", "For", "MOP", "Amt (Dr)", "Chq./Ref.No.", "Value Dt", "Amt (Cr)")

    # Find the header rows and columns
    $headerPositions = @{ }
    foreach ($header in $requiredHeaders.Keys) {
        for ($row = 1; $row -le 15; $row++) {
            for ($col = 1; $col -le $worksheet.UsedRange.Columns.Count; $col++) {
                $cellValue = $worksheet.Cells.Item($row, $col).Value2
                Write-Output "Checking header: $cellValue at row $row, column $col"
                if ($cellValue -eq $header) {
                    $headerPositions[$header] = @{ Row = $row; Column = $col }
                    Write-Output "Found required header: $header at row $row, column $col"
                    break
                }
            }
            if ($headerPositions.ContainsKey($header)) { break }
        }
    }

    if ($headerPositions.Count -ne $requiredHeaders.Count) {
        Write-Output "Not all required headers were found in the Excel file."
        $workbook.Close($false)
        continue
    }

    # Create a new workbook for the filtered data
    $newWorkbook = $excel.Workbooks.Add()
    $newWorksheet = $newWorkbook.Worksheets.Item(1)

    # Write the new headers to the new worksheet
    $colIndex = 1
    foreach ($newHeader in $newHeaders) {
        $newWorksheet.Cells.Item(1, $colIndex) = $newHeader
        Write-Output "Filtered header '$newHeader' written to new worksheet."
        $colIndex++
    }

# Write the filtered data rows to the new worksheet
$rowIndex = 2
for ($i = $headerPositions["Transaction Date"].Row + 1; $i -le $worksheet.UsedRange.Rows.Count; $i++) {
    $colIndex = 1
    $amtDr = ""
    $amtCr = ""
    $valueDt = ""  # To store Dr or Cr

    foreach ($newHeader in $newHeaders) {
        $data = ""

        # Handle the "Amount (INR)" column explicitly
        if ($newHeader -eq "Amt (Dr)" -or $newHeader -eq "Amt (Cr)") {
            $amountValue = $worksheet.Cells.Item($i, $headerPositions["Amount (INR)"].Column).Value2
            $amountValue = $amountValue -as [string]

            # Check for " Dr." or " Cr." and handle accordingly
            if ($amountValue -match "\sDr\.$") {
                $valueDt = "Dr."  # Store " Dr." in Value Dt
                $amtDr = $amountValue -replace "\sDr\.$", ""  # Remove " Dr." suffix for Amt (Dr)
                $amtDr = $amtDr.Trim()  # Trim extra spaces
                if ($newHeader -eq "Amt (Dr)") {
                    $newWorksheet.Cells.Item($rowIndex, $colIndex) = $amtDr
                }
            } elseif ($amountValue -match "\sCr\.$") {
                $valueDt = "Cr."  # Store " Cr." in Value Dt
                $amtCr = $amountValue -replace "\sCr\.$", ""  # Remove " Cr." suffix for Amt (Cr)
                $amtCr = $amtCr.Trim()  # Trim extra spaces
                if ($newHeader -eq "Amt (Cr)") {
                    $newWorksheet.Cells.Item($rowIndex, $colIndex) = $amtCr
                }
            }
        } else {
            if ($requiredHeaders.ContainsValue($newHeader)) {
                $oldHeader = $requiredHeaders.Keys | Where-Object { $requiredHeaders[$_] -eq $newHeader }
                $data = $worksheet.Cells.Item($i, $headerPositions[$oldHeader].Column).Value2
            }

            # Handle MOP and Value Dt
            if ($newHeader -eq "MOP") {
                $data = $MOP
            } elseif ($newHeader -eq "Value Dt") {
                $data = $valueDt  # Copy Dr. or Cr. to Value Dt column
            }

            # Apply the text format for the "Date" column
            if ($newHeader -eq "Date") {
                #$newWorksheet.Cells.Item($rowIndex, $colIndex).NumberFormat = "@"  # Set as text format
                $newWorksheet.Cells.Item($rowIndex, $colIndex).NumberFormat = "dd-MM-yyyy"  # Set as text format

                $newWorksheet.Cells.Item($rowIndex, $colIndex).Value2 = $data  # No conversion
                #$newWorksheet.Cells.Item($rowIndex, $colIndex).Value2 = [string]$data  # Convert date to string
                #$newWorksheet.Cells.Item($rowIndex, $colIndex).Text = [string]$data  # Convert date to string

            } else {
                $newWorksheet.Cells.Item($rowIndex, $colIndex) = $data
            }
        }

        Write-Output "Processed row $i, column $colIndex ($newHeader): $data"
        $colIndex++
    }
    $rowIndex++
}


    # Save the new workbook as XLSX
    try {
        $newWorkbook.SaveAs($newExcelFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault)
        Write-Output "Filtered data with specified headers has been written to $newExcelFilePath"
    } catch {
        Write-Output "Failed to save the new Excel file: $newExcelFilePath"
    }

    # Close the workbooks
    $newWorkbook.Close()
    $workbook.Close($false)
}

# Quit Excel and release COM object
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel
