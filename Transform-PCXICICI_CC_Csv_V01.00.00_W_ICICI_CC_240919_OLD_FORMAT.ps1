Clear-Host

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

    # Open the CSV file correctly using OpenText
    $excel.Workbooks.OpenText($inputCsv, 2, 1, 1, [Microsoft.Office.Interop.Excel.XlTextParsingType]::xlDelimited, $false, $false)

    # Save as Excel file
    $workbook = $excel.ActiveWorkbook
    $workbook.SaveAs($convertedXlsxFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault)
    
    # Close the workbook
    $workbook.Close($false)
}

#########################################################
# Now process all .xlsx files in the folder (including converted ones)
$excelFiles = Get-ChildItem -Path $scriptRoot -Filter ICICI_CC*.xlsx

foreach ($excelFile in $excelFiles) {
    $excelFilePath = $excelFile.FullName

    # Extract MOP from the file name
    $fileNameParts = (Get-Item $excelFilePath).BaseName -split '_'
    $MOP = "$($fileNameParts[0])_$($fileNameParts[1])_$($fileNameParts[2])"

    # Open the existing Excel file
    try {
        $workbook = $excel.Workbooks.Open($excelFilePath)
        Write-Host "Opened Excel file: $excelFilePath"
    } catch {
        Write-Host "Failed to open Excel file: $excelFilePath"
        continue
    }

    $worksheet = $workbook.Worksheets.Item(1)
    Write-Host "Accessed the first worksheet."

    # Define the headers to be found and their new names
    $requiredHeaders = @{
        "Date" = "Date"
        "Transaction Details" = "Narration"
        "Amount(in Rs)" = "Amt (Dr)"
        "Sr.No." = "Chq./Ref.No."
        "BillingAmountSign" = "Value Dt"
    }

    # Define the new headers and their order
    $newHeaders = @("Date", "Narration", "Item", "Catogery", "Place", "Freq", "For", "MOP", "Amt (Dr)", "Chq./Ref.No.", "Value Dt", "Amt (Cr)")

    # Find the header rows and columns
    $headerPositions = @{ }
    foreach ($header in $requiredHeaders.Keys) {
        for ($row = 1; $row -le 50; $row++) {  # Dynamically search within the first 50 rows
            for ($col = 1; $col -le $worksheet.UsedRange.Columns.Count; $col++) {
                $cellValue = $worksheet.Cells.Item($row, $col).Value2
                if ($cellValue -eq $header) {
                    $headerPositions[$header] = @{ Row = $row; Column = $col }
                    Write-Host "Found required header: $header at row $row, column $col"
                    break
                }
            }
            if ($headerPositions.ContainsKey($header)) { break }
        }
    }

    # Check if all required headers were found
    if ($headerPositions.Count -ne $requiredHeaders.Count) {
        $missingHeaders = $requiredHeaders.Keys | Where-Object { -not $headerPositions.ContainsKey($_) }
        Write-Host "Not all required headers were found: $($missingHeaders -join ', ')"
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
        Write-Host "Filtered header '$newHeader' written to new worksheet."
        $colIndex++
    }

    # Write the filtered data rows to the new worksheet
    $rowIndex = 2
    for ($i = $headerPositions["Date"].Row + 1; $i -le $worksheet.UsedRange.Rows.Count; $i++) {
        $colIndex = 1

        foreach ($newHeader in $newHeaders) {
            $data = ""

            switch ($newHeader) {
                "Date" {
                    $dateValue = $worksheet.Cells.Item($i, $headerPositions["Date"].Column).Value2
                    #$data = Get-FormattedDate $dateValue
                    $data = $dateValue
                }
                "Narration" {
                    $data = $worksheet.Cells.Item($i, $headerPositions["Transaction Details"].Column).Value2
                }
                "Amt (Dr)" {
                    $debitCredit = $worksheet.Cells.Item($i, $headerPositions["BillingAmountSign"].Column).Value2
                    if ($debitCredit -ne "Cr") {
                        $data = $worksheet.Cells.Item($i, $headerPositions["Amount(in Rs)"].Column).Value2
                    }
                }
                "Amt (Cr)" {
                    $debitCredit = $worksheet.Cells.Item($i, $headerPositions["BillingAmountSign"].Column).Value2
                    if ($debitCredit -eq "Cr") {
                        $data = $worksheet.Cells.Item($i, $headerPositions["Amount(in Rs)"].Column).Value2
                    }
                }
                "Chq./Ref.No." {
                    $data = $worksheet.Cells.Item($i, $headerPositions["Sr.No."].Column).Value2
                }
                "Value Dt" {
                    $data = $worksheet.Cells.Item($i, $headerPositions["BillingAmountSign"].Column).Value2
                }
                "MOP" {
                    $data = $MOP
                }
            }

            $newWorksheet.Cells.Item($rowIndex, $colIndex) = $data
            Write-Host "Processed row $i, column $colIndex ($newHeader): $data"
            $colIndex++
        }
        $rowIndex++
    }

    # Define transformed file path
    $newExcelFilePath = [System.IO.Path]::Combine($scriptRoot, "$($excelFile.BaseName)_Transformed.xlsx")

    # Save the new workbook as XLSX
    try {
        $newWorkbook.SaveAs($newExcelFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault)
        Write-Host "Filtered data written to $newExcelFilePath"
    } catch {
        Write-Host "Failed to save the new Excel file: $newExcelFilePath"
    }

    # Close the workbooks
    $newWorkbook.Close()
    $workbook.Close($false)
}

# Quit Excel
$excel.Quit()

# Release COM objects for transformation
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
