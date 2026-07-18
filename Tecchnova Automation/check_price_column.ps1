$excelPath = "C:\Users\this one\Documents\UiPath\Tecchnova Automation\SUPPLIER LIST TEMPLATE.xlsx"
$sheetName = "Supplier List"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)
$sheet = $workbook.Sheets.Item($sheetName)

$usedRange = $sheet.UsedRange
$rowCount = $usedRange.Rows.Count

Write-Host "Checking Column 7 (Index 6) for non-numeric values..."

# Start from row 4 (data starts after header in row 3)
for ($i = 4; $i -le $rowCount; $i++) {
    $cellValue = $sheet.Cells.Item($i, 7).Text
    if (-not [string]::IsNullOrWhiteSpace($cellValue)) {
        # Check if it's a valid number
        $isNumber = $cellValue -as [double]
        if ($null -eq $isNumber) {
            Write-Host "Row ${i}: Found non-numeric value in Column 7: '$cellValue'"
        }
    }
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
