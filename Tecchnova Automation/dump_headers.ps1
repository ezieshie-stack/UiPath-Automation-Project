$excelPath = "C:\Users\this one\Documents\UiPath\Tecchnova Automation\SUPPLIER LIST TEMPLATE.xlsx"
$sheetName = "Supplier List"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)
$sheet = $workbook.Sheets.Item($sheetName)

Write-Host "Dumping Row 3 (Headers):"
for ($j = 1; $j -le 10; $j++) {
    $val = $sheet.Cells.Item(3, $j).Text
    Write-Host "Col $j : '$val'"
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
