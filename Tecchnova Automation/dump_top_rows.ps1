$excelPath = "C:\Users\this one\Documents\UiPath\Tecchnova Automation\SUPPLIER LIST TEMPLATE.xlsx"
$sheetName = "Supplier List"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)
$sheet = $workbook.Sheets.Item($sheetName)

Write-Host "Dumping Row 1 and 2:"
for ($i = 1; $i -le 2; $i++) {
    Write-Host "Row ${i}:"
    for ($j = 1; $j -le 10; $j++) {
        $val = $sheet.Cells.Item($i, $j).Text
        Write-Host "  Col $j : '$val'"
    }
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
