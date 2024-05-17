# Load the Excel COM object
$excel = New-Object -ComObject Excel.Application

# Disable alerts and visibility
$excel.DisplayAlerts = $false
$excel.Visible = $false

$file_path = (Join-Path -Path (pwd) -ChildPath "PDF_Output.xlsx")

$workbook = $excel.Workbooks.Open($file_path)

# Set print options for the entire workbook
$excel.ActivePrinter = "Microsoft Print to PDF"
$excel.PrintCommunication = $true

foreach ($worksheet in $workbook.Worksheets) {
    $usedRange = $worksheet.UsedRange
    $Count += 1
    
    # Add borders to cells
    foreach ($cell in $usedRange.Cells) {
        if($Count -ne 2 -or $Count -ne 1){
        $cell.Borders.Weight = 2  # Thick borders
        $cell.WrapText = $true  # Enable text wrapping
        }
    }


    $usedRange.HorizontalAlignment = -4108  # Center alignment

    # Set print options to fit all columns on one page
    $worksheet.PageSetup.Zoom = $false
    $worksheet.PageSetup.FitToPagesWide = 1
    $worksheet.PageSetup.FitToPagesTall = $false  # Disable fitting to specific number of pages tall
    $worksheet.PageSetup.PaperSize = [Microsoft.Office.Interop.Excel.XlPaperSize]::xlPaperA3
}

# Print the entire workbook to PDF
$workbook.PrintOut()

# Restore print communication setting
$excel.PrintCommunication = $false

# Close the workbook and quit Excel
$workbook.Close($false)
$excel.Quit()

# Clean up COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel
