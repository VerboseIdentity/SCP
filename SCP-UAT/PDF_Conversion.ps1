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
    $foundServerName = $false  # Flag to indicate when "Server_Name" is found
    
    # Iterate over each row in the used range
    foreach ($row in $usedRange.Rows) {
        foreach ($cell in $row.Cells) {
            if ($worksheet.Name -eq "Cover Page" -or $worksheet.Name -eq "Knowledge") {
                continue
            }
            elseif ($cell.Value2 -eq "Server_Name") {
                $foundServerName = $true  # Set flag to true when "Server_Name" is found
            }
            # Apply borders and wrapping only if "Server_Name" has been found and the cell value is not "Analysis"
            elseif ($foundServerName -and $cell.Value2 -ne "Analysis") {
                $cell.Borders.Weight = 2  # Thick borders
                $cell.WrapText = $true  # Enable text wrapping
            }
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

Write-Output "Completed'."
