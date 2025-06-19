
# ----------------- to pdf
$csvFilePath = "$Home\Desktop\staffReport_20250522.csv"  # Replace with your CSV file path
$pdfFilePath = "$Home\Desktop\staffReport_20250522.pdf"  # Output PDF file path
# Create Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
# Open the CSV file in Excel
$workbook = $excel.Workbooks.Open($csvFilePath)
$worksheet = $workbook.Worksheets.Item(1)
$usedRange = $worksheet.UsedRange

# Set text size to 24 for all cells
$usedRange.Font.Size = 18

# Center text horizontally and vertically in all cells
$usedRange.HorizontalAlignment = -4108  # -4108 = Center alignment
$usedRange.VerticalAlignment = -4108    # -4108 = Middle alignment

# Bold headers (first row) and add shading
$headers = $worksheet.Rows.Item(1)
$headers.Font.Bold = $true
$headers.Interior.ColorIndex = 15  # Shading color

# Add border lines to all cells
$usedRange.Borders.LineStyle = 1  # 1 = Continuous line
$usedRange.Borders.Weight = 2     # 2 = Thin border

# Set row height to 40 for all rows
$usedRange.EntireRow.RowHeight = 25

# Set the width of the "No." column to 20
$worksheet.Columns.Item(1).ColumnWidth = 20

# Auto-size other columns for better readability
$usedRange.EntireColumn.AutoFit()

# Adjust page layout to fit all columns on one page (portrait orientation)
$worksheet.PageSetup.Orientation = 2  # 1 = Portrait orientation, 2 = Landscape orientation
$worksheet.PageSetup.Zoom = $false    # Disable zoom scaling
$worksheet.PageSetup.FitToPagesWide = 1  # Fit all columns to 1 page wide
$worksheet.PageSetup.FitToPagesTall = $false  # Do not limit the number of pages tall

# Repeat the header row on every page
$worksheet.PageSetup.PrintTitleRows = "1:1"  # Repeat the first row (header) on every page

# Center the table horizontally on the page
$worksheet.PageSetup.CenterHorizontally = $true

# Add footer and header with page number
$worksheet.PageSetup.CenterHeader ="&20 &B المدارس على الايسكول المسميات `n والأرقام الوطنية"  # "Page X of Y" format
$worksheet.PageSetup.CenterFooter = "&17 &P"  # "Page X of Y" format

# Set the top margin (distance from the top of the page to the table)
$worksheet.PageSetup.TopMargin = 80

# Export the worksheet to PDF
$worksheet.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF, $pdfFilePath)

# Clean up
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

# Display a message indicating the export is complete
Write-Host "PDF has been saved to: $pdfFilePath"
