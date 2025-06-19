
# Set Excel and Pdf files path  
$excelFilePath  = "$Home\Desktop\aa.xlsx"  # Replace with your CSV file path
$pdfFilePath = "$Home\Desktop\aa.pdf"  # Output PDF file path

# Create an Excel application object
$excelApp = New-Object -ComObject Excel.Application
$excelApp.Visible = $false
$excelApp.DisplayAlerts = $false

# Open the Excel file
$workbook = $excelApp.Workbooks.Open($excelFilePath)

# Set print options
$worksheet = $workbook.Worksheets.Item(1)
$worksheet = $workbook.Worksheets.Item(1)
$usedRange = $worksheet.UsedRange

# Set text size for all cells
$usedRange.Font.Size = 14

# Center text horizontally and vertically in all cells
$usedRange.HorizontalAlignment = -4108  # -4108 = Center alignment
$usedRange.VerticalAlignment = -4108    # -4108 = Middle alignment

# Bold headers (first row) 
$headers = $worksheet.Rows.Item(1)
$headers.Font.Bold = $true

# Add border lines to all cells
$usedRange.Borders.LineStyle = 1  # 1 = Continuous line
$usedRange.Borders.Weight = 2     # 2 = Thin border

# Set row height to 40 for all rows
$usedRange.EntireRow.RowHeight = 40

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

# Set Header and HeaderMargin
$worksheet.PageSetup.CenterHeader ="&30 &B إستمارة الطلبة ذوي الإعاقة للعام الدراسي 2024-2025"  # "Page X of Y" format
$worksheet.PageSetup.HeaderMargin = 10
$worksheet.PageSetup.TopMargin = 20

# Set Footer and FooterMargin 
$worksheet.PageSetup.CenterFooter = "&12 &P" 
$worksheet.PageSetup.FooterMargin  = 20

# Set the top margin (distance from the top of the page to the table)
$worksheet.PageSetup.TopMargin = 75

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
