Get-Service |
    Export-Excel "c:\temp\test2.xlsx" `
        -Show `
        -IncludePivotTable `
        -IncludePivotChart `
        -PivotRows status `
        -PivotData @{status='count'}

#General info about Excel file
Get-ExcelSheetInfo -Path "$Home\Desktop\HEV.xlsx"

#Import the specific sheet to a file
Import-Excel -Path "$Home\Desktop\HEV.xlsx" -WorksheetName "Sheet1" | Format-Table

#Import all sheets to a file
$AllSheets = Get-ExcelSheetInfo -Path "$Home\Desktop\HEV.xlsx"
foreach ($Sheet in $AllSheets) {
    Import-Excel -Path "$Home\Desktop\HEV.xlsx" -WorksheetName $Sheet.Name | Format-Table
}

#Import specific rows and columns from Excel file
Import-Excel -Path "$Home\Desktop\HEV.xlsx" -StartRow 1 -EndRow 5 -StartColumn 2 -EndColumn 3

#remove the headers
Import-Excel -Path  "$Home\Desktop\HEV.xlsx" -StartRow 2 -EndRow 5 -StartColumn 2 -EndColumn 3 -NoHeader

#Import specific columns
Import-Excel -Path  "$Home\Desktop\HEV.xlsx"  -WorksheetName "Sheet1" -ImportColumns 1 


#Import specific Excel worksheet to Out-GridView
Import-Excel -Path  "$Home\Desktop\HEV.xlsx" | Sort-Object Name | Out-GridView -Title "Employees HR"