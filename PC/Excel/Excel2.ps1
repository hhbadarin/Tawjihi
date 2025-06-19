#< Link: https://learn-powershell.net/2012/12/20/powershell-and-excel-adding-some-formatting-to-your-report/ >#

#General info about Excel file
Get-ExcelSheetInfo -Path "$Home\Desktop\HEV.xlsx"

#Import the specific sheet to a file
Import-Excel -Path "$Home\Desktop\HEV.xlsx" -WorksheetName "Sheet1" | Format-Table

#Create excel COM object
$excel = New-Object -ComObject excel.application
$excel.Visible = $True
$workbook = $excel.Workbooks.Add()
$serverInfoSheet = $workbook.Worksheets.Item(1)
$serverInfoSheet.Name = 'DiskInformation'
$serverInfoSheet.Activate() | Out-Null

#Create a Title for the first worksheet and adjust the font
$row = 1
$Column = 1
$serverInfoSheet.Cells.Item($row,$column)= 'Disk Space Information'

#Format Title and give a more “title” look
$serverInfoSheet.Cells.Item($row,$column).Font.Size = 18
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$serverInfoSheet.Cells.Item($row,$column).Font.Name = "Cambria"
$serverInfoSheet.Cells.Item($row,$column).Font.ThemeFont = 1
$serverInfoSheet.Cells.Item($row,$column).Font.ThemeColor = 4
$serverInfoSheet.Cells.Item($row,$column).Font.ColorIndex = 55
$serverInfoSheet.Cells.Item($row,$column).Font.Color = 8210719

$range = $serverInfoSheet.Range("a1","h2")
$range.Style = 'Title'
$range = $serverInfoSheet.Range("a1","g2")
$range.Merge() | Out-Null
$range.VerticalAlignment = -4160

# Add headers 
#Increment row for next set of data
$row++;$row++

#Save the initial row so it can be used later to create a border
$initalRow = $row
#Create a header for Disk Space Report; set each cell to Bold and add a background color
$serverInfoSheet.Cells.Item($row,$column)= 'Computername'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= 'DeviceID'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= 'VolumeName'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= 'TotalSize(GB)'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= 'UsedSpace(GB)'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= 'FreeSpace(GB)'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= 'PercentFree'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True


# Add data
$row++
$Column = 1
#Get the drives
$diskDrives = Get-CimInstance -ClassName Cim_LogicalDisk
#Process each disk in the collection and write to spreadsheet
ForEach ($disk in $diskDrives) {
    $serverInfoSheet.Cells.Item($row,$column)= $disk.__Server
    $Column++
    $serverInfoSheet.Cells.Item($row,$column)= $disk.DeviceID
    $Column++
    $serverInfoSheet.Cells.Item($row,$column)= $disk.VolumeName
    $Column++
    $serverInfoSheet.Cells.Item($row,$column)= [math]::Round(($disk.Size /1GB),2)
    $Column++
    $serverInfoSheet.Cells.Item($row,$column)= [math]::Round((($disk.Size - $disk.FreeSpace)/1GB),2)
    $Column++
    $serverInfoSheet.Cells.Item($row,$column)= [math]::Round(($disk.FreeSpace / 1GB),2)
    $Column++
    $serverInfoSheet.Cells.Item($row,$column)= ("{0:P}" -f ($disk.FreeSpace / $disk.Size))
    
    #Check to see if space is near empty and use appropriate background colors
    $range = $serverInfoSheet.Range(("A{0}"  -f $row),("G{0}"  -f $row))
    $range.Select() | Out-Null
    
    #Determine if disk needs to be flagged for warning or critical alert
    If ($disk.FreeSpace -lt 65GB -AND ($disk.FreeSpace / $disk.Size) -lt 75) {
        #Critical threshold 
        $range.Interior.ColorIndex = 3
    } ElseIf ($disk.FreeSpace -lt 80GB -AND ($disk.FreeSpace / $disk.Size) -lt 80) {
        #Warning threshold 
        $range.Interior.ColorIndex = 6
    }
    
    #Increment to next row and reset Column to 1
    $Column = 1
    $row++
}

#add some borders to this to give it a cleaner look.
$row--
$dataRange = $serverInfoSheet.Range(("A{0}"  -f $initalRow),("G{0}"  -f $row))
7..12 | ForEach {
    $dataRange.Borders.Item($_).LineStyle = 1
    $dataRange.Borders.Item($_).Weight = 2
}

# deciding what the LineStyle and Weight are for the borders.
[Enum]::getvalues([Microsoft.Office.Interop.Excel.XLLineStyle]) | 
Select-Object @{n="Name";e={"$_"}},value__

#Auto fit everything so it looks better
$usedRange = $serverInfoSheet.UsedRange	
$usedRange.EntireColumn.AutoFit() | Out-Null

#Save the file
$workbook.SaveAs("C:\temp\DiskSpace.xlsx")
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel) | Out-Null
