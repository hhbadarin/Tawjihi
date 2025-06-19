# 1. ----------- Report ----- 
$data = Import-Csv -Path "$Home\Desktop\Khatabi.csv"
$combinedData = @()
foreach ($row in $data) {
    for ($level = 1; $level -le 6; $level++) {
        $levelClass = "Level${level}Class"
        $levelStudentName = "Level${level}StudentName"
        if ($row.$levelClass -and $row.$levelStudentName) {
            $combinedData += [PSCustomObject]@{
                SchoolName       = $row.SchoolName
                SchoolID         = $row.SchoolID
                LevelClass       = $row.$levelClass
                LevelStudentName = $row.$levelStudentName
            }
        }
    }
}
$combinedData | Export-Csv  "$Home\Desktop\$("مسابقة الخطابة قائمة تجمعية").csv" -NoTypeInformation -Encoding UTF8
$combinedData | Format-Table -AutoSize

# 2. ------------ Removing duplicates ----- 
$data = Import-Csv -Path  "$Home\Desktop\Khatabi.csv"
$uniqueCombinations = @{}
$combinedData = @()
foreach ($row in $data) {
    for ($level = 1; $level -le 6; $level++) {
        $levelClass = "Level${level}Class"
        $levelStudentName = "Level${level}StudentName"
        if ($row.$levelClass -and $row.$levelStudentName) {
            $uniqueKey = "$($row.SchoolName)-$($row.SchoolID)-$($row.$levelClass)-$($row.$levelStudentName)"
            if (-not $uniqueCombinations.ContainsKey($uniqueKey)) {
                $uniqueCombinations[$uniqueKey] = $true
                $combinedData += [PSCustomObject]@{
                    SchoolName       = $row.SchoolName
                    SchoolID         = $row.SchoolID
                    LevelClass       = $row.$levelClass
                    LevelStudentName = $row.$levelStudentName
                }
            }
        }
    }
}
$combinedData | Export-Csv "$Home\Desktop\$("مسابقة الخطابة قائمة تجميعية بدون تكرار").csv"  -NoTypeInformation -Encoding UTF8
$combinedData | Format-Table -AutoSize


#3. ----- Values in two columns
$csvPath = "$Home\Desktop\ValuesinTwoCols.csv"  # Replace with the path to your CSV file
$csv = Import-Csv -Path $csvPath

# Define the columns to compare
$column1 = "Column1"  # Replace with the name of the first column
$column2 = "Column2"  # Replace with the name of the second column

# Create a hashtable to store values from the first column
$column1Values = @{}
foreach ($row in $csv) {
    $column1Values[$row.$column1] = $true
}

# Find values present in both columns
$commonValues = @()
foreach ($row in $csv) {
    $valueInColumn2 = $row.$column2
    if ($column1Values.ContainsKey($valueInColumn2)) {
        $commonValues += $valueInColumn2
    }
}

# Remove duplicates and sort the common values
$uniqueCommonValues = $commonValues | Sort-Object -Unique

# Export the common values to a new CSV file
$outputCsvPath = "$Home\Desktop\CommonValues.csv"  # Path to the output CSV file
$uniqueCommonValues | ForEach-Object {
    [PSCustomObject]@{ 
        CommonValue = $_
    }
} | Export-Csv -Path $outputCsvPath -NoTypeInformation -Encoding UTF8

Write-Output "Common values exported to $outputCsvPath."