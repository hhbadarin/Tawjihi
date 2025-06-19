# 1. ----------- Report ----- 
$data = Import-Csv -Path "$Home\Desktop\Quran.csv"
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
$combinedData | Export-Csv  "$Home\Desktop\AggregatedQuranList.csv" -NoTypeInformation -Encoding UTF8
$combinedData | Format-Table -AutoSize

# 2. ------------ Removing duplicates ----- 
$data = Import-Csv -Path "$Home\Desktop\Quran.csv"
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
$combinedData | Sort-Object SchoolID, LevelClass Export-Csv  "$Home\Desktop\UniqueQuranList.csv" -NoTypeInformation -Encoding UTF8
$combinedData | Format-Table -AutoSize