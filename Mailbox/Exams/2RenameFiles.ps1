
# Rename files in a folder based on a CSV reference file
$SourceFolder = "~/Desktop/NewAttachments" 
if (-not (Test-Path $SourceFolder)) {
    New-Item -ItemType Directory -Path $SourceFolder
}
$oldFolder = "~\Desktop\OldAttachments"
$newFolder = "~\Desktop\NewAttachments"
$csvData = Import-Csv -Path "$Home\Documents\GitHub\Microsoft-Graph\Mailbox\Exams\Rename_Reference.csv"
foreach ($row in $csvData) {
    $currentBaseName = $row.CurrentName
    $newBaseName = $row.NewName

    # Find the file in Old folder matching the base name (any extension)
    $currentFile = Get-ChildItem -Path $oldFolder -Filter "$currentBaseName.*" | Select-Object -First 1

    if ($null -ne $currentFile) {
        $extension = $currentFile.Extension
        $newFileName = $newBaseName + $extension
        $newFilePath = Join-Path -Path $newFolder -ChildPath $newFileName

        try {
            # Copy and rename
            Copy-Item -Path $currentFile.FullName -Destination $newFilePath -ErrorAction Stop
            Write-Host "Copied and renamed '$($currentFile.Name)' to '$newFileName'" -ForegroundColor Yellow
        }
        catch {
            Write-Host "Error copying '$($currentFile.Name)': $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "File starting with '$currentBaseName' not found in the folder." -ForegroundColor Red
    }
}
Write-Host "File copying and renaming completed." -ForegroundColor Green
