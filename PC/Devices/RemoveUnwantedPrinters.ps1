# List of printer names to remove
$printersToRemove = @(
    "OneNote (Desktop)",
    "Snagit 2023",
    "Snagit 2024",
    "RustDesk Printer"
)
foreach ($targetPrinter in $printersToRemove) {
    $matches = Get-Printer | Where-Object { $_.Name -like "*$targetPrinter*" }

    if ($matches) {
        foreach ($printer in $matches) {
            Write-Host "Removing printer: $($printer.Name)"
            Remove-Printer -Name $printer.Name -ErrorAction SilentlyContinue
        }
        Write-Host "✅ Removed printer(s) matching: '$targetPrinter'"
    } else {
        Write-Host "ℹ️ No printer found matching: '$targetPrinter'"
    }
}