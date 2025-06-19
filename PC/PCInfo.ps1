# Get PC basic information and export to CSV
$pcOwner = Read-Host "Enter the PC owner's name"
$csvPath = "$env:USERPROFILE\Desktop\$($pcOwner)_PCInfo.csv"
$Timestamp                = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$pcName = $env:COMPUTERNAME
$serialNumber = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber
$os = Get-CimInstance -ClassName Win32_OperatingSystem
$osName = $os.Caption
$productNumber = (Get-CimInstance -ClassName Win32_ComputerSystemProduct).IdentifyingNumber
$compSystem = Get-CimInstance -ClassName Win32_ComputerSystem
try {
    $displayVersion = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "DisplayVersion"
} catch {
    $displayVersion = "Unknown"
}
$totalDiskBytes = (Get-CimInstance -ClassName Win32_DiskDrive | Measure-Object -Property Size -Sum).Sum
$totalDiskGB = [math]::Round($totalDiskBytes / 1GB, 2)
$totalMemoryGB = [math]::Round($compSystem.TotalPhysicalMemory / 1GB, 2)
$pcSummary = [PSCustomObject]@{
    Owner            = $pcOwner
    Timestamp       = $Timestamp
    PCName            = $pcName
    SerialNumber      = $serialNumber
    ProductNumber     = $productNumber
    OS                = $osName
    Version    = $displayVersion
    TotalMemoryGB     = $totalMemoryGB
    TotalDiskSizeGB   = $totalDiskGB
}
$pcSummary | Format-Table -AutoSize
$pcSummary | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
