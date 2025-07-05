# Connect with only Mail.Send (User.Read removed)
Connect-MgGraph -Scopes "Mail.Send"

# === Step 1: Prompt for Directorate Names ===
$directorateEnglish = Read-Host "Enter the directorate name in English (for folder name)"
$directorateArabic = Read-Host "Enter the directorate name in Arabic (for email)"

# === Step 2: Create folder and zip path ===
$today = Get-Date -Format "dd-MM-yyyy"
$folderName = "$directorateEnglish-$today"
$desktop = [Environment]::GetFolderPath("Desktop")
$folderPath = Join-Path $desktop $folderName
$zipPath = "$folderPath.zip"

# === Step 3: Zip the Folder ===
if (Test-Path $folderPath) {
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    Compress-Archive -Path $folderPath -DestinationPath $zipPath
    Write-Host "✅ Folder zipped: $zipPath"
} else {
    Write-Warning "⚠️ Folder not found: $folderPath"
    exit
}

# === Step 4: Prepare the Email ===
$toEmail = "badarin2050@hebron.edu.ps"
$ccEmail = "badarin2050@gmail.com"
$subject = "الملفات الخاصة بمديرية $directorateArabic ليوم $today"
$bodyHtml = @"
<div style='direction:rtl; font-family:Segoe UI, sans-serif; font-size:16px;'>
<p>مرفق الملفات الخاصة بمديرية $directorateArabic ليوم $today.</p>
</div>
"@

# === Step 5: Create Email Body Structure ===
$mailBody = @{
    Subject = $subject
    Body = @{
        ContentType = "HTML"
        Content = $bodyHtml
    }
    ToRecipients = @(
        @{ EmailAddress = @{ Address = $toEmail } }
    )
    CcRecipients = @(
        @{ EmailAddress = @{ Address = $ccEmail } }
    )
    Attachments = @(
        @{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            Name = "$folderName.zip"
            ContentBytes = [Convert]::ToBase64String([IO.File]::ReadAllBytes($zipPath))
        }
    )
}

# === Step 6: Send Email with Error Handling ===
try {
    Send-MgUserMail -UserId "tawjihi2025@hebron.edu.ps" -Message $mailBody -SaveToSentItems
    Write-Host "📤 Email sent successfully to $toEmail with CC to $ccEmail."
} catch {
    Write-Warning "⚠️ Failed to send email: $_"
    exit 1
}
