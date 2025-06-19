# Connect to Microsoft Graph    
Connect-MgGraph -Scopes "Mail.Send"

# Send emails to multiple recipients
$csvData = Import-Csv -Path "$HOME\Desktop\Temp.csv"
$userId = "badarin2050@hebron.edu.ps"
$total = $csvData.Count
$count = 0

foreach ($row in $csvData) {
    $count++
    $to = "$($row.ID)@hebron.edu.ps"
    $schoolName = $row.SchoolName
    $bodyHtml = @"
    <div dir='rtl' style='font-family:Segoe UI, Arial, sans-serif; font-size:16px;'>
    <p>حضرة السيد/ة مدير/ة مدرسة <strong>$schoolName</strong> المحترم </p>
    <p>الموضوع: المكلفون بالعمل في امتحان الثانوية العامة لعام 2025 م </p>
    <p> تحية طيبة وبعد، </p>
    <p>مع الشكر<br></p>
    </div>
"@
    $desktop = [Environment]::GetFolderPath("Desktop")
    $file = Join-Path -Path $desktop -ChildPath $row.Attachment1

    Write-Progress -Activity "Sending emails" -Status "Sending to $to..." -PercentComplete (($count / $total) * 100)

    if (Test-Path $file) {
        try {
            $bytes = [IO.File]::ReadAllBytes($file)
            $base64Content = [Convert]::ToBase64String($bytes)
            $attachmentName = [IO.Path]::GetFileName($file)
            $message = @{
                Message = @{
                    Subject = "المكلفون بالعمل في الثانوية العامة 2025 (المراقبة والتصحيح)"
                    Body = @{
                        ContentType = "HTML"
                        Content = $bodyHtml
                    }
                    ToRecipients = @(@{ EmailAddress = @{ Address = $to } })
                    Attachments = @(@{
                        "@odata.type" = "#microsoft.graph.fileAttachment"
                        Name = [IO.Path]::GetFileName($file)
                        ContentBytes = $base64Content
                    })
                }
                SaveToSentItems = $true
            }
            Send-MgUserMail -UserId $userId -BodyParameter $message
            Write-Host "📧 Sent to $to => 📎 $attachmentName => 🏫 $schoolName"

        }
        catch {
            Write-Warning "❌ Failed to send to $to. Error: $_"
        }
    }
    else {
        Write-Warning "⚠ File not found: $file"
    }
}
Write-Host "✅ All emails processed."
