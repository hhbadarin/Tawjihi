# Connect to Microsoft Graph with required Mail.Send scope
Connect-MgGraph -Scopes "Mail.Send"

# CSV file with ID and SchoolName columns only
$csvPath = "$HOME\Documents\GitHub\Microsoft-Graph\Mailbox\Exams\Temp.csv"

if (!(Test-Path $csvPath)) {
    Write-Host "❌ File not found: $csvPath" -ForegroundColor Red
    return
}

$csvData = Import-Csv -Path $csvPath
$total = $csvData.Count
$count = 0
$userId = "exams@hebron.edu.ps"
$desktop = [Environment]::GetFolderPath("Desktop")

# Common values for attachment location and naming
$attachmentFolder = Join-Path $desktop "NewAttachments"
$attachmentPrefix = "المكلفون بالعمل في الثانوية العامة الدفعة السادسة -"

foreach ($row in $csvData) {
    $count++
    $id = $row.ID
    $schoolName = $row.SchoolName
    $to = "$id@hebron.edu.ps"

    # Build attachment path from ID
    $attachmentFileName = "$attachmentPrefix$id.pdf"
    $file = Join-Path -Path $attachmentFolder -ChildPath $attachmentFileName
    $attachmentName = $attachmentFileName

    # Email body
    $bodyHtml = @"
<div dir='rtl' style='font-family:Segoe UI, Arial, sans-serif; font-size:16px;'>
<p>حضرة السيد/ة مدير مدرسة <strong>$schoolName</strong> المحترم </p>
<p>الموضوع: ملحق المكلفون بالعمل في الثانوية العامة 2025 (المراقبة) / الدفعة السادسة </p>
<p> تحية طيبة وبعد، </p>

<p>نرسل لحضرتكم اسماء المكلفين/ الدفعة السادسة الذين تم اختيارهم للعمل في امتحان الثانوية العامة لعام 2025م، علما ان العمل في امتحان الثانوية العامة إلزامي للذين تم اختيارهم.</p>

<p>1. ملاحظة: نرجو التأكيد على المعلم إحضار التالي في أول يوم امتحان:</p>
<p>a. احضار الهوية الشخصية</p>
<p>b. قسيمة الراتب لاخر شهر للمعلم المثبت أو كتاب يظهر فيه IBAN رقم الحساب البنكي + كتاب التنسيب أو التعيين او مباشرة العمل من المدرسة.</p>
<p>2. ملاحظة: نرجو من جميع المكلفين التأكد من اسمه ورقم هويته ورقم جواله وان وجد اي اختلاف ابلاغ قسم الامتحانات بذلك.</p>
<p>3. ملاحظة: من يرغب في الاعتذار عن العمل في المراقبة أو التصحيح بعذر مشروع عليه احضار ما يثبت ذلك وسيتم عرضها على لجنة الامتحانات للنظر في طلبه.</p>
<p>مع الإحترام<br>قسم الإمتحانات</p>
</div>
"@

    Write-Progress -Activity "Sending emails..." -Status "Sending to $to..." -PercentComplete (($count / $total) * 100)

    if (Test-Path $file) {
        try {
            $bytes = [IO.File]::ReadAllBytes($file)
            $base64Content = [Convert]::ToBase64String($bytes)

            $message = @{
                Message = @{
                    Subject = "ملحق المكلفون بالعمل في الثانوية العامة 2025 (المراقبة) / الدفعة السادسة"
                    Body = @{
                        ContentType = "HTML"
                        Content = $bodyHtml
                    }
                    ToRecipients = @(@{ EmailAddress = @{ Address = $to } })
                    Attachments = @(@{
                        "@odata.type" = "#microsoft.graph.fileAttachment"
                        Name = $attachmentName
                        ContentBytes = $base64Content
                    })
                }
                SaveToSentItems = $true
            }

            Send-MgUserMail -UserId $userId -BodyParameter $message
            Write-Host "📧 Sent to $to | 📎 $attachmentName | 🏫 $schoolName"
        }
        catch {
            Write-Warning "❌ Failed to send to $to. Error: $_"
        }
    }
    else {
        Write-Warning "⚠ File not found: $file"
    }
}

Write-Host "`n✅ All emails processed."