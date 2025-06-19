#Connect to Graph with the scopes
Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All","Mail.Send"


# Send tawjihi data to multiple schools (Method 2)
$users = Import-Csv -Path "~\Desktop\Tawjihi2025.csv"
$schools = $users | Group-Object "المدرسة"
$schools 
foreach ($school in $schools) {
    $departmentName = $school.Name
    $successfulUsers = $school.Group 
    $postalCode = $successfulUsers[0]."الرقم الوطني"
    $schoolEmail = "$postalCode@hebron.edu.ps"
    $counter = 1
    $tableRows = foreach ($user in $successfulUsers) {
        $row = "<tr>
            <td>$counter</td>
            <td>$($user."اسم الموظف")</td>
            <td>$($user."الهوية")</td>
            <td>$($user."المدرسة")</td>
            <td>$($user."الوظيفة")</td>
            <td>$($user."المهام")</td>
            <td>$($user."القاعة")</td>
             <td>$($user."السكن")</td>
              <td>$($user."الهاتف")</td>
        </tr>"
        $counter++
        $row
    }

    $htmlTable = @"
<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse; width:100%; font-family:Segoe UI, sans-serif; font-size:16px; direction:rtl;'>
    <thead style='background-color:#f2f2f2;'>
        <tr>
            <th>الرقم</th>
            <th>اسم الموظف</th>
            <th>الهوية</th>
            <th>المدرسة</th>
            <th>الوظيفة</th>
            <th>المهام</th>
            <th>القاعة</th>
            <th>السكن</th>
            <th>الهاتف</th>
        </tr>
    </thead>
    <tbody>
        $(($tableRows -join "`n"))
    </tbody>
</table>
<p style='margin-top:20px;'>مع الشكر<br></p>
"@

    $mailBody = @{
        Subject = "المكلفين بالعمل في الثانوية العامة من مدرسة $departmentName"
        Body = @{
            ContentType = "HTML"
            Content = @"
<div style='direction:rtl; font-family:Segoe UI, sans-serif; font-size:16px;'>
<p>السادة مدرسة <strong>$departmentName</strong> المحترمين،</p>
<p>نرفق لكم قائمة المكلفين بالعمل في الثانوية العامة للعام 2025 من مدرستكم:</p>
$htmlTable
</div>
"@
        }
        ToRecipients = @(
            @{
                EmailAddress = @{
                    Address = $schoolEmail
                }
            }
        )
        # CcRecipients = @(
        #     @{
        #         EmailAddress = @{
        #             Address = "it@hebron.edu.ps"
        #         }
        #     }
        # )
    }

    try {
        Send-MgUserMail -UserId "badarin2050@hebron.edu.ps" -Message $mailBody -SaveToSentItems
        Write-Host "📤 Email sent to $schoolEmail (CC: it@hebron.edu.ps)"
    } catch {
        Write-Warning "⚠️ Failed to send email to $($schoolEmail): $_"
    }
}
