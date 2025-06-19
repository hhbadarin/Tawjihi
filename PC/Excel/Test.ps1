
#Connect to the Microsoft Graph using Device code interactive flow
Connect-Graph -Scopes "User.Read.All", "UserAuthenticationMethod.ReadWrite.All"


$Users = Import-CSV "$Home\Desktop\A.csv"
ForEach ($user in $Users) {
    $UserIDs = [pscustomobject]@{
        UserPrincipalName = $user.UserPrincipalName
        ID      = (Get-MgUser -UserId $user.UserPrincipalName).id
        }
    $UserIDs | Export-CSV -Path "$Home\Desktop\B.csv" -Append -NoTypeInformation   
    }


#Export users in a securtiy group
Get-MgGroupMember -GroupId '4bc85716-a5da-4481-9469-75fa02a6b47b' -All | Select-Object Id | Export-CSV -Path "$Home\Desktop\AA.csv" -Encoding UTF8 -NoTypeInformation

Get-mguser -All | Select-Object displayName,id, UserPrincipalName |Export-CSV -Path "$Home\Desktop\Allusers.csv" -Encoding UTF8 -NoTypeInformation


#Filter Secs 
Import-Csv "$Home\Desktop\roles.csv" | Where-Object {$_.role -eq 'student'} | Export-Csv -NoTypeInformation -Encoding UTF8 "$Home\Desktop\students.csv" 



# --------------- Excel ------ 

Install-Module -Name ImportExcel

$report = Import-Csv "$Home\Desktop\staffReport_20241118.csv" | Select-Object @{expression={$_."اسم الموظف"}; label='الإسم'},@{expression={$_."رقم الهوية"}; label='رقم الهوية'}, @{expression={$_."الوظيفة"}; label='الوظيفة'},@{expression={$_."اسم المدرسة"}; label='المدرسة'},@{expression={$_."الرقم الوطني"}; label='الرقم الوطني'} | Sort-Object {$_."الرقم الوطني"} -OutVariable staff | Export-Excel -Path "$Home\Desktop\staff_$((Get-Date).ToString('dd-MM-yyyy')).xlsx"  -WorksheetName "staff_$((Get-Date).ToString('dd-MM-yyyy'))" -TableName "staff" -TableStyle Light9 -AutoSize

Import-Csv "$Home\Documents\Github\Microsoft-Graph\Files\newUsers.csv" | Where-Object {$_.Surname -eq $null} | Export-Csv -NoTypeInformation -Encoding UTF8 "$Home\Desktop\Null.csv"


Import-Csv "$Home\Desktop\T1.csv" | Select-Object @{expression={$_.displayName}; label='الإسم'},@{expression={$_.jobTitle}; label='الوظيفة'},@{expression={$_.department}; label='المدرسة'},@{expression={$_.UserPrincipalName}; label='حساب تيمز'}  | Export-Csv -NoTypeInformation -Encoding UTF8 "$Home\Desktop\T2.csv"


$report = Import-Csv "$Home\Desktop\T2.csv"
$report | Export-Excel -Path "$Home\Desktop\T3.xlsx" -WorksheetName "حسابات تيمز لمعلمين جدد" -TableName "newteachers" -TableStyle Light9 -FreezeTopRow -BoldTopRow -AutoSize -Show


#Get all users created before a certain date 
Get-MgUser -All -Filter {$_.Enabled -like $false} | Select-Object Id, DisplayName, UserPrincipalName | Export-Csv -Path "$Home\Desktop\Diabled.csv" -Encoding UTF8



#Export members of Teachers security group
Get-MgGroupMember -GroupId '6ebc61b5-1252-4296-b440-2bc706411d92' -All | Select-Object Id | Export-Csv "$Home\Desktop\TeacherSecurityGroup.csv" -NoTypeInformation -Encoding UTF8 


