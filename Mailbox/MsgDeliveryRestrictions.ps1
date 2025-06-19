
#Connect to ExchangeOnline 
Connect-ExchangeOnline -UserPrincipalName badarin2050@hebron.edu.ps


#Require sender to be trusted
Set-Mailbox -Identity 26112276@hebron.edu.ps -RequireSenderAuthenticationEnabled $true

#Require senders to be trusted (bulk)
$Users=  Import-Csv -Path "$Home\Documents\Github\Microsoft-Graph\Files\RestrictdDeliveryUsers.csv"
$i = 0
foreach ($User in $Users) {

try {
    $i++
    Set-Mailbox -Identity $User.UserPrincipalName -RequireSenderAuthenticationEnabled $true
    Write-Host ("User {0} is now restricted from sending outside Your org ... ({1}/{2})" -f $User.UserPrincipalName, $i, $Users.Count) -ForegroundColor Yellow
}
catch {
    Write-Host ($_.Exception.Message) -ForegroundColor Yellow
}
}


############# 

#Don't require single sender to be trusted
Set-Mailbox -Identity 26112276@hebron.edu.ps -RequireSenderAuthenticationEnabled $false


#Don't require senders to be trusted bulk
$Users=  Import-Csv -Path "$Home\Documents\Github\Microsoft-Graph\Files\RestrictdDeliveryUsers.csv"
$i = 0
foreach ($User in $Users) {

try {
    $i++
    Set-Mailbox -Identity $User.UserPrincipalName -RequireSenderAuthenticationEnabled $false
    Write-Host ("User {0} is no longer restricted from sending outside Your org ... ({1}/{2})" -f $User.UserPrincipalName, $i, $Users.Count) -ForegroundColor Yellow
}
catch {
    Write-Host ($_.Exception.Message) -ForegroundColor Yellow
}
}