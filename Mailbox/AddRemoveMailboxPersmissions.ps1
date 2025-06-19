# Install if not already done
Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.2.0

# connect 
Connect-ExchangeOnline -UserPrincipalName badarin2050@hebron.edu.ps -ShowProgress $true

# Add full permission to the mailbox
Add-MailboxPermission -Identity zain.jabari@hebron.edu.ps -User badarin2050@hebron.edu.ps -AccessRights FullAccess -InheritanceType All

https://outlook.office.com/mail/zain.jabari@hebron.edu.ps

# Remove full permission from the mailbox
Remove-MailboxPermission  -Identity 26111019@hebron.edu.ps -User badarin2050@hebron.edu.ps -AccessRights FullAccess -InheritanceType All
