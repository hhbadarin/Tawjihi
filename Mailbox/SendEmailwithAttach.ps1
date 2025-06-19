#Connect to ExchangeOnline 
Connect-ExchangeOnline -UserPrincipalName badarin2050@hebron.edu.ps

# Send email to single user with cc and bcc 
$sender = "badarin2050@hebron.edu.ps"
$recipient = "123456@hebron.edu.ps"
$ccrecipient = "123456@hebron.edu.ps"
$bccrecipient = "123456@hebron.edu.ps"
$subject = "نشاط المعلمين على منصة تيمز"
$body = "السلام عليكم، مرفق نشاط المعلمين على منصة تيمز في الفترة الممتدة من "
$attachmentpath = "$Home\Documents\GitHub\Microsoft-Graph\Photos\Palestine.png"
$attachmentmessage = [Convert]::ToBase64String([IO.File]::ReadAllBytes($AttachmentPath))
$attachmentname = (Get-Item -Path $attachmentpath).Name
$type = "HTML" #Or you can choose "Text"
$save = "true" #Or you can choose "true"

# Connect to MgGraph with required permissions
Connect-MgGraph -Scopes 'Mail.Send', 'Mail.Send.Shared'

$params = @{
    Message         = @{
        Subject       = $subject
        Body          = @{
            ContentType = $type
            Content     = $body
        }
        ToRecipients  = @(
            @{
                EmailAddress = @{
                    Address = $recipient
                }
            }
        )
        CcRecipients  = @(
            @{
                EmailAddress = @{
                    Address = $ccrecipient
                }
            }
        )
        BccRecipients = @(
            @{
                EmailAddress = @{
                    Address = $bccrecipient
                }
            }
        )
        Attachments   = @(
            @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                Name          = $attachmentname
                ContentType   = "text/plain"
                ContentBytes  = $attachmentmessage
            }
        )
    }
    SaveToSentItems = $save
}
# Send message
Send-MgUserMail -UserId $sender -BodyParameter $params