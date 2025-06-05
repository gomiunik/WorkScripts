Add-Type -assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -comobject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")

$inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)
# $sent = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderSentMail)

$senderEmails = @{}

function Read-Email-Addresses($folder) {

    foreach ($mail in $folder.Items) {
        try {
            # Ensure it's a MailItem
            if ($mail -is [Microsoft.Office.Interop.Outlook.MailItem]) {
                $email = $mail.SenderEmailAddress
                if (![string]::IsNullOrWhiteSpace($email)) {
                    $senderEmails[$email] = $true
                }
            }
        } catch {
            # Ignore items that throw errors (e.g., MeetingItem or ReportItem)
            continue
        }
    }

    foreach ($subfolder in $folder.Folders) {
        Read-Email-Addresses($subfolder)
    }
}

Read-Email-Addresses($inbox)


# Output unique sender email addresses
$senderEmails.Keys | Sort-Object