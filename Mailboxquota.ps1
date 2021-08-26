# Set mailbox quota

function Set-UserMbxSize {

[CmdletBinding()]
param ($MailboxEmail)

    Set-Mailbox -Identity $MailboxEmail -ProhibitSendQuota 49.75GB -ProhibitSendReceiveQuota 50GB -IssueWarningQuota 49.5GB -WhatIf

}

Set-UserMbxSize