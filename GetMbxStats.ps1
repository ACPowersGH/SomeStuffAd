function cus_Get-Mbxstats {
    [CmdletBinding ()]
    param ($userID)

    Get-MailboxStatistics $userID | select @{label="DisplayName";e={Get-Mailbox -Identity $userID | select -ExpandProperty DisplayName}},@{label='USERID';e={Get-Mailbox $userID | select -ExpandProperty samaccountname}},@{label='Email';e={Get-Mailbox -Identity $userID | select -ExpandProperty primarysmtpaddress}},servername,database,itemcount

}