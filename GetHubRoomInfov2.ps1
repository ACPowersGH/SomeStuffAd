#function cus_Get-HubRoomInfo {

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true,
               HelpMessage='Enter conference room names',
               ValueFromPipeline=$true)]
               [array]$HubRoomNames
)

BEGIN {}

PROCESS {      
    
    #Variable for csv file name classifying date of export
    $unqName = Get-Date -UFormat %Y%m%d_%H%M%S

    $roomcounter = 0

    foreach ($hub in $HubRoomNames) {

    $roomcounter++
    
        Write-Verbose "==== Start of Gathering Data ===="
        Write-Verbose "Getting Mailbox Info for $hub"

        $mailboxInfo = Get-Mailbox -Identity $hub
        $mailboxCas = Get-CASMailbox -Identity $hub
        $mailboxPermissions = Get-MailboxPermission -Identity $hub -User "<>"
        $calendarProcessing = Get-CalendarProcessing -Identity $hub
        $userAD = (Get-ADUser -Identity ($mailboxInfo.SamAccountName) -Properties * -Server gtk.gtech.com) 

        $Properties = @{HubRoomName = $hub
                            DeviceAccount = $mailboxInfo.SamAccountName
                            DeviceAccountPassword = "3asp*cuY"
                            MailboxType = $mailboxInfo.RecipientType
                            MailboxTypeDetails = $mailboxInfo.RecipientTypeDetails
                            MailEnabled = $mailboxInfo.RoomMailboxAccountEnabled
                            MailboxHidden = $mailboxInfo.HiddenFromAddressListsEnabled
                            ActiveSyncEnabled = $mailboxCas.ActiveSyncEnabled
                            ActiveSyncPolicy = $mailboxCas.ActiveSyncMailboxPolicy
                            CalendarMeetingProcessing = $calendarProcessing.AutomateProcessing
                            ExternalMeetingProcessing = $calendarProcessing.ProcessExternalMeetingMessages
                            OrganizerVisible = $calendarProcessing.AddOrganizerToSubject
                            LyncSignInAddress = $userAD.UserPrincipalName
                            MailboxOwner = $mailboxPermissions.User
                            MailboxRights = $mailboxPermissions | select -ExpandProperty AccessRights
                            Notes = $userAD.Info
                            }

        $obj = New-Object -TypeName PSObject -Property $Properties
        
        Write-Output $obj

        Write-Progress -Activity 'Working' -Status "Getting info for $hub" -PercentComplete (($roomcounter/$HubRoomNames.count)*100)
        
        $result = $obj | select HubRoomName,DeviceAccount,DeviceAccountPassword,MailboxType,MailboxTypeDetails,MailEnabled,MailboxHidden,ActiveSyncEnabled,ActiveSyncPolicy,CalendarMeetingProcessing,ExternalMeetingProcessing,LyncSignInAddress,MailboxOwner,MailboxRights,OrganizerVisible,Notes #Export-Csv "\\$env:COMPUTERNAME\e$\Temp\$unqName.Hub_Configuration_Summary.csv" -NoTypeInformation -Append -NoClobber

        $result | Out-File "\\$env:COMPUTERNAME\e$\#_Posh\Info\_Results\_SurfaceHub\$unqName._SurfaceHub.Results.txt" -Append -NoClobber
    }


}

END {}

#}#end cus_Get-HubRoomInfo