# https://learn-powershell.net/2014/01/24/avoiding-system-object-or-similar-output-when-using-export-csv/

PS C:\Users\cuencaa> $belrooms |ForEach-Object {

    $calperms = Get-MailboxFolderPermission "$($_.primarysmtpaddress):\calendar"

    $calprops = @{
                    RoomName=$_.Name
                    FolderType=@($calperms.FolderName | Out-String).Trim()
                    User= @(($calperms.User).DisplayName | Out-String).Trim()
                    Rights=@($calperms.AccessRights | Out-String).Trim()
                }
                
                $cals=New-Object -TypeName psobject -Property $calprops
                
                Write-Output $cals
                
                $cals | select RoomName,FolderType,Rights,User | Export-Csv T:\Powershell\RSBelCalRoomPerm.csv -NoTypeInformation -Append
}

<#

Sample Run:


User                              Rights                 FolderType                   RoomName
----                              ------                 ----------                   --------
Default...                        Reviewer...            Calendar...                  RS Belgrade - Colosseum
Default...                        Reviewer...            Calendar...                  RS Belgrade - Hyatt
Default...                        Reviewer...            Calendar...                  RS Belgrade - MS
Default...                        Reviewer...            Calendar...                  RSBEGConfRoom2
Default...                        AvailabilityOnly...    Calendar...                  RSBEGOfficeDejan
Default...                        Reviewer...            Calendar...                  RSBEGConfRoom1
#>