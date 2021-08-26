Function Log-ScriptRun {

param (
            [string]$Message,
            [string]$path="e:\Data\CSV_Exports\Logs\ScriptGetPoshGroupMembers.log"
            )

            $Message | Out-File -Filepath $path -Append
}

function Get-PoshGroupMembers ([string[]]$DGName) {

    #$AllusersWorld = $DGName | Get-DistributionGroupMember -ResultSize unlimited

    foreach ($DG in $DGName) {

        $RawDGData = $DG | Get-DistributionGroupMember -ResultSize unlimited

            foreach ($Member in $RawDGData) {

            $Unique = Get-Date -Format yyyymmdd


                Switch ($Member.RecipientType) {

                    "MailUniversalDistributionGroup" {
        
                        "$($Member.Name) is a DG"
    
                        Get-DistributionGroupMember -Identity $Member.Name -ResultSize unlimited | select @{l='DGName';e={$Member.Name}},Name,PrimarySmtpAddress,Office,Title,CustomAttribute1,CustomAttribute5,CustomAttribute6,RecipientType | Export-Csv "\\$env:COMPUTERNAME\e$\Data\CSV_Exports\Announcements_TS\$Unique.$($Member.Name).csv" -NoTypeInformation -Append
    
    
                    }

                    "DynamicDistributionGroup" {
    
                        "$($Member.Name) is a DDG"

                        $DDG = Get-DynamicDistributionGroup $Member.Name
                        Get-Recipient -RecipientPreviewFilter $DDG.RecipientFilter -ResultSize unlimited | select @{l='DDGName';e={$($DDG.Name)}},Name,PrimarySmtpAddress,Office,Title,CustomAttribute1,CustomAttribute5,CustomAttribute6,RecipientType | Export-Csv "\\$env:COMPUTERNAME\e$\Data\CSV_Exports\Announcements_TS\$Unique.$($Member.Name).csv" -NoTypeInformation -Append    

                    }

                    "MailContact" {
            
            
                        Write-Verbose "[MAILCONTACT OBJECT FOUND] No action required for $($Member.Name)" -Verbose

                        Log-ScriptRun -Message "$($Member.Name) is a MAILCONTACT"
                                           
                    }

                    Default {
            
            
                    Write-Warning "No data found matching RecipientType MailUniversalDistributionGroup or DynamicDistributionGroup for $($Member.Name)"
            
            
            
                    }


                }

            }

    } #end foreach $DGNAME

} # end function