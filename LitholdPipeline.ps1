[CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()][string[]]$EmailAddress
    )

BEGIN 
{

    $unqName = Get-Date -UFormat %Y%m%d_%H%M%S

    # Dot source
    . T:\Powershell\Scripts\Gather\GetLog.ps1

    # Check/Connect PSSession

    . T:\Powershell\Scripts\PSSessions\PSSessionALL.ps1
}

PROCESS 
{

#Write-Debug -Message 'Stepping to foreach validation loop' -Debug

    foreach ($Email in $EmailAddress) 
    {

            try 
            {
                Get-AcScriptRunLog -Message "Validating $($Email)" -ScriptName 'LitHold'
                Write-Verbose "Validating $($Email)" -Verbose

                $MailboxInfo = Get-EXOMailbox -Identity $Email -ErrorAction Stop

                $MailboxProps = @{
                    MailboxValid = 'YES'
                    Name = $MailboxInfo.Name
                    Email = $MailboxInfo.PrimarySmtpAddress
                    OnHold = $MailboxInfo.LitigationHoldEnabled
                    PlacedOnDate = $MailboxInfo.LitigationHoldDate
                    HoldDuration = $MailboxInfo.LitigationHoldDuration
                    Setby = $MailboxInfo.LitigationHoldOwner
                    MailServer = $MailboxInfo.Database
                    MailSource = $MailboxInfo.OriginatingServer.split('.')[3]
                }

            }
            catch 
            {
            
                Get-AcScriptRunLog -Message "Errors found?........." -ScriptName 'LitHold'
                Get-AcScriptRunLog -Message "[INVALID] Mailbox $($Email)"
                Write-Warning "Mailbox $($Email) Not Valid" -Verbose
                $MailboxProps = @{
                    MailboxValid = 'NO'
                    Name = "No Mailbox found for $($Email)"
                    Email = $MailboxInfo.PrimarySmtpAddress
                    OnHold = 'N/A'
                    PlacedOnDate = $MailboxInfo.LitigationHoldDate
                    HoldDuration = 'N/A'
                    Setby = 'N/A'
                    MailServer = 'N/A'
                    MailSource = 'N/A'
                }

            }
            finally 
            {
                Get-AcScriptRunLog -Message "[PRE CONFIG] Compiling and exporting data....." -ScriptName "LitHold"
                #Write-Verbose "[PRE CONFIG] Compiling and exporting data....." -Verbose

                $MailboxObjectPre = New-Object -TypeName psobject -Property $MailboxProps

                Write-Output $MailboxObjectPre

                $MailboxObjectPre | Select-Object MailboxValid, Name, Email,OnHold, PlacedOnDate, HoldDuration,Setby,MailServer,MailSource | Export-Csv "T:\Powershell\CSV\Results\$unqName.LitResultCheckPre.csv" -NoTypeInformation -Append

                foreach ($Object in $MailboxObjectPre) 
                {

                    if (($Object.MailboxValid) -eq 'YES' -and ($Object.OnHold -like 'FALSE')) 
                    {
                
                        Get-AcScriptRunLog "$($Object.Email) can be placed on hold" -ScriptName "LitHold"

                        Get-AcScriptRunLog "Placing $($Object.Email) on hold" -ScriptName "Lithold"

                        Set-EXOMailbox -Identity $Object.Email -LitigationHoldEnabled $true -UseDatabaseQuotaDefaults $false

                    }


                    if (($object.MailboxValid) -eq 'YES' -and ($Object.OnHold -like 'TRUE')) 
                    {
                        Get-AcScriptRunLog -Message "$($Object.Email) is a valid mailbox and is already on hold" -ScriptName "LitHold"
                        #Write-Verbose "$($Object.Email) is Valid and is already on hold" -Verbose
                    }


                    if (($object.MailboxValid) -like 'NO' -and ($object.OnHold -like 'N/A')) 
                    {
                        Get-AcScriptRunLog -Message "Object $($Object.email) is not valid and cannot be processed" -ScriptName "LitHold"
                        #Write-Verbose "Object $($Object.email) is not valid and cannot be processed" -Verbose

                    }
                
                        Get-AcScriptRunLog "[POST CONFIG] Gathering and Exporting data for $($Email.Email)" -ScriptName "LitHold"
                
                    foreach ($email in $Email) 
                    {

                        try 
                        {

                          $MailboxInfoPost = Get-EXOMailbox -Identity $Email -ErrorAction stop

                            $MailboxPropsResult = @{
                                MailboxValid = 'Yes'
                                Name = $MailboxInfoPost.Name
                                Email = $MailboxInfoPost.PrimarySmtpAddress
                                OnHold = $MailboxInfoPost.LitigationHoldEnabled
                                PlacedOnDate = $MailboxInfoPost.LitigationHoldDate
                                HoldDuration = $MailboxInfoPost.LitigationHoldDuration
                                Setby = $MailboxInfoPost.LitigationHoldOwner
                                MailServer = $MailboxInfoPost.Database
                                MailSource = $MailboxInfo.OriginatingServer.split('.')[3]
                          }
                        } 
                        catch 
                        {

                            $MailboxPropsResult = @{
                                MailboxValid = 'No'
                                Name = "No Mailbox found for $($Email)"
                                Email = "$($email.email)"
                                PlacedOnDate = 'N/A'
                                OnHold = 'N/A'
                                HoldDuration = 'N/A'
                                Setby = 'N/A'
                                MailServer = 'N/A'
                                MailSource = 'N/A'
                            }
                        }

                        finally 
                        {
                        
                          $MailboxObjectPost = New-Object -TypeName psobject -Property $MailboxPropsResult

                          Write-Output $MailboxObjectPost

                          $Results = $MailboxObjectPost | Select-Object MailboxValid, Name, Email, OnHold, PlacedOnDate, HoldDuration,Setby,MailServer,MailSource | Out-String
                          
                          Get-AcScriptRunLog -Message "Results: `r`n $Results" -ScriptName "LitHold"

                          $Results | Export-Csv "T:\Powershell\CSV\Results\$unqName.LitResultCheckPost.csv" -NoTypeInformation -Append

                       } # end try/catch $MailboxInfoPost
                    } # end foreach try/catch $MailboxInfoPost
                } # end foreach $MailboxObject
            } # end finally
        } #end foreach-object $MailboxInfo
} # end PROCESS

END {
    Write-Verbose "Script ended" -Verbose

    # END PSSsession

    $EndSesh = Read-Host "Would you like to remove all PSSessions? Y/N"

    if ($EndSesh -like 'y') {
    
        Get-PSSession | Remove-PSSession -Verbose
    
    } else {
    
    break
    
    }

}

