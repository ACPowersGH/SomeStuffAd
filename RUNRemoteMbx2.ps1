# Must establish connection to MYIGT ECP to create the remote mailbox
# [VIP(Field Not Required for Non-VIP’s)].[CountryAbbr].[State/ProvAbbr(optional)].[City].[AdditionalIdentifier(optional)].[RoomName].[Device Type(Optional - Display Name Only)]

# Observation:
# AD replication completed before O365: AD account was created before O365

#function New-PoshRemoteMailbox {

    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Medium')]
    param
    (
     [Parameter()]
     [ValidateNotNullOrEmpty()]
     [switch]$Force
    )

BEGIN {
    
    $InformationPreference = "Continue"
    $VerbosePreference = "Continue"
    $DebugPreference = "Continue"

    Write-Information "IN BEGIN BLOCK"

    Write-Information "Checking PSSessions"

    . T:\Powershell\Scripts\PSSessions\PSSessionALL.ps1

    # https://becomelotr.wordpress.com/2013/05/01/supports-should-process-oh-really/
    $PSBoundParameters.Remove('Force') | Out-Null
    $PSBoundParameters.Confirm = $false

    # Import saved keys
    $MyIGTId = Get-Content 'T:\Powershell\Key\WSIGTUN.txt'
    $MyIGTPW = Get-Content 'T:\Powershell\Key\WSIGTPW.txt' | ConvertTo-SecureString
    $MyIGT = New-Object System.Management.Automation.PSCredential ($MyIGTId,$MyIGTPw)

    # DOT Source log function

    . T:\Powershell\Scripts\Gather\GetLog.ps1

    # SET CSV PATH

    $path = 'T:\Powershell\CSV\RunData\Book.csv'

    # IMPORT CSV

    $Data = Import-Csv -Path $path

    $Data 

    #pause

    $countO365 = 1
    $countAD = 1

} PROCESS {

        # CREATE REMOTE MAILBOX IN MYIGT ECP

        Write-Information "IN PROCESS BLOCK"

    foreach ($object in $Data) 
    {
    $Error.clear()

        Write-Information "PROCESSING FOREACH"

            if ($Force -or $PSCmdlet.ShouldProcess($object.Name,"Create Mailbox")){

                try {

                    Get-PoshRunLog -Message "Creating mailbox $($object.UPN) > $(Get-Date -UFormat %h%dth%Y_T_%H%M%S)" -ScriptName 'RUNRemoteMbx'

                    #Write-Debug "Creating mailbox $($object.MbxSMTP)"
                    Write-Verbose "Creating mailbox $($object.MbxSMTP)"

                    New-MyIGTRemoteMailbox -Name $object.Name -Alias $object.Alias -SamAccountName $object.SamAccountName  -Room -RemoteRoutingAddress $object.RemoteRouting -PrimarySmtpAddress $object.MbxSMTP -UserPrincipalName $object.UPN -Password (ConvertTo-SecureString -String $object.Password -AsPlainText -Force) -OnPremisesOrganizationalUnit $object.MbxOU -ErrorAction Stop | Out-Null
                    $checkRemote = Get-MyIGTRemoteMailbox -Identity $object.Alias

                } catch {

                    Write-Warning "[ERROR] $($object.MbxSMTP) Please review log for details"
                    $Errorflag = $Error.exception.message.ToUpper()
                    $Errorflag
                    Get-PoshRunLog -Message "[ERROR (See exported csv error file)]: $($object.MbxSMTP), $Errorflag" -ScriptName 'RUNRemoteMbx'
                    break

                }
            } # End $force Create Mailbox

            # CHECK IF MAILBOX IS IN O365
            If ($checkRemote) 
            {

                Get-PoshRunLog -Message "$($checkRemote.Name) was successfully created > $(Get-Date -UFormat %h%dth%Y_T_%H%M%S)" -ScriptName 'RUNRemoteMbx'
                    Write-Verbose "$($checkRemote.Name) was successfully created"

                        do 
                        {
                            
                            #Write-Debug -Message "Get-MsolUser for $($object.UPN)"
                            Get-PoshRunLog -Message "Checking if $($object.UPN) exists in O365..sleeping for 60 seconds > $(Get-Date -UFormat %h%dth%Y_T_%H%M%S)" -Scriptname 'RUNRemoteMbx'
                            Write-Verbose -Message "Checking if $($object.UPN) exists in O365..sleeping for 60 seconds > $(Get-Date -UFormat %h%dth%Y_T_%H%M%S)"

                            $checkO365 = Get-MsolUser -UserPrincipalName $object.UPN -ErrorAction SilentlyContinue

                            Start-Sleep -s 60

                            $numberO365 = $countO365++

                            Write-Verbose -Message "O365 has been checked $($numberO365) time(s); Time = $($numberO365 * 60)secs."
                            Get-PoshRunLog -Message "O365 has been checked $($numberO365) time(s); Time = $($numberO365 * 60)secs" -ScriptName 'RUNRemoteMbx'

                       } 
                       until (($checkO365 | Measure-Object).count -eq 1)


                        Get-PoshRunLog -Message "Room Mailbox $($object.Name) is now in O365 found on  > $(Get-Date -UFormat %h%dth%Y_T_%H%M%S).`r`nOK to proceed with configurations" -ScriptName 'RUNRemoteMbx'

                        Write-Verbose -Message "Room Mailbox $($object.Name) is now in O365.. found.`r`nOK to proceed with configurations"

                        Get-MsolUser -UserPrincipalName $object.UPN | Select-Object UserPrincipalName,DisplayName

                        #pause 

                        # ASSIGN ACTIVESYNC POLICY

                        Write-Verbose -Message "Assigning AS policy to $($object.UPN)"
                        Get-PoshRunLog -Message "Assigning AS policy to $($object.UPN)" -ScriptName 'RUNRemoteMbx'

                        if ($Force -or $PSCmdlet.ShouldProcess($($object.Name),"Assign AS policy Surfacehub")) 
                        {

                                Write-Verbose -Message "Assigning ActiveSync Policy to $($object.UPN)" 
                                Set-O365CasMailbox -Identity $object.UPN -ActiveSyncMailboxPolicy Surfacehub

     
                            }

            }
                       #pause

                        # CONFIGURE CALENDAR PROCESSING

                        Write-Verbose -Message "Configuring Cal Processing $($object.UPN)"
                        Get-PoshRunLog -Message "Configuring Cal Processing $($object.UPN) > $(Get-Date -UFormat %h%dth%Y_T_%H%M%S)" -ScriptName 'RUNRemoteMbx'

                        if ($Force -or $PSCmdlet.ShouldProcess($($object.UPN),"Configure calendar processing")) 
                        {

                            try 
                            {

                                Set-O365CalendarProcessing -Identity $object.UPN -AutomateProcessing AutoAccept -AddOrganizerToSubject $false -AllowConflicts   $false -DeleteComments $false -DeleteSubject $false -RemovePrivateProperty $false -AddAdditionalResponse $true -AdditionalResponse "This room is equipped with a Surface Hub"

                            } 
                            catch 
                            {

                                throw

                            }
                        }

                        #pause

                        # ASSIGN O365 USAGE LOCATION

                        #Write-Debug -Message "Setting O365 Usage Location $($object.UPN)"
                        Write-Verbose -Message "Setting O365 Usage location $($object.UPN)"

                        Get-PoshRunLog -Message "Setting O365 Usage location $($object.UPN) > $(Get-Date -UFormat %h%dth%Y_T_%H%M%S)" -ScriptName 'RUNRemoteMbx'

                        if ($Force -or $PSCmdlet.ShouldProcess($($object.UPN),'Assign UsageLocation')) 
                        {

                            try 
                            {

                                Set-MsolUser -UserPrincipalName $object.UPN -UsageLocation US

                            } 
                            catch 
                            {

                                throw

                            }

                        }

                        #pause

                        # ASSIGN O365 TEAMS LICENSE

                        #Write-Debug -Message "Assigning O365 license $($object.UPN)"
                        Write-Verbose -Message "Assigning O365 license $($object.UPN)"

                        Get-PoshRunLog -Message "Assigning O365 license $($object.UPN)" -ScriptName 'RUNRemoteMbx'

                        if ($Force -or $PSCmdlet.ShouldProcess($($object.Name),'Assign MsolUserLicense > gtechcorp:TEAMS_COMMERCIAL_TRIAL')) 
                        {

                            try 
                            {

                                Set-MsolUserLicense -UserPrincipalName $object.UPN -AddLicenses "gtechcorp:TEAMS_COMMERCIAL_TRIAL"

                            } 
                            catch 
                            {

                                $teamsErrors = $Error

                                $teamsErrors | Select-Object @{l='Mailbox_Email';e={$object.MbxSMTP}},@{l='ErrorName';e={$_.categoryinfo.Category}},@{l='Object';e={$_.categoryinfo.TargetName}} | Export-Csv T:\Powershell\Results\ErrorRUNRemoteMailbox.csv -Append -Force -NoTypeInformation

                            }

                        }

                        #pause

                        # Check for S4B PSSession

                        $S4B = Get-PSSession | Where-Object {$_.ComputerName -like '*admin0a*' -and $_.Availability -eq 'Available'}

                        if (!$S4b) 
                        {
                            Write-Warning -Message 'No S4B Session found. Import PSSession'
                            . T:\Powershell\Scripts\PSSessions\S4B.ps1                                                    
                        }                   
                        else
                        {
                            Write-Verbose -Message "S4B Session found. Please continue"
                        }

                        # Enable for MS Teams

                        #Write-Debug -Message "Enable $($object.UPN) for MS Teams"
                        Write-Verbose "Enabling $($object.UPN) for MS Teams"

                        Get-PoshRunLog -Message "Enabling $($object.UPN) for MS Teams > $(Get-Date -UFormat %h%dth%Y_T_%H%M%S)" -ScriptName 'RUNRemoteMbx'

                        if ($Force -or $PSCmdlet.ShouldProcess($($object.Name),"Enable for MS Teams")) 
                        {
                            Do 
                            {
                                $checkEnabled = Get-S4BCsMeetingRoom -Identity $object.Name -ErrorAction SilentlyContinue

                                try 
                                {
                                        
                                    Write-Verbose "Attempting to enable $($object.Name) for MS Teams"
                                    Enable-S4BCsMeetingRoom -Identity $object.Name -SipAddressType "EmailAddress" -RegistrarPool "sippoolDM10a24.infra.lync.com" -ErrorAction SilentlyContinue
                                    
                                    Start-Sleep -s 30
                                } 
                                catch [System.Management.Automation.RemoteException]
                                {
                                        
                                    Write-Error -Message "$($object.UPN) not found; propagation to O365 failed"
                                    $enableErrors = $Error
                                    $enableErrors | Select-Object @{l='Mailbox_Email';e={$object.MbxSMTP}},@{l='ErrorName';e={$_.categoryinfo.Category}},@{l='Object';e={$_.categoryinfo.TargetName}} | Export-Csv T:\Powershell\Results\ErrorRUNRemoteMailbox.csv -Append -Force -NoTypeInformation
                                }
                            }
                            Until ($checkEnabled.Enabled -eq $true)

                                Write-Verbose "$($object.Name) has been enabled for MS Teams"
                        }
                        
                        #pause
                  

                    # CHECK IF AD object exist in MYIGT

                    Write-Information -MessageData "Checking MYIGT Domain for $($object.Alias)" 

                    if ($Force -or $PSCmdlet.ShouldProcess($object.Alias,"Enabled MYIGT AD Account")) 
                    {
                        Do
                        {
                            $checkADUser = Get-ADUser -Identity $object.SamAccountName
                                try
                                {
                                    Write-Verbose 'Attempting to enable AD object'                            
                                    Enable-ADAccount -Identity $object.SamAccountName -Credential $MyIGT
                                }
                                catch
                                {
                                    Write-Error -Message "$($object.Alias) not found in MYIGT AD"
                                    $adErrors = $error.Exception.Message
                                    Get-PoshRunLog -Message "$adErrors" -ScriptName 'RUNRemoteMbx'
                                }
                        }
                        Until ($checkADUser.Enabled -eq $true)

                            Write-Verbose -Message "AD object $($object.Alias) is enabled"
                            
                    }
                    Write-Information -MessageData "Script execution for remote mailbox creation complete" 
                    Get-PoshRunLog -Message "Script execution for remote mailbox creation complete" -ScriptName 'RUNRemoteMbx'
            } # END IF ELSE
    } # END FOREACH

 } # END PROCESS

END {


}

#} # END function

#New-PoshRemoteMailbox