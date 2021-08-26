# Require -runas administrator
# Require Connection to Active Directory for domain
# Updated for Connect-O365PSSession

[CmdletBinding(SupportsShouldProcess=$true)]
param
(
 [parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$File,
 $ScriptSave,$SavePath,
 [switch]$Force
)

begin {

# https://becomelotr.wordpress.com/2013/05/01/supports-should-process-oh-really/
$PSBoundParameters.Remove('Force') | Out-Null            
$PSBoundParameters.Confirm = $false 


# DOT Source log function

. T:\Powershell\Scripts\Gather\GetLog.ps1

$unqName = Get-Date -Format yyyyMMdd_hhmm


# Import data to parse

$Data = Import-Csv -Path "$File"

# Declare counters

$CheckValid = 1
$CheckRemove = 1
$CheckNew = 1
$AddNew = 1
$DelMem = 1

# Get members for distribution group that will need to updated

$GetGroup = Read-Host "Please enter a valid Distribution Group"

$DistroGroup = Get-O365DistributionGroup -Identity $GetGroup -ErrorAction silentlycontinue

# Check if $GetGroup value enterd is a valid Distribution List
# If invalid continue to prompt

if (!$DistroGroup){
    
    Do {
    
        $GetGroup = Read-Host "ENTRY INVALID: Please re-enter a valid Distribution Group"

        $DistroGroup = Get-O365DistributionGroup -Identity $GetGroup -ErrorAction silentlycontinue

    } Until ($DistroGroup)
}

Write-Verbose "$DistroGroup is a valid Distributution Group. Getting members for $($DistroGroup.PrimarySmtpAddress)" -Verbose

$CurrentDGMembers = @(Get-O365DistributionGroupMember -Identity $DistroGroup.PrimarySmtpAddress -ResultSize Unlimited)

# Export current members POST update

$CurrentDGMembers | Select-Object Name,Alias,PrimarySmtpAddress,RecipientType,RecipientTypeDetails | Export-Csv "$SavePath\$DistroGroup.MembersPREUpdate_$unqName.csv" -NoTypeInformation -Append

$StartTime = ($stopwatch =  [system.diagnostics.stopwatch]::StartNew()).Elapsed.ToString()

Get-PoshRunLog -Message ">>>>>> Script started @ $(Get-Date -UFormat %Y%m%d_%H%M%S) and timer begins @ $($StartTime) <<<<<<`r" -ScriptName "$ScriptSave"

# Declare empty array

$InValRecipients= @()
$ValRecipients = @()
$NewRecipients = @()
$RemRecipients = @()

}
process {

Get-PoshRunLog -Message "=================================== CHECKING PROVIDED LIST FOR VALID O365 RECIPIENTS ===================`r" -ScriptName "$ScriptSave"

Write-Verbose -Message "====================================== CHECKING PROVIDED LIST FOR VALID O365 RECIPIENTS ===================`r" -Verbose


    foreach ($member in $Data) 
    {
    $Error.Clear()
                
        $Recipient = Get-O365Recipient -Identity $member.Email -ErrorAction SilentlyContinue | Where-Object {$_.RecipientTypeDetails -match "UserMailbox|MailContact|MailUniversalDistributionGroup|MailUniversalSecurityGroup"}
       
        if ($Recipient)
        {
            Write-Verbose "[VALID RECIPIENT]: $($Recipient.PrimarySmtpAddress)" -Verbose
            $RecipientInfo = Get-O365Recipient -Identity $Recipient.DistinguishedName
            $RecipientInfo | Select-Object Name,Alias,PrimarySmtpAddress,RecipientType,RecipientTypeDetails | Export-Csv "$SavePath\$DistroGroup.VALID_$unqName.csv" -NoTypeInformation -Append 
            $ValRecipients += $Recipient.DistinguishedName            
            Get-PoshRunLog -Message "[VALID RECIPIENT]: $($Recipient.PrimarySmtpAddress) exist in O365" -ScriptName "$ScriptSave"
        }
        else
        {            
            Write-Verbose "[INVALID RECIPIENT]: $($member.Email) < not found" -Verbose
            $InValRecipients += $member.email
            $InValRecipients | Select-Object @{n='Email';e={$_}},@{n='Reason';e={'EmailAddressNotFound'}} | Export-Csv "T:\Powershell\Results\$DistroGroup.INVALID_$unqName.csv" -NoTypeInformation -Append
            Get-PoshRunLog -Message "[INVALID RECIPIENT]: $($member.Email) does not exist in O365" -ScriptName "$ScriptSave"
        }
            Write-Progress -Activity "Checking if object is a valid O365 recipient" -Status "[$($member.Email)]: $CheckValid of $($Data.count)" -PercentComplete ($($CheckValid/$Data.count)*100) 
            $CheckValid++
    }

Get-PoshRunLog -Message "=================================== END CHECKING PROVIDED LIST FOR VALID RECIPIENTS ====================`r" -ScriptName "$ScriptSave"

Get-PoshRunLog -Message "=================================== CHECKING DISTRO MEMBERSHIP FOR ACCOUNTS NOT IN PROVIDED LIST =======`r" -ScriptName "$ScriptSave"

Write-Verbose -Message "====================================== CHECKING DISTRO MEMBERSHIP FOR ACCOUNTS NOT IN PROVIDED LIST =======`r" -Verbose


    foreach ($CurrentDgMember in $CurrentDGMembers)
    {
    $Error.Clear()

        if ($ValRecipients -notcontains $CurrentDgMember.DistinguishedName)
        {
            $CurrentDgMemberInfo = Get-O365Recipient -Identity $CurrentDgMember.DistinguishedName
            Get-PoshRunLog -Message "[NOT IN LIST]: $($CurrentDgMemberInfo.PrimarySmtpAddress) is an existing member of $DistroGroup, but is not in the provided list and will be removed." -ScriptName "$ScriptSave"
            Write-Verbose "[NOT IN LIST]: $($CurrentDgMemberInfo.PrimarySmtpAddress)" -Verbose
            $RemRecipients += $CurrentDgMember.DistinguishedName
        }

        if ($ValRecipients -contains $CurrentDgMember.DistinguishedName)
        {
            $CurrentDgMemberInfo = Get-O365Recipient -Identity $CurrentDgMember.DistinguishedName
            Write-Verbose -Message "[OBJECT IN LIST]: $($CurrentDgMemberInfo.PrimarySmtpAddress)" -Verbose
            Get-PoshRunLog -Message "[OBJECT IN LIST]: $($CurrentDgMemberInfo.PrimarySmtpAddress) is an existing member of $DistroGroup and is included in the provided list and will not be removed." -ScriptName "$ScriptSave"
            
            
        }

        Write-Progress -Id 1 -Activity 'Checking if object is in list provided' -Status "[$($CurrentDgMemberInfo.PrimarySmtpAddress)]: $CheckRemove of $($CurrentDgMembers.count)" -PercentComplete ($($CheckRemove/$CurrentDgMembers.count)*100)
        $CheckRemove++
    }

Get-PoshRunLog -Message "============= END CHECKING DISTRO MEMBERSHIP FOR ACCOUNTS NOT IN PROVIDED LIST =========================`r" -ScriptName "$ScriptSave"

Get-PoshRunLog -Message "=================================== CHECKING FOR NEW MEMBERS ===========================================`r" -ScriptName "$ScriptSave"

Write-Verbose -Message "====================================== CHECKING FOR NEW MEMBERS ===========================================`r" -Verbose


    foreach ($ValRecipient in $ValRecipients)
    {
        if ($CurrentDgMembers.DistinguishedName -notcontains $ValRecipient)
        {
            $ValRecipientInfo = Get-O365Recipient -Identity $ValRecipient            
            Write-Verbose "[NEW MEMBER FOUND]: $($ValRecipientInfo.PrimarySmtpAddress)" -Verbose
            Get-PoshRunLog -Message "[NEW MEMBER FOUND]:$($ValRecipientInfo.PrimarySmtpAddress) is not a current member and will be added to $DistroGroup" -ScriptName "$ScriptSave"
            $NewRecipients += $ValRecipient

        }
        if ($CurrentDgMembers.DistinguishedName -contains $ValRecipient)
        {

            $ValRecipientInfo = Get-O365Recipient -Identity $ValRecipient
            Get-O365Recipient -Identity $ValRecipient | Select-Object Name,Alias,PrimarySmtpAddress,RecipientType,RecipientTypeDetails | Export-Csv "$SavePath\$DistroGroup.EXISTING_MembersAsOf$unqName.csv" -NoTypeInformation -Append            
            Write-Verbose -Message "[EXISTING MEMBER]: $($ValRecipientInfo.PrimarySmtpAddress)" -Verbose
            Get-PoshRunLog -Message "[EXISTING MEMBER]: $($ValRecipientInfo.PrimarySmtpAddress) is a current member of $DistroGroup" -ScriptName "$ScriptSave"
        }

        Write-Progress -Id 2 -Activity 'Checking for new members' -Status "[$($ValRecipientInfo.PrimarySmtpAddress)]: $CheckNew of $($ValRecipients.count)" -PercentComplete ($($CheckNew/$ValRecipients.count)*100)
        $CheckNew++

    }

Get-PoshRunLog -Message "=================================== END CHECKING FOR NEW MEMBERS =======================================`r" -ScriptName "$ScriptSave"

Get-PoshRunLog -Message "=================================== DATA REVIEW SUMMARY ================================================`r" -ScriptName "$ScriptSave"

Write-Verbose -Message "====================================== DATA REVIEW SUMMARY ================================================`r" -Verbose

Write-Verbose -Message "Total of object in provided list = $($Data.count)" -Verbose

Get-PoshRunLog -Message "Total of object in provided list = $($Data.count)" -ScriptName "$ScriptSave"

Write-Verbose -Message "Total of current $DistroGroup members = $($CurrentDgMembers.count)" -Verbose

Get-PoshRunLog -Message "Total of current $DistroGroup members = $($CurrentDgMembers.count)" -ScriptName "$ScriptSave"

Write-Verbose "Total of VALID recipients = $($ValRecipients.count)" -Verbose

Get-PoshRunLog -Message "Total of VALID recipients = $($ValRecipients.count)" -ScriptName "$ScriptSave"

Write-Verbose "Total of INVALID recipients = $($InValRecipients.count)" -Verbose

Get-PoshRunLog -Message "Total of INVALID recipients = $($InValRecipients.count)" -ScriptName "$ScriptSave"

Write-Verbose "Total of NEW MEMBERS to be added =  $($NewRecipients.count)" -Verbose

Get-PoshRunLog -Message "Total of NEW MEMBERS to be added =  $($NewRecipients.count)" -ScriptName "$ScriptSave"

Write-Verbose "Total of recipients to be removed =  $($RemRecipients.count)" -Verbose

Get-PoshRunLog -Message "Total of recipients to be removed =  $($RemRecipients.count)" -ScriptName "$ScriptSave"

Get-PoshRunLog -Message "=================================== END DATA REVIEW SUMMARY ============================================`r" -ScriptName "$ScriptSave"

Get-PoshRunLog -Message "=================================== ADDING NEW MEMBERS =================================================`r" -ScriptName "$ScriptSave"

Write-Verbose -Message "====================================== ADDING NEW MEMBERS =================================================`r" -Verbose

    if ($NewRecipients)
    {

        foreach ($NewRecipient in $NewRecipients)
        {

        $Error.Clear()

            if ($Force -or $PSCmdlet.ShouldProcess($NewRecipient, "Adding as memberof $DistroGroup"))
            {
               try
                {
                    $NewRecipientInfo = Get-O365Recipient -Identity $NewRecipient
                    Write-Verbose -Message "++++++++++ [ADDED] $($NewRecipientInfo.PrimarySmtpAddress)" -Verbose                    
                    Get-O365Recipient -Identity $NewRecipient | Select-Object Name,Alias,PrimarySmtpAddress,RecipientType,RecipientTypeDetails | Export-Csv "T:\Powershell\Results\$DistroGroup.SUCCESS_AddNewMembers.$unqName.csv" -NoTypeInformation -Append
                    Add-O365DistributionGroupMember -Identity $DistroGroup.PrimarySmtpAddress -Member $NewRecipient -BypassSecurityGroupManagerCheck -Confirm: $false -ErrorAction Stop
                    Get-PoshRunLog -Message "+++++++ [ADDED]: $($NewRecipientInfo.PrimarySmtpAddress) as a memberof $DistroGroup" -ScriptName "$ScriptSave"

            
                }
                catch
                {

                    $NewRecipientInfo = Get-O365Recipient -Identity $NewRecipient
                    Write-Warning "[ERROR] $($NewRecipientInfo.PrimarySmtpAddress) Please review log for details"            
                    $Errorflag = $Error.exception.message.ToUpper()      
                    #$flag += $Errorflag
                    Get-O365Recipient -Identity $NewRecipient | Select-Object @{l='Error';e={"$Errorflag"}},Name,Alias,PrimarySmtpAddress,RecipientType,RecipientTypeDetails | Export-Csv "T:\Powershell\Results\$DistroGroup.ERRORS_AddNewMembers.$unqName.csv" -NoTypeInformation -Append
                    Get-PoshRunLog -Message "[ERROR (See exported csv error file)]: $($NewRecipientInfo.PrimarySmtpAddress)" -ScriptName "$ScriptSave"

                }

            }              

                Write-Progress -Id 2 -Activity "Adding object as new member" -Status "[$($NewRecipientInfo.PrimarySmtpAddress)]: $AddNew of $($NewRecipients.count)" -PercentComplete ($($AddNew/$NewRecipients.count)*100)
                $AddNew++
        }
    }
    else
    {
        Write-Verbose -Message "No new members found" -Verbose
        Get-PoshRunLog -Message "No new members found" -ScriptName "$ScriptSave"
    }

Get-PoshRunLog -Message "=================================== END ADDING NEW MEMBERS =============================================`r" -ScriptName "$ScriptSave"

Get-PoshRunLog -Message "=================================== REMOVING INVALID MEMBERS ===========================================`r" -ScriptName "$ScriptSave"

Write-Verbose -Message "====================================== REMOVING INVALID MEMBERS ===========================================`r" -Verbose
    if ($RemRecipients)
    {

        foreach ($RemRecipient in $RemRecipients)
        {
        $Error.Clear()

            if ($Force -or $PSCmdlet.ShouldProcess($RemRecipient, "Removing as memberof $DistroGroup"))
            {
                try
                {
                    $RemRecipientInfo = Get-O365Recipient -Identity $RemRecipient
                    Write-Verbose -Message "[REMOVING]: $($RemRecipientInfo.PrimarySmtpAddress)" -Verbose
                    Get-O365Recipient -Identity $RemRecipient | Select-Object Name,Alias,PrimarySmtpAddress,RecipientType,RecipientTypeDetails | Export-Csv "T:\Powershell\Results\$DistroGroup.SUCCESS_RemoveMembers.$unqName.csv" -NoTypeInformation -Append        
                    Remove-O365DistributionGroupMember -Identity $DistroGroup.PrimarySmtpAddress -Member $RemRecipient -BypassSecurityGroupManagerCheck -Confirm: $false -ErrorAction Stop
                    Get-PoshRunLog -Message "------- [REMOVED] $($RemRecipientInfo.PrimarySmtpAddress) from $DistroGroup" -ScriptName "$ScriptSave"
                }
                catch
                {
                    $RemRecipientInfo = Get-O365Recipient -Identity $RemRecipient
                    Write-Warning "[ERROR ENCOUNTERED]: $($RemRecipientInfo.PrimarySmtpAddress). Please review log for details"            
                    $Errorflag = $Error.exception.message.ToUpper()          
                    #$flag += $Errorflag
                    Get-O365Recipient -Identity $RemRecipient | Select-Object @{l='Error';e={"$Errorflag"}},Name,Alias,PrimarySmtpAddress,RecipientType,RecipientTypeDetails | Export-Csv "T:\Powershell\Results\$DistroGroup.ERRORS_RemoveMembers.$unqName.csv" -NoTypeInformation -Append            
                    Get-PoshRunLog -Message "[ERROR (See exported csv error file)] $($RemRecipientInfo.PrimarySmtpAddress)" -ScriptName "$ScriptSave"
        
                }
            }

                Write-Progress -Id 2 -Activity "Removing object" -Status "[$($RemRecipientInfo.PrimarySmtpAddress)]: $DelMem of $($RemRecipients.count)" -PercentComplete ($($DelMem/$RemRecipients.count)*100)
                $DelMem++
        }
    }
    else
    {
        Write-Verbose -Message "No members need to be removed" -Verbose
        Get-PoshRunLog -Message "No members need to be removed`r" -ScriptName "$ScriptSave"
    }

Get-PoshRunLog -Message "=================================== END REMOVING INVALID MEMBERS =======================================`r" -ScriptName "$ScriptSave"

}


end {


Write-Verbose -Message "===================================== GETTING GROUP MEMBERSHIP & EXPORT TO CSV ============================`r" -Verbose

Get-PoshRunLog -Message "================================== GETTING GROUP MEMBERSHIP & EXPORT TO CSV ============================`r" -ScriptName "$ScriptSave"

$Results = Get-O365DistributionGroupMember -Identity $DistroGroup.Name | Select-Object Name,PrimarySmtpAddress | Out-String

$Results | Select-Object Name,Alias,PrimarySmtpAddress,RecipientType,RecipientTypeDetails | Export-Csv "$SavePath\$DistroGroup.MembersPOSTUpdate_$unqName.csv" -NoTypeInformation -Append

Get-PoshRunLog -Message "Group members for $($DistroGroup):`r`n $Results" -ScriptName "$ScriptSave"    

$EndTime = $stopwatch.elapsed.ToString()
    
Write-Verbose -Message "Script ends on $(Get-Date -UFormat %Y%m%d_%H%M%S) and timer ends @ $($EndTime)" -Verbose
    
Get-PoshRunLog -Message "Script ends on $(Get-Date -UFormat %Y%m%d_%H%M%S) and timer ends @ $($EndTime)" -ScriptName "$ScriptSave"

}

