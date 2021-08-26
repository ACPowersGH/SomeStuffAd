# Run AuditLog against specific searches

[CmdletBinding()]
param
(
[Parameter(Mandatory=$false)]
[ValidateNotNullOrEmpty()]$StartDate,
[Parameter(Mandatory=$false)]
[ValidateNotNullOrEmpty()]$EndDate,
[Parameter(Mandatory=$true)]
[ValidateNotNullOrEmpty()]$Commands,
[Parameter(Mandatory=$true)]
[ValidateNotNullOrEmpty()]$Params,
[Parameter(Mandatory=$true)]
[ValidateNotNullOrEmpty()]$TargetObject
)

Search-AdminAuditLog -StartDate $StartDate -EndDate $EndDate -Cmdlets $Commands -Parameters $Params -ObjectIds $TargetObject | select IsValid, ObjectModified, Caller, CmdletName, @{Label="Run Date";e={($_.Rundate).ToUniversalTime()}}, Succeeded, @{Label="Executed Command";e={$_ | Select-Object -ExpandProperty CmdletParameters}} | Export-Csv "T:\Powershell\CSV\Results\$(Get-Date -UFormat %Y%d%m%_T_%H%M%S).AuditLog.csv" -NoTypeInformation

<# SAMPLE RUN

PS T:\Powershell\Scripts> .\Audit\SearchAdminAudtiLog.ps1

cmdlet SearchAdminAudtiLog.ps1 at command pipeline position 1
Supply values for the following parameters:
StartDate: 12/16/2019
EndDate: 12/17/2019
Commands: Set-Mailbox
Params: LitigationHoldEnabled
TargetObject: Morrison.Don.O
PS T:\Powershell\Scripts>

#>