# Add multiple DG owners

[CmdletBinding()]
param ([string[]]$Owner,$Group)

$DG = Get-EXODistributionGroup -Identity $Group

$DGManagedBy = $DG.Managedby

$New = @()

foreach ($owner in $owner) {
    
    $AddNew = (Get-EXORecipient -Identity $Owner).DistinguishedName
    $New += $AddNew

    Set-EXODistributionGroup -Identity $Group -ManagedBy $New
}

