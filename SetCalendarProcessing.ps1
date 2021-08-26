[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory=$true,
    ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    $RoomEmailAddress,$SecurityGroupDelegate
    )

Begin{}

Process{
    Get-CalendarProcessing -Identity $RoomEmailAddress | Select-Object auto*,*policy,*delegates | Format-Table -AutoSize
    
    Write-Verbose 'Review current calendar configuration' -Verbose

    pause

    $UserInput = Read-Host "Would you like to set calendar processing to manual approval? (Y/N)"

    if ($UserInput -like 'y'){
        Write-Verbose "Setting manual approval/decline configuration for $($RoomEmailAddress)" -Verbose
        Set-CalendarProcessing -Identity $RoomEmailAddress -AllBookInPolicy $false -AllRequestInPolicy $true -BookInPolicy $SecurityGroupDelegate
    }else {
        Write-Verbose "Setting automatic approval/decline configuration for $($RoomEmailAddress)" -Verbose
        Set-CalendarProcessing -Identity $RoomEmailAddress -AllBookInPolicy $true -AllRequestInPolicy $false
    }

    Write-Verbose 'Waiting for replication of permission update' -Verbose

    Start-Sleep -s 10

    Get-CalendarProcessing -Identity $RoomEmailAddress | Select-Object auto*,*policy,*delegates | Format-Table -AutoSize
}

End{}