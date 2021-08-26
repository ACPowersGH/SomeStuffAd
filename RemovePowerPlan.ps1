<#
.SYNOPSIS
   Script to remove rogue Powerscheme/Plan
.DESCRIPTION
   Script uses multiple functions to locate, replace, and remove rogue Powerscheme/Plan
.PARAMETER <paramName>
   <Description of script parameter>
.EXAMPLE
   <An example of using the script>
#>

Function Test-FolderPath {

$Folder=Test-Path -Path 'c:\_Temp\'

            if ($Folder -eq $false)
            {
                New-Item -Path 'c:\_Temp\' -ItemType 'directory'

                #Once folder is created call Function to check powerschemes/plans

                Get-PowerSchemes

            }
            else {

                exit
            }

}

Function Log-ScriptRun {

param (
            [string]$Message,
            [string]$path="c:\_Temp\REMOVAL_LOG.$env:COMPUTERNAME.$(Get-Date -format yyyyddmm).log"
            )

            $Message | Out-File -Filepath $path -Append
}

# Footer

Function Show-ScriptFooter {

$FooterPurpose='Script Purpose: Locate and Remove Power Scheme Disable_PC_Lock'
$FooterExecutionDetails='Script Execution: Executes at device startup via GPO'
$Team='Responsible Team: IT Collaboration Operations'
$Contact='For all inquiries please contact the Helpdesk'
$Date="Script ran on $((Get-Date).ToString())"
$message = @"

===================================================================================================

		    $FooterPurpose
		    $FooterExecutionDetails
		    $Team

===================================================================================================

		        $Contact

               		$Date

"@

$message

}

# Function to check and/or create folder to store logs

Function Change-ActivePowerScheme {

Log-ScriptRun  "Changing Active Power Scheme to Balanced.....................................`r`n"

Log-ScriptRun	"> Locate additional Power Scheme to set as active"

		$NotActive= Get-WmiObject -Class win32_PowerPlan -Namespace root\cimv2\power -Filter "ElementName = 'Balanced'"
		$NotActiveGuid= $NotActive.InstanceID.split('{')[1]
		$SetActivePowerSchemeGuid= $NotActiveGuid -replace '}',""

Log-ScriptRun "> Power Scheme $($NotActive.ElementName) with Guid $SetActivePowerSchemeGuid is not active"

# Set as active power plan

powercfg /setactive "$SetActivePowerSchemeGuid"

Log-ScriptRun "> Power Scheme $($NotActive.ElementName) with Guid $SetActivePlanGuid has been set to active`r`n"

Log-ScriptRun "Active Power Scheme changed......................................`r`n"

Remove-PowerScheme

}

# Remove Disable_PC_Lock

Function Remove-PowerScheme {

Log-ScriptRun "Removing Power Scheme $($PowerScheme.ElementName)................`r`n"

Log-ScriptRun "> Removing Disable_PC_Lock Power Scheme from $env:COMPUTERNAME"

    $PowerSchemeToRemove= Get-WmiObject -Class win32_PowerPlan -Namespace root\cimv2\power -Filter "ElementName = 'Disable_PC_Lock'"
    $PowerSchemeGuid= $PowerSchemeToRemove.InstanceID.split('{')[1]
    $RemovePowerSchemeGuid= $PowerSchemeGuid -replace '}',""

    #Remove Power Scheme

    powercfg /delete "$RemovePowerSchemeGuid"

Log-ScriptRun "> Power Scheme $($PowerSchemeToRemove.ElementName) with Guid $RemovePowerSchemeGuid has been removed`r`n"

Log-ScriptRun "List of all Power Scheme present on $($env:COMPUTERNAME) after removal of (Disable_PC_Lock) power scheme....`r`n"

$PostPowerScheme = Get-WmiObject -Class win32_PowerPlan -Namespace root\cimv2\power | Select-Object ElementName,InstanceID,IsActive | Out-String

Log-ScriptRun $PostPowerScheme

Log-ScriptRun "Logging Process stopped......................................`r`n"

Log-ScriptRun (Show-ScriptFooter)

}

# Main Functon

Function Get-PowerSchemes {

Log-ScriptRun  "Logging process started........................................`r`n"


Log-ScriptRun "List of all Power Schemes present on $($env:COMPUTERNAME)`r`n"

    $PrePowerScheme = Get-WmiObject -Class win32_PowerPlan -Namespace root\cimv2\power | Select-Object ElementName,InstanceID,IsActive | Out-String

    Log-ScriptRun  $PrePowerScheme

    $PowerScheme=Get-WmiObject -Class win32_PowerPlan -Namespace root\cimv2\power -Filter "ElementName = 'Disable_PC_Lock'"

    if (!($PowerScheme)) {

        Log-ScriptRun "> WARNING MESSAGE: Power Scheme 'Disable_PC_Lock' was not found on $env:COMPUTERNAME`r`n"

        Log-ScriptRun "Script stopped...............`r`n"

        Log-ScriptRun (Show-ScriptFooter)

        exit

    }

    else {

        Switch ($PowerScheme.IsActive) {

            'True' {

                Log-ScriptRun "> $($PowerScheme.ElementName) is the active Scheme and must be changed before continuing`r`n" #| Out-File "c:\$Temp\REMOVAL_LOG.$env:COMPUTERNAME.log" -Append

                Change-ActivePowerScheme

            }

            'False'{

                Log-ScriptRun "> $($PowerScheme.ElementName) is not the active Scheme`r`n> Ok to remove power Scheme`r`n" #| Out-File "c:\$Temp\REMOVAL_LOG.$env:COMPUTERNAME.log" -Append

                Remove-PowerScheme

            }
        }
    }
}

# Call Function to start script execution

Test-FolderPath