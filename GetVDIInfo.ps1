# Script logging

Function Get-AcScriptRunLog {

param (
            [string]$Message,
            [string]$ScriptName,
            [string]$path="\\$env:COMPUTERNAME\c$\temp\$ScriptName.$(Get-Date -Format MMddyyyyHHmm).log"
            )

            $Message | Out-File -Filepath $path -Append
}

function Remove-Files {

param ($path) 

Remove-Item -Path $path

}


function Get-AcBrokerMachine {

    # Gather all VDI machine data

    $AllVDIs = Get-BrokerMachine * | ?{$_.MachineName -like "igtmaster\crnov-qc*"}
    
    # HTML formating for email

    $format = @'
<html>
<head>
	<title></title>
</head>
<body>
<h1>VDI Status Report</h1>

<p>Please see below for VDIs reporting as <strong><em><span style="background-color:#FF0000;">&quot;Unregistered&quot;</span></em></strong></p>
</body>
</html>
'@

    # Loop through each data and get the information

    Get-AcScriptRunLog -ScriptName 'cus_Get-VDIInfo' -Message "Running foreach loop"

    foreach ($VDI in $AllVDIs) {

        
        Get-AcScriptRunLog -ScriptName 'cus_Get-VDIInfo' -Message "In Switch method for $($VDI.MachineName)"
        
        # Switch

        Switch ($VDI.RegistrationState) {
  
              'Registered'{

                Get-BrokerMachine -MachineName $VDI.MachineName | Select `
                @{Name="VDI_Name";Expression={$VDI.MachineName}},`
                @{Name="AssignedUsers";Expression={$VDI.AssociatedUserNames -join ";"}},`
                @{Name="Registration";Expression={$VDI.RegistrationState}} | Export-Csv "\\$env:COMPUTERNAME\c$\temp\$(Get-Date -Format MMMddyyy).Registered_VDIs.csv" -NoTypeInformation -Append -NoClobber -Force

                $RegisteredData = "\\$env:COMPUTERNAME\c$\temp\$(Get-Date -Format MMMddyyyy).Registered_VDIs.csv"
            
              } 

              'Unregistered' {
            
                Get-BrokerMachine -MachineName $VDI.MachineName | Select `
                @{Name="VDI_Name";Expression={$VDI.MachineName}},`
                @{Name="AssignedUsers";Expression={$VDI.AssociatedUserNames -join ";"}},`
                @{Name="Registration";Expression={$VDI.RegistrationState}} | Export-Csv "\\$env:COMPUTERNAME\c$\temp\$(Get-Date -Format MMMddyyyy).UnRegistered_VDIs.csv" -NoTypeInformation -Append -NoClobber -Force

                $UnRegisteredData = "\\$env:COMPUTERNAME\c$\temp\$(Get-Date -Format MMMddyyyy).UnRegistered_VDIs.csv"

              }

        } # End Switch

    } # End foreach
    
    Get-AcScriptRunLog -ScriptName 'cus_Get-VDIInfo' -Message "Prepping for send-mail"
    
    # Send email

        Get-AcScriptRunLog -ScriptName 'cus_Get-VDIInfo' -Message "Importing CSV Data"

        # Import CSV data for Unregistered VDIs and assign to variable
            
        $Alerts = Import-Csv $UnRegisteredData

        Get-AcScriptRunLog -ScriptName 'cus_Get-VDIInfo' -Message "Converting Unregistered data to html to add to email body"

        # Convert CSV data into html

        $UnregisteredAlerts = $Alerts | ConvertTo-Html -As list -Fragment -PreContent "<h4>Report ran on $(Get-Date -Format g) PST......</h4>"

        Get-AcScriptRunLog -ScriptName 'cus_Get-VDIInfo' -Message "Creating HTML file"

        # Create html file from all html converted data

        ConvertTo-Html -Head $format -PreContent "<h1>Unregistered VDIs</h1>" -PostContent $UnregisteredAlerts  > "\\$env:COMPUTERNAME\c$\temp\$(Get-Date -Format MMMddyyyy).UnRegistered_VDIs.htm"
        
        Get-AcScriptRunLog -ScriptName 'cus_Get-VDIInfo' -Message "Creating Mailbody based on html file"
                
        # Create mailbody for email message 

        $MailBody = Get-Content "\\$env:COMPUTERNAME\c$\temp\$(Get-Date -Format MMMddyyyy).UnRegistered_VDIs.htm"
        
        Get-AcScriptRunLog -ScriptName 'cus_Get-VDIInfo' -Message "Sending email" 

        Send-MailMessage -To alexander.cuenca@igt.com -From rnop-xeni04@igt.com -Attachments $RegisteredData,$UnRegisteredData -Subject "VDI machine status for report ran on $(Get-date -Format g) PST" -Body "$MailBody" -BodyAsHtml -SmtpServer smtp.igt.com

        Sleep -Milliseconds 100
        
    # Remove ALL files

    Get-AcScriptRunLog -ScriptName 'cus_Get-VDIInfo' -Message "Removed Registered CSV file"

    Remove-Files -path "$UnRegisteredData"

    sleep -Milliseconds 100

    Get-AcScriptRunLog -ScriptName 'cus_Get-VDIInfo' -Message "Removed UnRegistered CSV file"

    Remove-Files -path "$RegisteredData"
    
    sleep -Milliseconds 100

    Get-AcScriptRunLog -ScriptName 'cus_Get-VDIInfo' -Message "Removed HTML file"

    Remove-Files -path "\\$env:COMPUTERNAME\c$\temp\$(Get-Date -Format MMMddyyyy).UnRegistered_VDIs.htm"
    
    Get-AcScriptRunLog -ScriptName 'cus_Get-VDIInfo' -Message "Script complete $(Get-Date)"

} # End function