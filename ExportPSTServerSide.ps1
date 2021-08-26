# Run from LIGT Exchange Server rnop-exci01

Function Log-ScriptRun {

param (
            [string]$Message,
            [string]$path="\\$env:COMPUTERNAME\E$\ADHOC_PST\Export_$(Get-Date -format yyyyddmm).log"
            )

            $Message | Out-File -Filepath $path -Append
}


function cus_Export-PST {

    [CmdletBinding()]
        param (
            
            [Parameter(Mandatory=$true,HelpMessage="Enter CSV file name")]$CSVFileName
        )

    $data = Import-Csv "\\$env:COMPUTERNAME\E$\ADHOC_PST\$CSVFileName.csv"

    #$domain = list domains

    $log=@()

    $valid=@()
    $unqName = Get-Date -UFormat %Y%m%d_%H%M%S

    #$ErrorActionPreference = 'SilentlyContinue'

    foreach ($data in $data.Email) {

    $Error.Clear()

    Log-ScriptRun "Verify mailbox $($data)"

    Write-Host "Checking valid mailbox" -BackgroundColor Cyan -ForegroundColor Black

        Get-Mailbox -Identity $data

            if ($Error) {

                $log += "ERROR < $($data.toUpper()) >","$Error" -join "::"
                $log > "\\$env:COMPUTERNAME\E$\ADHOC_PST\$unqName.NotValid.txt"

            } else {
            
                $valid  += $data 
            
                }

    }

    #Write-Host "Getting DisplayName" -BackgroundColor Magenta -ForegroundColor Black

    Log-ScriptRun "Checking number of valid accounts;Valid Recipients = $($valid.count)" 

    Write-Host "There is a total of $($valid.count) valid recipients" -BackgroundColor cyan -ForegroundColor black

    sleep -Milliseconds 100

    #$progress=0

    $StartTime = Get-Date

    foreach ($object in $valid) {

    $fileloc = "\\$env:COMPUTERNAME\E$\ADHOC_PST"
    $filename  = "PST_" + $object + ".$unqName"
    $filepath = $fileloc + "\" + $filename
    $ExportStats = $object + "\" + $filename

        #Get-Mailbox -Identity $result | select DisplayName

        New-MailboxExportRequest -Mailbox $object -Name $filename -BadItemLimit 9999 -AcceptLargeDataLoss -FilePath "$filepath.pst" #-WhatIf
    
        <#Get-MailboxStatistics -Identity $object | select displayname,`
        @{label='TotalMailboxSize (MB)';e={$_.TotalItemSize.value.ToMB()}},`
        @{label='USERID';e={Get-Mailbox $object| select -ExpandProperty samaccountname}},`
        @{label='Email';e={Get-Mailbox -Identity $object | select -ExpandProperty primarysmtpaddress}},servername,database,itemcount | Export-Csv \\$env:COMPUTERNAME\d$\Temp\$unqName.MailboxStats.csv -NoTypeInformation -Append #>

        Log-ScriptRun "Checking export status for $($object)"

        Do 
        {

        

	    Write-Verbose "Checking status $($object)" -Verbose            
    
            $checkStatus = ((Get-MailboxExportRequestStatistics -Identity $ExportStats -ErrorAction silentlycontinue).StatusDetail -eq "Completed")

            sleep -s 180

        } Until ($checkStatus -eq $true) 
                
                Log-ScriptRun "Export for $($object) complete. Moving to next object"
                
                Write-Host "Export for $($object) complete. Moving to next object" -BackgroundColor Cyan -ForegroundColor Black
                
                continue

        #$progress++

        #$SecondsElapsed = ((Get-Date) - $StartTime).TotalSeconds
        #$SecondsRemaining = ($SecondsElapsed * ($valid.count – $progress))/$progress

        #Write-Progress -Activity "Creating Mailbox Export" -Status "Exporting $($object) :: $progress of $($valid.count)" -CurrentOperation "Exporting" -PercentComplete (($progress/$valid.count)*100)
        
        #Write-Progress -Activity "Mailbox Export requests" -Status "Creating Export for $object : ($progress of $($valid.count))" -CurrentOperation "$("{0:N2}" -f (($progress/$($valid.Count)) * 100),2)% Complete" -PercentComplete (($progress/$valid.count)*100)  -SecondsRemaining $SecondsRemaining


        #Start-Sleep -Seconds 3000
    }

}#end function

#Call function
cus_Export-PST
