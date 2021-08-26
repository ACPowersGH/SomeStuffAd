function Get-ServerCertificates {
    
    [CmdletBinding()]
        param()
    $Error.Clear()
    $Hostname = (Get-Content "E:\#Powershell\_Data\_LyncServerReport\LyncServerList.txt")
    try {

    $results = Invoke-Command {Get-ChildItem -Path cert:\localmachine\my -Recurse | select PSComputerName,Subject,Issuer,FriendlyName,NotAfter} -ComputerName $HostName
    }catch {

    $Error.errordetails.message | Out-File $env:COMPUTERNAME\d$\Error.txt

    } finally {
    $results | Select PSComputerName,Subject,FriendlyName,NotAfter | Export-Csv -Path E:\#Powershell\_Results\_LyncServerReport\LyncServer2010.$(Get-date -Format "mmddyyyy").csv -NoTypeInformation
    }
}
