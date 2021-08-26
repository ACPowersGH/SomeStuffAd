$DomainServers = (get-adcomputer -LDAPFilter "(&(objectCategory=computer)(operatingSystem=Windows Server*)(!serviceprincipalname=*MSClusterVirtualServer*)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))" -Property name | sort-object Name)


$ValidServer = @()


For($i=0; $i -lt $DomainServers.count; $i++){


    Write-Progress -Activity "checking $($DomainServers[$i].DNSHostName)" -Status "Getting info for $($DomainServers[$i].DNSHostName)" -PercentComplete (($i/$DomainServers.count)*100) 

    $ValidServer += Test-Connection -ComputerName $DomainServers[$i].DNSHostName -Count 1 -ErrorAction:SilentlyContinue

}


For($i=0; $i -lt $ValidServer.count; $i++){


    Write-Progress -Activity "checking $($ValidServer[$i].Address)" -Status "Getting info for $($ValidServer[$i].Address)" -PercentComplete (($i/$ValidServer.count)*100) 

    $CertInfo = Invoke-Command -ComputerName $ValidServer[$i].Address -Command {get-childitem cert:LocalMachine\My -recurse | where-object {$_.NotAfter -gt (get-date)} | select Subject,FriendlyName,Thumbprint,Issuer,NotBefore,NotAfter,@{Name="Expires in (Days)";Expression={($_.NotAfter).subtract([DateTime]::Now).days}}} -ErrorAction:SilentlyContinue

    $CertInfo | Export-Csv C:\Temp\DomainServer-CertInfo.csv -NoTypeInformation -Append

}