function Search-PoshMailbox {

     [CmdletBinding()]
     param(
         [Parameter(Mandatory=$true,
                    HelpMessage='Enter email address to search againsts',
                    ValueFromPipeline=$true)]
                    [string[]]$EmailAddress,
                    [string]$SearchQuerySubject,
                    [string]$TargetMailboxUserID,
                    [string]$TargetFolder
     )



     begin {}
     process {
         foreach ($email in $EmailAddress){


             try {

                 Get-EXOMailbox -Identity $email

             } catch {

                 Write-Warning "$($_.CategoryInfo) on $domain"

             } finally {

                 Search-EXOMailbox $email -SearchQuery $SearchQuerySubject -TargetMailbox "$TargetMailboxUserID" -TargetFolder "$TargetFolder" -LogOnly -LogLevel full

             }
         }

     }

 }