#Import-Module ActiveDirectory
#$users = Get-ADUser -Filter {(Co -like "Sri Lanka" -and EmailAddress -like "*")} -Properties * -SearchBase "OU=Test,DC=Dag,DC=com" -ResultSetSize $null | Sort-Object UserPrincipalName 
#$results = foreach ($user in $users) {
 #           [pscustomobject]@{
  #          UPN = $user.userprincipalname
   #         }
    #      }  

$users = import-csv -Path "C:\Users\Basim\Documents\H-ONe\LOLC\Calendar\users.csv"
$s = 0
$e = 0
$users | Sort-Object UPN
#Write-EventLog -LogName Application -Source ADDPERMISSION -EventId 4 -EntryType Information -Message "Started Adding user permission"
$cred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
Import-PSSession $Session
foreach($user in $users){
    try{
        Add-MailboxFolderPermission -Identity "$($user.UPN):\calendar" -user basim@honelab.onmicrosoft.com -AccessRights Editor
        if($? -eq $false){
        #Write-Host "True"
        throw $Error[0].Exception
        }           
        #Write-EventLog -LogName Application -Source ADDPERMISSION -EventId 4 -EntryType Information -Message "Adding Permisson completed on : $($user.UPN)"
        $user | Export-Csv "C:\Users\Basim\Documents\H-ONe\LOLC\Calendar\permissionSuccessUsers.csv" -Append
        $s++
    }
    catch{
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        $errorentry = "Exception occurred on user $($user.UPN). Message ==> $ErrorMessage"
        #Write-EventLog -LogName Application -Source ADDPERMISSION -EventId 1004 -EntryType Error -Message $errorentry
        $user | Export-Csv "C:\Users\Basim\Documents\H-ONe\LOLC\Calendar\permissionErrorUsers.csv" -Append
        $e++
        continue
    }
}
Remove-PSSession $Session