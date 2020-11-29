$users = import-csv -Path "C:\Users\Basim\Documents\H-ONe\LOLC\Calendar\users.csv"
$users | Sort-Object UPN
#Write-EventLog -LogName Application -Source ADDPERMISSION -EventId 4 -EntryType Information -Message "Started Adding user permission"
#------------------------------------------------------------------
$eventfile = "C:\Users\Basim\Documents\H-ONe\LOLC\Calendar\Holidays-2018.csv"
$psmodule = "C:\Users\Basim\Documents\H-ONe\LOLC\Calendar\NewOSCEXOAppointment\NewOSCEXOAppointment.psm1"
$events = Import-Csv $eventfile -Encoding UTF8
Import-Module $psmodule
#-------------------------------------------------------------------
$cred = Get-Credential
foreach($user in $users){
#----------------------Adding Delegation-------------------------------------
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
	Import-PSSession $Session -AllowClobber
    try{
        Add-MailboxFolderPermission -Identity "$($user.UPN):\calendar" -user basim@honelab.onmicrosoft.com -AccessRights Editor
        if($? -eq $false){
        #Write-Host "True"
        throw $Error[0].Exception
        }           
        #Write-EventLog -LogName Application -Source ADDPERMISSION -EventId 4 -EntryType Information -Message "Adding Permisson completed on : $($user.UPN)"
        $user | Export-Csv "C:\Users\Basim\Documents\H-ONe\LOLC\Calendar\permissionSuccessUsers.csv" -Append
    }
    catch{
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        $errorentry = "Exception occurred on user $($user.UPN). Message ==> $ErrorMessage"
        #Write-EventLog -LogName Application -Source ADDPERMISSION -EventId 1004 -EntryType Error -Message $errorentry
        $user | Export-Csv "C:\Users\Basim\Documents\H-ONe\LOLC\Calendar\permissionErrorUsers.csv" -Append
        continue
    }
#---------------------Adding Calendar Events-------------------------------------
	Connect-OSCEXOWebService -Credential $cred
	try{
        $events  | % {New-OSCEXOAppointment -Identity $user.UPN -Subject $_.Subject -StartDate $_.StartDate -AllDayEvent}
        #Write-EventLog -LogName Application -Source SETCALEVENT -EventId 3 -EntryType Information -Message "Event Update completed for the user: $($user.UPN)"
        if($? -eq $false){
        #Write-Host "True"
        throw $error[0].Exception
        }
        $user | Export-Csv "C:\Users\Basim\Documents\H-ONe\LOLC\Calendar\eventAddSuccessUsers.csv" -NoTypeInformation -Append
    }
    catch{
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        $errorentry = "Exception occurred on user $($user.UPN). Message ==> $ErrorMessage"
        #Write-EventLog -LogName Application -Source SETCALEVENT -EventId 1003 -EntryType Error -Message $errorentry
        $user | Export-Csv "C:\Users\Basim\Documents\H-ONe\LOLC\Calendar\eventAddErrorUsers.csv" -NoTypeInformation -Append
        continue
    }
#-------------------Removing Delegation--------------------------------------------------------
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
	Import-PSSession $Session -AllowClobber
	try{
        Remove-MailboxFolderPermission -Identity "$($user.UPN):\calendar" -user basim@honelab.onmicrosoft.com -Confirm:$false
        if($? -eq $false){
            #Write-Host "True"
            throw $Error[0].Exception
        } 
        #Write-EventLog -LogName Application -Source REMOVEPERMISSION -EventId 4 -EntryType Information -Message "Adding Permisson completed on : $($user.UPN)"
        $user | Export-Csv "C:\Users\Basim\Documents\H-ONe\LOLC\Calendar\permissionRemoveSuccessUsers.csv" -Append
    }
    catch{
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        $errorentry = "Exception occurred on user $($user.UPN). Message ==> $ErrorMessage"
        #Write-EventLog -LogName Application -Source REMOVEPERMISSION -EventId 1004 -EntryType Error -Message $errorentry
        $user | Export-Csv "C:\Users\Basim\Documents\H-ONe\LOLC\Calendar\permissionRemoveErrorUsers.csv" -Append
        continue
    }	
}
Remove-PSSession $Session