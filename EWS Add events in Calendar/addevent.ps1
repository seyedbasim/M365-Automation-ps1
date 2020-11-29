$users = Import-Csv "C:\Users\Basim\Documents\H-ONe\LOLC\Calendar\users.csv"
 
$s = 0
$e = 0
$c = 0 
$eventfile = "C:\Users\Basim\Documents\H-ONe\LOLC\Calendar\Holidays-2018.csv"
$psmodule = "C:\Users\Basim\Documents\H-ONe\LOLC\Calendar\NewOSCEXOAppointment\NewOSCEXOAppointment.psm1"
$events = Import-Csv $eventfile -Encoding UTF8
Import-Module $psmodule
$users | Sort-Object UPN
#Write-EventLog -LogName Application -Source ADDPERMISSION -EventId 4 -EntryType Information -Message "Started Adding user permission"
$cred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber
Connect-OSCEXOWebService -Credential $cred
foreach($user in $users){
        
    $c++

    if( $c -eq 50){
        $c = 0
    }

    try{
        
        $events  | % {New-OSCEXOAppointment -Identity $user.UPN -Subject $_.Subject -StartDate $_.StartDate -AllDayEvent}
        #Write-EventLog -LogName Application -Source SETCALEVENT -EventId 3 -EntryType Information -Message "Event Update completed for the user: $($user.UPN)"
        if($? -eq $false){
        #Write-Host "True"
        throw $error[0].Exception
        }
        $user | Export-Csv "C:\Users\Basim\Documents\H-ONe\LOLC\Calendar\eventAddSuccessUsers.csv" -NoTypeInformation -Append
        $s++

    }
    catch{
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        $errorentry = "Exception occurred on user $($user.UPN). Message ==> $ErrorMessage"
        #Write-EventLog -LogName Application -Source SETCALEVENT -EventId 1003 -EntryType Error -Message $errorentry
        $user | Export-Csv "C:\Users\Basim\Documents\H-ONe\LOLC\Calendar\eventAddErrorUsers.csv" -NoTypeInformation -Append
        $e++
        continue

    }
}
Remove-PSSession $Session