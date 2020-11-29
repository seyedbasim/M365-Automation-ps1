Set-ExecutionPolicy RemoteSigned

$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking

$BoardGroup = Import-Csv "C:\Users\Basim\Documents\H-ONe\Projects\AMCHAM\Board Group.csv"
$OBGroup = Import-Csv "C:\Users\Basim\Documents\H-ONe\Projects\AMCHAM\OB Group.csv"
$AMCHAM = Import-Csv "C:\Users\Basim\Documents\H-ONe\Projects\AMCHAM\AMCHAM Membership.csv"

New-DistributionGroup -Name "AMCHAM Membership" -PrimarySmtpAddress "AMCHAM.Membership@amcham.lk" -MemberJoinRestriction "Closed"
foreach ($users in $AMCHAM){
    New-MailContact -Name $users.NAME -ExternalEmailAddress $users.UPN
    Add-DistributionGroupMember -Identity "AMCHAM Membership" -Member $users.UPN
}

New-DistributionGroup -Name "OB Group" -PrimarySmtpAddress "OB.Group@amcham.lk" -MemberJoinRestriction "Closed"
foreach ($users1 in $OBGroup){
    New-MailContact -Name $users1.NAME -ExternalEmailAddress $users1.UPN
    Add-DistributionGroupMember -Identity "OB Group" -Member $users1.UPN
}

New-DistributionGroup -Name "Board Group" -PrimarySmtpAddress "Board.Group@amcham.lk" -MemberJoinRestriction "Closed"
foreach ($users2 in $BoardGroup){
    New-MailContact -Name $users2.NAME -ExternalEmailAddress $users2.UPN
    Add-DistributionGroupMember -Identity "Board Group" -Member $users2.UPN
}
