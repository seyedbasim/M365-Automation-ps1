$users = Import-Csv "C:\Users\Basim\Documents\H-ONe\Projects\Talent Fort\users.csv"

Connect-MsolService

$password = "123@password"

foreach($user in $users){
  $DName = $user.First_Name + " " + $user.Last_Name
  #Remove-MsolUser -UserPrincipalName $user.Email_Address -Force
  New-MailUser -Name $DName -ExternalEmailAddress $user.Email_Address -MicrosoftOnlineServicesID $user.Email_Address -Password (ConvertTo-SecureString -String $password -AsPlainText -Force)
  #$user.First_Name + " " + $user.Last_Name + " " + $user.Email_Address
}
New-MsolUser -FirstName -LastName -UserPrincipalName
