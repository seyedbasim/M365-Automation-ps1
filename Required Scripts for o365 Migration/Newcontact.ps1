$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking

$password = "123@password"

$s = $password.Split("@")

$s[0]

foreach($user in $users){
  #$DName = $user.First_Name + " " + $user.Last_Name
  #New-MailUser -Name $DName -ExternalEmailAddress $user.Email_Address -MicrosoftOnlineServicesID $user.Email_Address -Password (ConvertTo-SecureString -String $password -AsPlainText -Force)
  #New-MsolUser -FirstName $user.First_Name -LastName $user.Last_Name -UserPrincipalName $user.Email_Address -DisplayName $DName
  #$user.First_Name + " " + $user.Last_Name + " " + $user.Email_Address
  $Email = $user.Email_Address.split("@")
  $EmailA = $Email[0] + "@o365.talentfort.com"
  Set-MailUser $user.Email_Address -EmailAddresses @{add=$EmailA}
}

New-MailUser -Name "Chandaka" -ExternalEmailAddress "chandaka@acquest.lk" -MicrosoftOnlineServicesID "chandaka@acquest.lk" -Password (ConvertTo-SecureString -String $password -AsPlainText -Force)

Set-MailUser "chandaka@acquest.lk" -EmailAddresses @{add="chandaka@o365.acquest.lk"}
