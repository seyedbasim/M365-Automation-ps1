#Read-Host -AsSecureString | ConvertFrom-SecureString | Out-File EXOpassword.txt

$password=get-content EXOpassword.txt | ConvertTo-SecureString

$username = "@.onmicrosoft.com"

$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password

Connect-ExchangeOnline -Credential $Credential

Get-Mailbox -ResultSize Unlimited -Filter "RecipientTypeDetails -eq 'UserMailbox'" | Where-Object LitigationHoldEnabled -eq $false | Set-Mailbox -LitigationHoldEnabled $true -LitigationHoldDuration 5475
