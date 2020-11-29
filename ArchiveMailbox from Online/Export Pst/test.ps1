$psmodule = "C:\Users\Basim\Documents\H-ONe\Scripts\Export Pst\export.psm1"
Import-Module $psmodule

$cred = Get-Credential #Example user: user@example.com
$domain = 'honelab.onmicrosoft.com'

# Set how many emails we want to read at a time
$PageSizeNumOfEmails = 10
$OffSet = 0
$PageIndexLimit = 2000
$MailFolder = 'Inbox'
$DeltaTimeStamp = (Get-Date).AddDays(-30) #Go how many days back?

try
{
  $Session = Enter-ExchangeOnlineSession -Credential $Cred -MailDomain $domain
}
catch
{
  $Message = $_.exception.message
  Write-Host -ForegroundColor Yellow $Message
}

$Mails = Get-ExchangeOnlineMailContent -ServiceObject $Session `
                                        -PageSize $PageSizeNumOfEmails `
                                        -Offset $OffSet `
                                        -PageIndexLimit $PageIndexLimit `
                                        -WellKnownFolderName $MailFolder `
                                        -ParseOriginalRecipient `
                                        -OriginalRecipientAddressOnly `
                                        -MailFromDate $DeltaTimeStamp | Where-Object {$_.DateTimeReceived -gt $DeltaTimeStamp} | select *