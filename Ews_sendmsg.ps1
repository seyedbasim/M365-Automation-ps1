
Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
$Creds = Get-Credential
$Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($Creds)
$exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$exchService.Credentials = $Credentials
$exchService.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
$Message = $null
$Message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($exchService)
$message.Subject = "Dupli";
$message.Body = "EwsTest2";
$message.ToRecipients.Add("basim.moulana@hirdaramani.com");
$message.SendAndSaveCopy();
