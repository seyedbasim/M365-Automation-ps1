Import-Module -Name “C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll”
$Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials(“basim@honelab.onmicrosoft.com”,”pass@123”)
$exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$exchService.Credentials = $Credentials
$exchService.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
$Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchservice,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
$Calendar | fl