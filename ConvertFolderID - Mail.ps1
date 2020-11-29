$Mailbox = "basim@honelab.onmicrosoft.com"

$aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId        
$aiItem.Mailbox = $MailboxName        
$aiItem.UniqueId = $EwsId    
$aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId       
$convertedId = $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId)   
return $convertedId.UniqueId  
