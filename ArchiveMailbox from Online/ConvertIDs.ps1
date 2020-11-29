
$id = "AQMkAGY0NjQ0ADUwZS1jNWUwLTQwMTYtYTM5Ny1mNTk2NTE2NzEyMWIALgAAA9taM4BtCqpLtuY%2BfELEFMsBANw%2FJbVTfWZBuqgCOO5JlHsAAAIBRwAAAA"

$Mailboxname = "basim@honelab.onmicrosoft.com"

$aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateID     
$aiItem.Mailbox = $MailboxName        
$aiItem.UniqueId = $Id   
$aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::OwaId       
$convertedId = $exchService.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::StoreId)   
$convertedId.UniqueId  