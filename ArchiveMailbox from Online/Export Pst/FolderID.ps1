Import-Module -Name “C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll”
$Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials(“basim@honelab.onmicrosoft.com”,”pass@123”)
$exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$exchService.Credentials = $Credentials
$exchService.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
$Archive = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchservice,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot)
$Archive | fl
$Copyto = New-Object Microsoft.Exchange.WebServices.Data.FolderID([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveRoot,"basim@honelab.onmicrosoft.com")
$copyto | fl
$str = $copyto.ToString()
$str | fl
$Archive.copy([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)

#--------------coversion----------------------
$encoding= [System.Text.Encoding]::GetEncoding("us-ascii")
$nibbler= $encoding.GetBytes("0123456789ABCDEF");
$folderIdBytes = [Convert]::FromBase64String($Archive.Id);
$indexIdBytes = New-Object byte[] 48;
$indexIdIdx=0;
$folderIdBytes | select -skip 23 -First 24 | %{$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -shr 4];$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -band 0xF]}
$folderQuery = "folderid:$($encoding.GetString($indexIdBytes))";

$folderQuery
#-----------------------------------------------

$Mailboxname = "basim@honelab.onmicrosoft.com"

$aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateID     
$aiItem.Mailbox = $MailboxName        
$aiItem.UniqueId = $Archive.Id   
$aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId       
$convertedId = $exchService.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::StoreId)   
$convertedId.UniqueId  



