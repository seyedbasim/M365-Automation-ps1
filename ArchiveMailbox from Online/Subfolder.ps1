$fv = New-Object Microsoft.Exchange.WebServices.Data.Folderview(1000)

$list = $exchservice.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot,$fv)

$list | select DisplayName, Id