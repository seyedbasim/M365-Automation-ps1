$convertedId.UniqueId 

$encoding= [System.Text.Encoding]::GetEncoding("us-ascii")
$nibbler= $encoding.GetBytes("0123456789ABCDEF");
$folderIdBytes = [Convert]::FromBase64String("LgAAAADbWjOAbQqqS7bmPnxCxBTLAQDcPyW1U31mQbqoAjjuSZR7AAAAAAFHAAAB");
$indexIdBytes = New-Object byte[] 48;
$indexIdIdx=0;
$folderIdBytes | select -skip 23 -First 24 | %{$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -shr 4];$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -band 0xF]}
$folderQuery = "folderid:$($encoding.GetString($indexIdBytes))";

$folderQuery