$ErrorActionPreference = 'SilentlyContinue'
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking

Add-Content -Path C:\Mailboxdetails.csv  -Value '"Name","Primary_Size","Primary_Item_Count","Archive_Size","Archive_Item_Count"'

$users = Get-Mailbox -resultsize unlimited
$users | ForEach-Object {
 $user = $_.UserPrincipalName
 $Name = $_.Name
 $ArchiveCount = 0
 $Mailbox = Get-MailboxStatistics $user
 $MailboxArchive = Get-MailboxStatistics -archive $user
 $MailboxArchiveFolders = Get-MailboxFolderStatistics -Identity $user -Archive
 $MailboxArchiveFolders | ForEach-Object {
   $ItemCount = $_.ItemsInFolder
   $ArchiveCount = $ItemCount + $ArchiveCount
 }
 if($ArchiveCount -ne 0){
    $ArchiveCount = $ArchiveCount - 1
 }
 $hash = @{
             "Name" = $Name 
             "Primary_Size" = $Mailbox.totalItemsize  
             "Primary_Item_Count" = $Mailbox.ItemCount  
             "Archive_Size" = $MailboxArchive.TotalItemSize 
             "Archive_Item_Count" = $ArchiveCount
             }
$newRow = New-Object PsObject -Property $hash
Export-Csv C:\Mailboxdetails.csv -inputobject $newrow -append -Force
}
$ErrorActionPreference = 'Continue'