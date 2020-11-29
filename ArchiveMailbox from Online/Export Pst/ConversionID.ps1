#---------------------------------Inputs---------------------------------
$SearchN = "FinalScript"
$fw = Import-Csv C:\users\basim\desktop\Licenses.csv
#------------------------------------------------------------------------
$UserCredential = Get-Credential
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $UserCredential -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking

$folderQuery = $null
$fw | ForEach-Object{
$folderStatistics = Get-MailboxFolderStatistics $_.UserPrincipalName -Archive
    foreach ($folderStatistic in $folderStatistics)
    {
        $folderId = $folderStatistic.FolderId;
        $folderPath = $folderStatistic.FolderPath;
        $encoding= [System.Text.Encoding]::GetEncoding("us-ascii")
        $nibbler= $encoding.GetBytes("0123456789ABCDEF");
        $folderIdBytes = [Convert]::FromBase64String($folderId);
        $indexIdBytes = New-Object byte[] 48;
        $indexIdIdx=0;
        $folderIdBytes | select -skip 23 -First 24 | %{$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -shr 4];$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -band 0xF]}
        $folderQuery = $folderQuery + "folderid:$($encoding.GetString($indexIdBytes)) OR ";
    }
}
$folderQuery = $folderQuery.substring(0,$folderQuery.length-4)

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking
New-ComplianceSearch -ExchangeLocation All -Name $SearchN -ContentMatchQuery $folderquery
Start-ComplianceSearch -Identity $SearchN
Do
{
   $tes = Get-ComplianceSearch -Identity $SearchN
} While ($tes.status -eq "InProgress")

New-ComplianceSearchAction -SearchName $SearchN -Export -Format FxStream -ArchiveFormat PerUserPST -Scope BothIndexedAndUnindexedItems
                                                                  #FxStream: Export to PST files
$ExportName = $SearchN + "_Export"
Do
{
   $complete = Get-ComplianceSearchAction -Identity $ExportName
} While ($complete.Status -ne "InProgress")