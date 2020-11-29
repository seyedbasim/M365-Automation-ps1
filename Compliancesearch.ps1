$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

Start-ComplianceSearch -Identity "test123"

New-ComplianceSearch -ExchangeLocation All -Name "test123" -ContentMatchQuery $folderquery

$tes = Get-ComplianceSearch -Identity "test123"

$SearchName = "test123"

New-ComplianceSearchAction -SearchName $SearchName -Export -Format FxStream -ArchiveFormat PerUserPST
                                                                  #FxStream: Export to PST files

$ExportName = $SearchName + "_Export"

$complete = Get-ComplianceSearchAction -Identity $ExportName

$complete.Status
