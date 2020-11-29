$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking

Add-MailboxPermission -Identity "Irfan.siddique@midassafety.onmicrosoft.com" -User "rehan.alam@midassafety.com" -AccessRights ReadPermission -InheritanceType All

Get-MailboxFolderPermission -Identity "Irfan.siddique@midassafety.onmicrosoft.com"

Add-MailboxFolderPermission -Identity Irfan.siddique@midassafety.onmicrosoft.com -User rehan.alam@midassafety.com -AccessRights Reviewer

Add-MailboxFolderPermission -Identity Irfan.siddique@midassafety.onmicrosoft.com:\inbox -User rehan.alam@midassafety.com -AccessRights Reviewer
