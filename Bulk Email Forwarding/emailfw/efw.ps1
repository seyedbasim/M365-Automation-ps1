$fw = Import-Csv C:\emailfw\emailfw.csv
$fw | ForEach-Object{ New-MailContact -Name $_.newexternalcontactname -ExternalEmailAddress $_.ForwardingSMTPAddress; 
					  Set-Mailbox -Identity $_.Identity -DeliverToMailboxAndForward $true -ForwardingSMTPAddress $_.ForwardingSMTPAddress }

