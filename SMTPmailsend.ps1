# Example of how to send a message using the .NET SmtpClient and MailMessage classes

$O365Cred = (Get-Credential)
$MsgFrom = "mail@h1lab.club"
$MsgTo = "seyedbasim95@outlook.com"
$SmtpServer = "smtp.office365.com" ; $SmtpPort = "587"
# Build Message Properties
$Message = New-Object System.Net.Mail.MailMessage $MsgFrom, $MsgTo
$Message.Subject = "Example Message" 
#$Message.Attachments.Add("C:\Temp\Office365TenantUsage.csv")
$Message.Headers.Add("X-O365ITPros-Header","Important Email")
$Message.Body = ""
$Message.IsBodyHTML = $True
# Build the SMTP client object and send the message off
$Smtp = New-Object Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
$Smtp.EnableSsl = $True
$Smtp.Credentials = $O365Cred
$Smtp.Send($Message)
