#------------------------Update the Details as required----------------------------------------------------
$Group1 = "E3sub1"
$Group2 = "E1sub1"
$Subject = "Duplicates"
$ReportingEmailAddress = "basim.moulana@hirdaramani.com"
$Creds = Get-Credential
Import-Module -Name “C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll”
#----------------------------------------------------------------------------------------------------------
$Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($Creds)
#$Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials("emailaddress","password")
$exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$exchService.Credentials = $Credentials
$exchService.Url = "https://outlook.office365.com/EWS/Exchange.asmx"

$tt = Get-EventLog -InstanceId 4728 -LogName security -Newest 1
$tt.ReplacementStrings[2]

if($tt.ReplacementStrings[2] -eq $Group1){
   $names = Get-AdGroupMember -identity $Group2
   $names | ForEach-Object {
      $name = $_
      if($tt.ReplacementStrings[1] -eq $name.SID){
        $Message = $null
        $Message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($exchService)
        $message.Subject = "Duplicates";
        $message.Body = $name.name + " is duplicated";
        $message.ToRecipients.Add("basim.moulana@hirdaramani.com");
        $message.SendAndSaveCopy();
     }
   } 
}
if($tt.ReplacementStrings[2] -eq $Group2){
    $names = Get-AdGroupMember -identity $Group1
    $names| ForEach-Object {
      $name = $_
      if($tt.ReplacementStrings[1] -eq $name.SID){
        $Message = $null
        $Message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($exchService)
        $message.Subject = $Subject;
        $message.Body = $name.name + " is duplicated";
        $message.ToRecipients.Add($ReportingEmailAddress);
        $message.SendAndSaveCopy();
     }
   }
}