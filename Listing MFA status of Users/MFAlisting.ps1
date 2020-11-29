$File = "C:\users\basim\desktop\MFA.txt"

Connect-MsolService -Credential $credential

$users = Get-MsolUser -All
$users | ForEach-Object {

$user = $_
	If ($user.StrongAuthenticationRequirements.State -ne $null){
	$user.UserPrincipalName + " Enabled" >> $file	
	}else{
	$user.UserPrincipalName + " Disabled" >> $file	
	}
}
