$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking 

Remove-DistributionGroupMember -Identity 'Everyone_region_China@midassafety.com' -Member 'herbert.sun@midassafety.onmicrosoft.com'

