$path = "C:\userinfo.csv"

$creds  = Get-Credential

$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $creds -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking

#Get-Mailbox -Identity "basim@honelab.onmicrosoft.com" | fl

Add-Content -Path $path  -Value '"Name","Title","CompanyName","Office","UserPrincipalName","Mail","MFA Status","AssignedLicense"'

Connect-MsolService -Credential $creds 
$users = Get-MsolUser -All

foreach($user in $users){
        ############Exchange online#############################
        $mail = (Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction Ignore).PrimarySmtpAddress 

        ########################################################
        $Name = $user.DisplayName
        $Title = $user.Title
        $Department = $user.Department
        $UserPrincipleName = $user.UserPrincipalName
        $Office = $user.Office
        #$mail = $user.ProxyAddresses[1]
        $Company = $null
        ##################MFA###################################
        If ($user.StrongAuthenticationRequirements.State -ne $null){
	        $mfa = "Enabled"	
	        }else{
	        $mfa = "Disabled"	
        }
        #####License#############################################
        $AssignedLicense = $null
        $Licenses=$User.Licenses.AccountSkuId
        $FriendlyNameHash = Get-Content -Raw -Path .\LicenseFriendlyName.txt -ErrorAction Stop | ConvertFrom-StringData 

         foreach($License in $Licenses) 
         { 
            $Count++ 
            $LicenseItem= $License -Split ":" | Select-Object -Last 1 
            $EasyName=$FriendlyNameHash[$LicenseItem] 
            if(!($EasyName)) 
            {$NamePrint=$LicenseItem} 
            else 
            {$NamePrint=$EasyName} 
            $AssignedLicense=$AssignedLicense+$NamePrint 
            if($count -lt $licenses.count) 
            { 
              $AssignedLicense=$AssignedLicense+"," 
            } 
         } 
         ##########################################################
          $hash = @{
             "Name" = $Name 
             "Title" = $Title
             "CompanyName" = $Company
             "Office" = $Office
             "UserPrincipalName" = $UserPrincipleName
             "Mail" = $mail
             "MFA Status" = $mfa
             "AssignedLicense" = $AssignedLicense
             }
         $newRow = New-Object PsObject -Property $hash
         Export-Csv $path -inputobject $newrow -append -Force
         ###########################################################
}


