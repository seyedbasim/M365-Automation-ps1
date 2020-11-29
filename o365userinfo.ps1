$path = "C:\userinfo.csv"

Add-Content -Path $path  -Value '"Name","Title","CompanyName","Office","UserPrincipalName","Mail","MFA Status","AssignedLicense"'

Connect-MsolService

$users = Get-MsolUser -All

foreach($user in $users){
        ########################################################
        $Name = $user.DisplayName
        $Title = $user.Title
        $Department = $user.Department
        $UserPrincipleName = $user.UserPrincipalName
        $Office = $user.Office
        $mail = $user.ProxyAddresses[1]
        $Company = $null
        ##################MFA###################################
        If ($user.StrongAuthenticationRequirements.State -ne $null){
	        $mfa = "Enabled"	
	        }else{
	        $mfa = "Disabled"	
        }
        #####License#############################################
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
         if($Licenses.count -eq 0) 
         { 
          $AssignedLicense="No License Assigned" 
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

