Connect-MsolService

$users = Get-Msoluser -All

foreach($user in $users){

$Licenses = $User.Licenses.AccountSkuId

foreach($License in $Licenses){
    $Sku = $License.split(":")[1]
    if($sku -eq "ENTERPRISEPACK"){
#*****************Add Log*********************************************************************
        $user.UserPrincipalName >> $logpath
        $AllLicenses = $User.Licenses | where AccountSkuId -eq $License
        "License: " + $AllLicenses.AccountSkuId >> $logpath
        $AllLicenses.ServiceStatus | Format-Table >> $logpath
        "" >> $logpath
#*****************Set Apps********************************************************************
        $LO = New-MsolLicenseOptions -AccountSkuId $License
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -LicenseOptions $LO
#*********************************************************************************************
    }elseif($sku -eq "STANDARDPACK"){
#*****************Add Log*********************************************************************
        $user.UserPrincipalName >> $logpath
        $AllLicenses = $User.Licenses | where AccountSkuId -eq $License
        "License: " + $AllLicenses.AccountSkuId >> $logpath
        $AllLicenses.ServiceStatus | Format-Table >> $logpath
        "" >> $logpath
#*****************Set Apps********************************************************************
        $LO = New-MsolLicenseOptions -AccountSkuId $License
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -LicenseOptions $LO
#*********************************************************************************************
    }elseif($sku -eq "DESKLESSPACK"){
#*****************Add Log*********************************************************************
        $user.UserPrincipalName >> $logpath
        $AllLicenses = $User.Licenses | where AccountSkuId -eq $License
        "License: " + $AllLicenses.AccountSkuId >> $logpath
        $AllLicenses.ServiceStatus | Format-Table >> $logpath
        "" >> $logpath
#*****************Set Apps********************************************************************
        $LO = New-MsolLicenseOptions -AccountSkuId $License
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -LicenseOptions $LO
#*********************************************************************************************
    }elseif($sku -eq "M365_F1_COMM"){
#*****************Add Log*********************************************************************
        $user.UserPrincipalName >> $logpath
        $AllLicenses = $User.Licenses | where AccountSkuId -eq $License
        "License: " + $AllLicenses.AccountSkuId >> $logpath
        $AllLicenses.ServiceStatus | Format-Table >> $logpath
        "" >> $logpath
#*****************Set Apps********************************************************************
        $LO = New-MsolLicenseOptions -AccountSkuId $License
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -LicenseOptions $LO
#*********************************************************************************************
    }
}

}

