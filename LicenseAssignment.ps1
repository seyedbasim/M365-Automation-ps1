#--Get-MsolAccountSku------------
#--------Create Licenses.CSV------------------------
#-----------------Add the path----------------------
$path = "C:\Users\Basim\Desktop\Licenses.csv"
#---------------------------------------------------
#Connect-MsolService
$Folder = "C:\LicenseAssign"
if (!(Test-path $Folder -PathType Container)){
    New-Item -ItemType Directory -Force -Path $Folder
}
Copy-Item $path -Destination $Folder

$path1 = "C:\LicenseAssign\Licenses.csv"
$Users = Import-Csv $path1
$Users | ForEach-Object{ 
        Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses honelab:POWER_BI_PRO,honelab:ENTERPRISEPACK -RemoveLicenses honelab:FLOW_FREE
}

Add-Content -Path C:\LicenseAssign\LicensesAfter.csv  -Value '"UserPrincipalName","Licenses"'
$User1 = Import-Csv $path1
$User1 | ForEach-Object{ 
   $lic = Get-MsolUser -UserPrincipalName $_.UserPrincipalName
   $licen = $null
   $lic.Licenses.AccountSkuId | ForEach-Object{
        $licen = $_ + "," + $licen
   }
   $hash = @{
            "UserPrincipalName" = $_.UserPrincipalName 
            "Licenses" = $licen
            }
   $newRow = New-Object PsObject -Property $hash
   Export-Csv C:\LicenseAssign\LicensesAfter.csv -inputobject $newrow -append -Force
}
