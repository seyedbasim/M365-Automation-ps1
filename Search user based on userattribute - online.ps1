Connect-MsolService

$proxyaddress = "smtp:user01@contoso.onmicrosoft.com"
get-msoluser -all | where {[string] $str = ($_.proxyaddresses); $str.tolower().Contains($proxyaddress.tolower()) -eq $true}