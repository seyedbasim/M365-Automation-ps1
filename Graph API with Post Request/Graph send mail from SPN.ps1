$date = Get-Date
$datep = $date.AddDays(-30)
$dates = $datep | Get-Date -Format "yyyy-MM-dd"

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/x-www-form-urlencoded")
$headers.Add("Cookie", "x-ms-gateway-slice=prod; stsservicecookie=ests; fpc=AtoBfWCWsvBMjQrly80bOHlXF97mAQAAAL8UK9cOAAAA")

$body = "client_id=b4d8c7aa-27b6-4289-b2bd-1ef445027c0f&scope=https%3A//graph.microsoft.com/.default&client_secret=mTx%7E5FFE295IS1-spSs4Pc-%7Eu%7Ebf3QkkH1&grant_type=client_credentials"

$response = Invoke-RestMethod 'https://login.microsoftonline.com/h1lab.onmicrosoft.com/oauth2/v2.0/token' -Method 'POST' -Headers $headers -Body $body
#$response | ConvertTo-Json
$addbody = "The new cafeteria is open."

$headers = $null
$body = $null
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$token = "Bearer " + $response.access_token
$headers.Add("Authorization", $token)
$headers.Add("Content-Type", "application/json")


$Header = @"
    <style>
    table{
        border-collaps: collapse;
    }
    th,td{
        text-align: left;
        padding: 5px;
        color: black;
        background-color: white;
    }
    tr:nth-child(even){background-color: #f2f2f2}
    th{
        background-color: #ffd500;
        color: black;
    }
    </style>
"@

$mailmessage = "<p>Dear Sir,<br><br>

Please Find Details of users logged into PowerBI:: <br>

</p>"

$mailsignature = "<p>
Thanks & Regards,<br>
Azure Automation Team<br><br>
</p>"

[System.Collections.ArrayList]$ArrayWithHeader = @()
 #   foreach($server in $GroupByServerName){
 #       foreach($upd in $server.group){
        $val = [pscustomobject]@{'Server'="server1";'Update'="update1"}
        $ArrayWithHeader.add($val) | Out-Null
        $val=$null
        $val = [pscustomobject]@{'Server'="server1";'Update'="update1"}
        $ArrayWithHeader.add($val) | Out-Null
        $val=$null
#        }  
 #   }


$body1 = $ArrayWithHeader | ConvertTo-Html -Property 'Server','Update' -As Table -Head $Header | Out-String
$full_body = $null
$full_body = $mailmessage + $body1 + $mailsignature


$body = $null
$body = "{
`n                `"message`":
`n                {
`n                    `"subject`": `"Power BI Logon Report`",
`n                    `"body`": 
`n                    {
`n                        `"contentType`": `"HTML`",
`n                        `"content`": `'"+$full_body+".`'
`n                    },
`n                    `"toRecipients`": [
`n                    {
`n                        `"emailAddress`": 
`n                        {
`n                            `"address`": `"basim.moulana@hirdaramani.com`"
`n                        }
`n                    }
`n                    ]
`n                }
`n}"

$response = Invoke-RestMethod 'https://graph.microsoft.com/v1.0/users/basim@h1lab.club/sendMail' -Method 'POST' -Headers $headers -Body $body
