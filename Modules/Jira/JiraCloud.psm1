Invoke-RestMethod  
$cred = Get-Credential 
$Text = "rob.looman@jumbo.com:kDCvIX4kq8LDezNVtQrM3C89"
$Bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
$EncodedText = [Convert]::ToBase64String($Bytes)
$EncodedText

 $t = Invoke-RestMethod -Uri "https://jumbo-supermarkten.atlassian.net:rob.looman@kDCvIX4kq8LDezNVtQrM3C89/rest/api/2/issue/IAMD-1189" -Method GET -Headers @{Accept = "application/json" } -Credential $cred
  -UseDefaultCredentials -Headers @{Accept = "application/json" } 
 -Body $authJson.ToString() -Method Post -UseDefaultCredentials -Headers @{Accept = "application/json" } -SessionVariable session | Out-Null

     Invoke-RestMethod -Uri "$Global:OIM_BaseURL/auth/apphost"  -Method Post -UseDefaultCredentials -Headers @{Accept = "application/json" ;Authorization= "Basic $($EncodedText)"} -SessionVariable session | Out-Null
     
     $url = "https://jumbo-supermarkten.atlassian.net/rest/api/3/issue/IAMD-1189"
     $tmp = Invoke-RestMethod -UseDefaultCredentials -Uri $url  -Method Get  -Headers @{Accept = "application/json" ;Authorization= "Basic $($EncodedText)"} 

     Invoke-WebRequest