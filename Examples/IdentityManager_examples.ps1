# Configuration 
$OIMServer   = "192.168.56.101 " 
$OIMAppNamme =  "D1IMAppServer"

$cred = Get-Credential
$con = Connect-OIM -AppServer $OIMServer  -AppName $OIMAppNamme   -Credential $cred

Get-OIMObject -ObjectName Person -Where "Lastname like 'Lo%' "   -First -Full
Get-OIMObjectfromURI -uri "/D1IMAppServer/api/entity/Person/604b2bad-a34c-4c58-a0d8-6ea86e61ba5c" 

New-OIMObject -ObjectName Person -Properties @{Firstname="test";Lastname="test"}

Add-OIMObjectMember -TableName PersonInOrg -TableColumn UID_Person -UID xxx-xxx -Properties @('aaaa','ssss','ddd')

Remove-OIMObjectMember -TableName PersonInOrg -TableColumn UID_Person -UID xxx-xxx -Properties @( 'aaaa','ssss','ddd')
Remove-OIMObject $obj

Start-OIMScript -ScriptName QER_GetWebBaseURL
Start-OIMMethod -Object $obj -MethodName ExecuteTemplates -Parameters @()
Start-OIMMethod -Object $obj -MethodName ExecuteTemplate -Parameters @("UID_ORg")
Start-OIMEvent  -Object $obj -EventName ExecuteTemplates -Parameters @{}
Start-OIMScript -ScriptName CCC_GetHostNameFromSystem -Parameters @("sasfdsa")

Disconnect-OIM $con

Connect-OIMSQL -servername sbx-IAMDB-9001.sandbox.local -database D2IMv7 -Cred $cred
Run-OIMSQLQuery -sqlquery "select top 1 * from person"





Set-OIMConfigParameter -FullPath "Custom\SourceSystems\YouForceAPI\UseLocalInputFiles" -Value "False"