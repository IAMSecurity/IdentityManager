
$dicEnvironment = @{
    DEV = @{Server="WIN-DMTVK12KPU5";AppName="D1IMAppServer"}
    TST = @{Server="sbx-iam-9001.sandbox.local";AppName="D1IMAppServer"}
    ACC = @{Server="abx-iam-8001.accbox.local";AppName="D1IMAppServer"}
    PRD = @{Server="HKT-iam-0001.jumbo.local";AppName="D1IMAppServer"}


}

if([string]::isnullorEmpty(  $SelectedEnvironment)){
    $SelectedEnvironment = $dicEnvironment | Out-GridView  -OutputMode Single 
    $OIMServer   = $SelectedEnvironment.Value["Server"]
    $OIMAppName = $SelectedEnvironment.Value["AppName"] 
    $cred = Get-Credential -Message "Crendtials for environment $($SelectedEnvironment.Name)"
}
<#
$SelectedEnvironment = $null
#>


$env:PSModulePath += ";\\jumbo.local\hkt-usrdata\r\roblooman\Documents\GitHub\IdentityManager\Modules"
Import-Module IdentityManager -force
$con = Connect-OIM -AppServer $OIMServer  -AppName $OIMAppName   -Credential $cred

Get-OIMObject -ObjectName UNSAccountB -Where "cn like '0042%' "   -Full | Export-Csv -path c:\temp\UNSAccountB_0042_Start.csv 
Get-OIMObject -ObjectName UNSAccountB -Where "cn like '9999%' "   -Full | Export-Csv -path c:\temp\UNSAccountB_9999_Start.csv 
Get-OIMObject -ObjectName UNSGroupB -Where "cn like '%0042' "   -Full | Export-Csv -path c:\temp\UNSGroupB_0042_Start.csv 
Get-OIMObject -ObjectName UNSGroupB -Where "cn like '%9999' "   -Full | Export-Csv -path c:\temp\UNSGroupB_9999_Start.csv 
Get-OIMObject -ObjectName Org -Where "CustomProperty03 = '0042' "   -Full | Export-Csv -path c:\temp\Org_0042_Start.csv 
Get-OIMObject -ObjectName Org -Where "CustomProperty03 = '9999' "   -Full | Export-Csv -path c:\temp\Org_9999_Start.csv 


<#

Get-OIMObject -ObjectName Person -Where "Lastname like 'Lo%' "   -First -Full
Get-OIMObject -ObjectName department -Where "Lastname like 'Lo%' "   -First -Full
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
Invoke-OIMSQLQuery -sqlquery "select top 1 * from person"



#>