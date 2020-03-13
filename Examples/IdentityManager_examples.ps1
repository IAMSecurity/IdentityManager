# Configuration 
$RealseFolder = "C:\Temp\" 
$ProgramsFolder = "C:\Temp\One Identity Manager v8.1"
$DatabaseConnectionString= "Data Source=WIN-DMTVK12KPU5;Initial Catalog=D2IMv7;User Id=sa_d1im;Password=Passw0rd!;"


$dicEnvironment = @{
    "1_DEV" = @{Server="WIN-DMTVK12KPU5";AppName="D1IMAppServer";Database="D2IMv7"}
    "2_TST" = @{Server="sbx-iam-9001.sandbox.local";AppName="D1IMAppServer";Database="D2IMv7"}
    "3_ACC" = @{Server="abx-iam-8001.accbox.local";AppName="D1IMAppServer";Database="D2IMv7"}
    "4_PRD" = @{Server="HKT-iam-0001.jumbo.local";AppName="D1IMAppServer";Database="D1IMv7"}
}

if([string]::isnullorEmpty(  $SelectedEnvironment)){
    $SelectedEnvironment = $dicEnvironment | Out-GridView  -OutputMode Single 
    $OIMServer   = $SelectedEnvironment["Server"]
    $OIMAppNamme = $SelectedEnvironment["AppName"] 
    $cred = Get-Credential -Message "Crendtials for environment $($SelectedEnvironment.Name)"
}
<#
$SelectedEnvironment = $null
#>
Import-Module IdentityManager


$con = Connect-OIM -AppServer $OIMServer  -AppName $OIMAppNamme   -Credential $cred

Get-OIMObject -ObjectName Person -Where "Lastname like 'Lo%' "   -First -Full
Get-OIMObject -ObjectName department -Where "Lastname like 'Lo%' "   -First -Full
Get-OIMObject -ObjectName CCC_HRFunction -Where "ccc_medewerker like '9%'"
Get-OIMObjectfromURI -uri "/D1IMAppServer/api/entity/Person/604b2bad-a34c-4c58-a0d8-6ea86e61ba5c" 

Set-OIMConfigParameter -FullPath "Custom\SourceSystems\YouForceAPI\UseLocalInputFiles" -Value "False"

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

Connect-OIMSQL -servername HKT-IAMDB-0001.jumbo.local -database D1IMv7

$CredProdDB = Get-Credential
Connect-OIMSQL -servername HKT-IAMDB-0001.jumbo.local -database D1IMv7 -Cred $CredProdDB
Invoke-OIMSQLQuery -sqlquery "select top 1 * from Org"

Invoke-OIMSQLQuery -sqlquery "select top 1 * from person"


Install-OIMTransportFiles -TransportFile  $TransportFile -ProgramsFolder $ProgramsFolder  -DatabaseConnectionString $DatabaseConnectionString -Credential $Credential




