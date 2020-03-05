# Configuration 

$dicEnvironment = @{
    DEV = @{Server="WIN-DMTVK12KPU5";AppName="D1IMAppServer"}
    TST = @{Server="sbx-iam-9001.sandbox.local";AppName="D1IMAppServer"}
    ACC = @{Server="abx-iam-8001.accbox.local";AppName="D1IMAppServer"}
    PRD = @{Server="HKT-iam-0001.jumbo.local";AppName="D1IMAppServer"}


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



$con = Connect-OIM -AppServer $OIMServer  -AppName $OIMAppNamme   -Credential $cred

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





Set-OIMConfigParameter -FullPath "Custom\SourceSystems\YouForceAPI\UseLocalInputFiles" -Value "False"

Get-OIMObject -ObjectName CCC_HRFunction -Where "ccc_medewerker like '9%'"
#select UID_ccc_HRfunction from CCC_HRFunction where ccc_medewerker like '" + personnel_number + "%'"


$TransportFile = "C:\Temp\20200203 R10 Update Password fix.zip" 
$ProgramsFolder = "C:\Temp\One Identity Manager v8.1"
$DatabaseConnectionString= "Data Source=WIN-DMTVK12KPU5;Initial Catalog=D2IMv7;User Id=sa_d1im;Password=Passw0rd!;"
 #$Credential = Get-Credential
 . Modules\IdentityManager\IdentityManager_CMD.ps1
$process = Install-OIMTransportFiles -TransportFile  $TransportFile -ProgramsFolder $ProgramsFolder  -DatabaseConnectionString $DatabaseConnectionString -Credential $Credential
$process.ExitCode

$obj = Get-OIMPerson -Lastname Laagland
Get-OIMPerson  -Object $obj -full
 -FirstName test%