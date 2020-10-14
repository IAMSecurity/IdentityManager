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