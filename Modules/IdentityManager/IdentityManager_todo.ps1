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