
Function Run-OIMDataImporter($ImportFile, $DefintionFile, $DatabaseConnectionString, $Credential, $loglevel = "off" , $Culture =  (Get-Culture).Name, $ProgramsFolder   ){

    if ($null -eq $Cred ) {
        #Single sign
        $AuthString = "Module=RoleBasedADSAccount" 
    }
    else {
        $user = $Cred.Username
        $Pass = $Cred.GetNetworkCredential().password
        $AuthString = "Module=DialogUser;User=$user;Password=$Pass" 

    }


    $Arguments = @(
        "/conn '$DatabaseConnectionString'"        
        "/ImportFile '$ImportFile'"
        "/DefintionFile '$DefintionFile'"
        "/Auth '$AuthString'"
        "/loglevel $loglevel"
        "/culture $culture"


    )
    

    $ArgumentList =  [string]::join(" ",$Arguments)
    $runFile = Join-Path $ProgramsFolder DataImporterCMD.exe
    $process = Start-Process -FilePath $runFile -ArgumentList $ArgumentList -WorkingDirectory $ProgramsFolder -NoNewWindow -PassThru -Wait

    $process.ExitCode
}
$ProgramsFolder = "C:\Temp\One Identity Manager v8.1"
Run-OIMDataImporter -ProgramsFolder "C:\Temp\One Identity Manager v8.1" -loglevel Debug 

