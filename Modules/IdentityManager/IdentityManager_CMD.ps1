function Install-OIMTransportFiles {
    [CmdletBinding()]
    param (
        # Path to the directory containing the transport files.
        [Parameter(Mandatory=$true)]
        [Alias('Source')]
        [ValidateScript ({ Test-Path -Path $_  })]
        [string] $TransportFile,

        # Path to the directory where the workstation tools are installed
        [Parameter(Mandatory=$true)]
        [Alias('Tools')]
        [ValidateScript ({ (Test-Path -Path (Join-Path  -Path $_  -ChildPath 'DBTransporterCmd.exe' ) -PathType Leaf) })]
        [string]
        $ProgramsFolder,      

        # Database server name or FQDN
        [Parameter(Mandatory=$true)]
        [string]
        $DatabaseConnectionString,

        $Credential


      

    )


    if ($null -eq $Credential ) {
        #Single sign
        $AuthString = "Module=RoleBasedADSAccount" 
    }
    else {
        $user = $Credential.Username
        $Pass = $Credential.GetNetworkCredential().password
        $AuthString = "Module=DialogUser;User=$user;Password=$Pass" 

    }


    $Arguments = @(   
        "/File=""$TransportFile"""
        "/Conn=""$DatabaseConnectionString"""        
        "/Auth=""$AuthString"""
        "-V"


    )
    

    $ArgumentList =  [string]::join(" ",$Arguments)
    Write-Host $ArgumentList 
    $runFile = Join-Path $ProgramsFolder DBTransporterCmd.exe
    #$scriptblock = {Param ($runFile,$ArgumentList,$ProgramsFolder)  ;Start-Process -FilePath $runFile -ArgumentList $ArgumentList -WorkingDirectory $ProgramsFolder -NoNewWindow -PassThru -Wait }
    #$process = Invoke-Command -ScriptBlock $scriptblock -ArgumentList $runFile,$ArgumentList,$ProgramsFolder
    $process = Start-Process -FilePath $runFile -ArgumentList $ArgumentList -WorkingDirectory $ProgramsFolder -PassThru -Wait -RedirectStandardOutput C:\Temp\IAM\test.txt -RedirectStandardError C:\Temp\IAM\test1.txt -WindowStyle Hidden
    $process

}


Function Run-OIMDataImporter($ImportFile, $DefintionFile, $DatabaseConnectionString, $Credential, $loglevel = "off" , $Culture =  (Get-Culture).Name, $ProgramsFolder ,$ComputerName = 'localhost', $RemoteCredential ){

    if ($null -eq $Credential ) {
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
    $scriptblock = {Param ($runFile,$ArgumentList,$ProgramsFolder)  ;Start-Process -FilePath $runFile -ArgumentList $ArgumentList -WorkingDirectory $ProgramsFolder -NoNewWindow -PassThru -Wait }
    $process = Invoke-Command -ScriptBlock $scriptblock -ComputerName $ComputerName -Credential $RemoteCredential -ArgumentList $runFile,$ArgumentList,$ProgramsFolder

    $process.ExitCode
}
