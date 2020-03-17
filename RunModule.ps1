
$SourcePath = "$env:HOMEDRIVE\Documents\GitHub\IdentityManager\Modules"
$SourcePath  = "$(Split-Path -parent $PSCommandPath)\Modules"
$DestinationPath =  "$env:HOMESHARE\Documents\PowerShell\Modules"
$SourcePath
$DestinationPath 
Copy-Item -Path  $SourcePath -Destination $DestinationPath   -Container -Force
