$ModulePath  = "$(Split-Path -parent $PSCommandPath)\Modules"
$ModulePath = "$env:HOMEDRIVE\Documents\GitHub\IdentityManager\Modules"
Copy-Item -Path  $ModulePath -Destination $env:HOMEDRIVE\Documents\WindowsPowerShell -Force -Recurse
