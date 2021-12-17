$apikey = Read-Host -Prompt "Fill in the PowerShell Gallery API key"

$config = Get-PSPackageProjectConfiguration -ConfigPath $PSScriptRoot
$config = Get-PSPackageProjectConfiguration -ConfigPath "C:\Users\rloom_q\OneDrive\Documenten\Git\IdentityManager\IdentityManager"


$script:SrcModulePath = $config.SourcePath +"\" + $config.ModuleName + ".psd1" 

$script:BuildOutputPath =  $config.BuildOutputPath + "\*"
$script:SrcPath = $config.SourcePath + "\*"
$IMPath = "C:\Users\rloom_q\OneDrive\Documenten\Git\IdentityManager\IdentityManager\out\IdentityManager"

New-Item C:\Users\rloom_q\OneDrive\Documenten\Git\IdentityManager\IdentityManager\out\IdentityManager -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
Copy-Item -Path C:\Users\rloom_q\OneDrive\Documenten\Git\IdentityManager\IdentityManager\src\* -destination $IMPath 
Publish-Module   -NuGetApiKey $apikey -Path $IMPath
Remove-Item $IMPath\* -Force



Import-Module  $IMPath
Test-ModuleManifest -Path $script:SrcModulePath -Verbose 

Import-Module C:\Users\rloom_q\OneDrive\Documenten\Git\IdentityManager\IdentityManager\out\IdentityManager\IdentityManager.psd1 