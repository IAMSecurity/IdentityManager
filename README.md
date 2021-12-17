Important Note
==============
This version of Identity Manager  is currently under development and is not feature complete.

Introduction
============

Identity Manager is a PowerShell module with commands for controlling One Identity Manager.


Requirements
============

- PowerShell 5.0 or higher.

Get IdentityManager Module
========================

Please refer to PowerShell gallery

Get IdentityManager Source
========================

#### Steps

* Obtain the source
    - Download the latest source code from the release page (https://github.com/IAMSecurity/IdentityManager) OR
    - Clone the repository (needs git)
    ```powershell
    git clone https://github.com/IAMSecurity/IdentityManager
    ```

* Navigate to the local repository directory

```powershell
PS C:\> cd c:\Repos\IdentityManager
PS C:\Repos\IdentityManager>
```

* Install PSPackageProject module if needed

```powershell
if ((Get-Module -Name PSPackageProject -ListAvailable).Count -eq 0) {
    Install-Module -Name PSPackageProject -Repository PSGallery
}
```

* Build the project

```powershell
# Build for the netstandard2.0 framework
PS C:\Repos\IdentityManager> .\build.ps1 -Clean -Build -BuildConfiguration Debug -BuildFramework netstandard2.0
```

* Publish the module to a local repository

```powershell
PS C:\Repos\IdentityManager> .\build.ps1 -Publish
```

* Run functional tests

```powershell
PS C:\Repos\IdentityManager> Invoke-PSPackageProjectTest -Type Functional
```

* Import the module into a new PowerShell session

```powershell
# If running PowerShell 6+
C:\> Import-Module C:\Repos\IdentityManager\out\IdentityManager

# If running Windows PowerShell
C:\> Import-Module C:\Repos\IdentityManager\out\IdentityManager\IdentityManager.psd1
```

**Note**  
