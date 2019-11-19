

$GPOBackupFolder = "C:\Users\A9321871\Documents\_temp\backup"
$GPOBackupFolderGptmlBackup = "C:\Users\A9321871\Documents\_temp\backup"

Function Add-GPUsersToGptmpl {
    Param(
        $SourcePath,
        $DestinationPath,
        $oldDomain = "VERZ", 
        $newDomain ="INS"
    )
    Process{

        Get-Content $SourcePath |ForEach-Object{
            $line = $_

            if($_.Contains("*S-1-")){
                $addUsers = ""
                $KeyValues = $line.Split("=")
                if($KeyValues.Count -ne 2){
                    Write-Warning "Could not parse line: $line"
                    continue
                }
                $Key = $KeyValues[0]
                $Values = $KeyValues[1]
                $arValues = $Values.Split(",")
                ForEach($value in $arValues){
                    $sid = $value.Replace("*","").Trim()
                    if(![String]::IsNullOrEmpty($sid)){
                        $objSID = New-Object System.Security.Principal.SecurityIdentifier($sid)
                        $objUser = $objSID.Translate( [System.Security.Principal.NTAccount]) 
                        if($objUser.Value -like "$oldDomain\*"){
                            $insUser = $objUser.Value.Replace("$oldDomain\","")
                            $objUser = New-Object System.Security.Principal.NTAccount($newDomain, $insUser)
                                $strSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier]) 
                            $addUsers  += ",*$($strSID.Value)"
        
                        } 
                    }
            
                }#End arValues
                $line += $addUsers
            }
            $line
    

        }| Out-File $DestinationPath -Force
    }#end Process
}


ForEach($GPOFolder in Get-ChildItem $GPOBackupFolder -Directory){
    $GPOBackupFolderGptmlBackup = Join-Path $GPOBackupFolderGptmlBackup "$($GPOFolder.Name )GptTmpl.inf"
    $GptTmplPath = Join-Path $GPOFolder.FullName "DomainSysvol\GPO\Machine\microsoft\windows nt\SecEdit\GptTmpl.inf"

    Copy-Item $GptTmplPath $GPOBackupFolderGptmlBackup -Force
    if(Test-Path $GptTmplPath){
        Add-GPUsersToGptmpl -SourcePath $GPOBackupFolderGptmlBackup -DestinationPath $GptTmplPath -oldDomain "VERZ" -newDomain "INS"

    }

}


Function Get-GPORestrictedGroups($domain){
    $result = @()
    ForEach($gpo in Get-GPO -All -Domain $domain){
        $xmlReport = [xml] ($gpo |Get-GPOReport -ReportType Xml -domain $domain)
        ForEach($group in $xmlReport.GPO.Computer.ExtensionData.Extension.RestrictedGroups){
            $groupName = $group.GroupName.Name.'#text'
            forEach($member in $group.Member){cd 
                $properties = @{
                    "GPOName"=$gpo.DisplayName;
                    "GroupName"=$groupName;
                    "Member"=$member.Name.'#text'
                    "Domain"=$domain

                }
                $result += New-Object -TypeName PSObject -Property $properties
                Write-Host "$($gpo.DisplayName) : $groupName : $($member.Name.'#text')"
            }
        }
    }

   

}

Function Get-GPOUnlinked($domain) {
    $GPOs = Get-GPO -All  -Domain $domain
    ForEach ($GPO in $GPOs) { 
        If ( $GPO | Get-GPOReport -ReportType XML -Domain $domain | Select-String -NotMatch "<LinksTo>" ) {
            $list += $GPO.DisplayName        
            $GPO.DisplayName 
        }
    }

    $list | Sort-Object 
}


function Out-IniFile($InputObject, $FilePath)
{
    $outFile = New-Item -ItemType file -Path $Filepath
    foreach ($i in $InputObject.keys)
    {
        if (!($($InputObject[$i].GetType().Name) -eq "Hashtable"))
        {
            #No Sections
            Add-Content -Path $outFile -Value "$i=$($InputObject[$i])"
        } else {
            #Sections
            Add-Content -Path $outFile -Value "[$i]"
            Foreach ($j in ($InputObject[$i].keys | Sort-Object))
            {
                if ($j -match "^Comment[\d]+") {
                    Add-Content -Path $outFile -Value "$($InputObject[$i][$j])"
                } else {
                    Add-Content -Path $outFile -Value "$j=$($InputObject[$i][$j])" 
                }
 
            }
            Add-Content -Path $outFile -Value ""
        }
    }
}

function Get-IniContent ($filePath)
{
    $ini = @{}
    switch -regex -file $FilePath
    {
        "^\[(.+)\]" # Section
        {
            $section = $matches[1]
            $ini[$section] = @{}
            $CommentCount = 0
        }
        "^(;.*)$" # Comment
        {
            $value = $matches[1]
            $CommentCount = $CommentCount + 1
            $name = "Comment" + $CommentCount
            $ini[$section][$name] = $value
        } 
        "(.+?)\s*=(.*)" # Key
        {
            $name,$value = $matches[1..2]
            $ini[$section][$name] = $value
        }
    }
    return $ini
}


$Path = "C:\Users\A9321871\Documents\_temp\backup\{585F6187-0686-4FEA-AF31-2ED4314197F6}\DomainSysvol\GPO\Machine\microsoft\windows nt\SecEdit\GptTmpl.inf"
$PathNew = "C:\Users\A9321871\Documents\_temp\backup\{585F6187-0686-4FEA-AF31-2ED4314197F6}\DomainSysvol\GPO\Machine\microsoft\windows nt\SecEdit\GptTmplNew.inf"


$iniGptTmpl = Get-IniContent $Path
$iniGptTmplNew =  Get-IniContent $Path
$iniGptTmpl["Group Membership"].Keys |%{
    $new = $iniGptTmpl["Group Membership"][$_]
    $Accounts = $iniGptTmpl["Group Membership"][$_].Split(",")
    ForEach($account in $Accounts){
        $sid = $account.Replace("*","").Trim()
        if(![String]::IsNullOrEmpty($sid)){
            $objSID = New-Object System.Security.Principal.SecurityIdentifier($sid)
            $objUser = $objSID.Translate( [System.Security.Principal.NTAccount]) 
            if($objUser.Value -like "VERZ\*"){
                $insUser = $objUser.Value.Replace("VERZ\","")
                $objUser = New-Object System.Security.Principal.NTAccount("INS.LOCAL", $insUser)
                 $strSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier]) 
                $new  += ",*$($strSID.Value)"
        
            } 
        }
    }
    $iniGptTmplNew["Group Membership"][$_] = $new
}
$iniGptTmpl["Privilege Rights"].Keys |%{
    $new = $iniGptTmpl["Privilege Rights"][$_]
    $Accounts = $iniGptTmpl["Privilege Rights"][$_].Split(",")
    ForEach($account in $Accounts){
        $sid = $account.Replace("*","").Trim()
        $objSID = New-Object System.Security.Principal.SecurityIdentifier($sid)
        $objUser = $objSID.Translate( [System.Security.Principal.NTAccount]) 
        if($objUser.Value -like "VERZ\*"){
            $insUser = $objUser.Value.Replace("VERZ\","")
            $objUser = New-Object System.Security.Principal.NTAccount("INS.LOCAL", $insUser)
             $strSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier]) 
            $new  += ",*$($strSID.Value)"
        
        } 
    }
    $iniGptTmplNew["Privilege Rights"][$_] = $new
}

 Out-IniFile $iniGptTmplNew $PathNew

 Get-ADUser -filter {((UserAccountControl -band 0x10000) -eq 0) -and (samAccountName -eq "9321871")} 