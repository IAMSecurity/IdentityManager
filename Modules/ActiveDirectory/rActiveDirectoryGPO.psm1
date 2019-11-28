

Function Add-rGPUsersToGptmpl {
    Param(
        $SourcePath,
        $DestinationPath,
        $oldDomain = "", 
        $newDomain =""
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
                #$Key = $KeyValues[0]
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

Rename-rGPOGroupMembership($OldGptTmlinf, $newGptTmlinf,$oldDomain,$newDomain){
    $Path = "C:\Windows\Temp\backup\{xxx1}\DomainSysvol\GPO\Machine\microsoft\windows nt\SecEdit\GptTmpl.inf"
    $PathNew = "C:\Windows\Temp\backup\{xxx2}\DomainSysvol\GPO\Machine\microsoft\windows nt\SecEdit\GptTmplNew.inf"


    $iniGptTmpl = Get-IniContent $OldGptTmlinf
    $iniGptTmplNew =  Get-IniContent $newGptTmlinf
    $iniGptTmpl["Group Membership"].Keys |%{
        $new = $iniGptTmpl["Group Membership"][$_]
        $Accounts = $iniGptTmpl["Group Membership"][$_].Split(",")
        ForEach($account in $Accounts){
            $sid = $account.Replace("*","").Trim()
            if(![String]::IsNullOrEmpty($sid)){
                $objSID = New-Object System.Security.Principal.SecurityIdentifier($sid)
                $objUser = $objSID.Translate( [System.Security.Principal.NTAccount]) 
                if($objUser.Value -like "$oldDomain\*"){
                    $insUser = $objUser.Value.Replace("$oldDomain\","")
                    $objUser = New-Object System.Security.Principal.NTAccount("$newDomain", $insUser)
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
            if($objUser.Value -like "$oldDomain\*"){
                $insUser = $objUser.Value.Replace("$oldDomain\","")
                $objUser = New-Object System.Security.Principal.NTAccount("$newDomain", $insUser)
                $strSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier]) 
                $new  += ",*$($strSID.Value)"
            
            } 
        }
        $iniGptTmplNew["Privilege Rights"][$_] = $new
    Out-IniFile $iniGptTmplNew $newGptTmlinf
}


