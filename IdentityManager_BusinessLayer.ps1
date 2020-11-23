Function Get-OIMPerson($personnelnumber, $Session = $Global:OIM_Session, [switch]$First, [switch]$Full){
    Get-OIMObject -ObjectName Person -Where "PersonnelNumber = '$personnelnumber' " -Full:$full -Session $Session
}

Function Get-OIMEset($name, $Session = $Global:OIM_Session, [switch]$First, [switch]$Full){
    Get-OIMObject -ObjectName Eset -Where "ident_eset = '$name' " -Full:$full -Session $Session
}

Function Get-OIMUNSAccountB($name, $Session = $Global:OIM_Session, [switch]$First, [switch]$Full){
    Get-OIMObject -ObjectName UNSAccountB -Where "cn = '$name' " -Full:$full -Session $Session
}

Function Get-OIMUNSGroupB($name, $Session = $Global:OIM_Session, [switch]$First, [switch]$Full){
    Get-OIMObject -ObjectName UNSGroupB -Where "cn = '$name' " -Full:$full -Session $Session
}
Function Get-OIMUNSAccountBInUNSGroupB{
    [CmdletBinding()]     
    Param($UNSAccountB,$UNSGroupB, $Session = $Global:OIM_Session, [switch]$First, [switch]$Full)
    
    if([string]::IsNullOrEmpty($UNSAccountB.uid)){
        $UNSAccountB = Get-OIMUNSAccountB -name $UNSAccountB
    }
    if([string]::IsNullOrEmpty($UNSGroupB.uid)){
        $UNSGroupB = Get-OIMUNSGroupB -name $UNSGroupB
    }
    Get-OIMObject -ObjectName UNSAccountBInUNSGroupB -Where "UID_UNSAccountB = '$($UNSAccountB.uid)' AND UID_UNSGroupB = '$($UNSGroupB.uid)' " -Full:$Full
}


Function Add-OIMPersonHasEset{
    [CmdletBinding()] 
    Param($Person,$Eset)
    if($Person -isnot [array]){
        $listperson = @($person)
    }else{$listperson= $Person}

    if($Eset -isnot [array]){
        $listEset = @($Eset)
    }else{$listEset = $eset}

    ForEach($PersonItem in $listperson){
        ForEach($EsetItem in $listEset){
            if(-not [string]::IsNullOrEmpty($PersonItem.UID) -and  -not [string]::IsNullOrEmpty($EsetItem.UID)){
                Write-Host "New PersonHasEset Person:$($PErsonItem.UID) ESET:$($EsetItem.UID)"
                New-OIMObject -ObjectName PersonHasEset -Properties @{UID_Eset=$EsetItem.UID_ESet;UID_Person=$PersonItem.UID_Person} 
            }else{
                Write-Warning "Invalid UID Person:$($PErsonItem.UID) ESET:$($EsetItem.UID)"
            }
        }

    }

}