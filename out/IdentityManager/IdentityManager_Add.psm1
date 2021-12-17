
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