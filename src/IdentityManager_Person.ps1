
Function Get-OIMPerson($Object, $UID, $CentralAccount,$PersonnelNumber, $FirstName,$Lastname, [switch] $full,[switch] $first){

    $arWhere = @()
    if(-not [string]::IsNullOrEmpty($Object)){
        $arWhere += "UID_Person = '$($Object.UID)'"
    }
    if(-not [string]::IsNullOrEmpty($UID)){
        $arWhere += "UID_Person = '$UID'"
    }
    if(-not [string]::IsNullOrEmpty($PersonnelNumber)){
        $arWhere += "PersonnelNumber = '$PersonnelNumber'"
    }
    if(-not [string]::IsNullOrEmpty($FirstName)){
        $arWhere += "FirstName = '$FirstName'"
    }
    if(-not [string]::IsNullOrEmpty($Lastname)){
        $arWhere += "Lastname  = '$Lastname'"
    }


    $Where  = ""
    ForEach($obj in $arWhere){
        if($Where -ne ""){$Where += " AND "}
        if($obj.Contains("%")){
            $Where  += $obj.Replace("=","like")
        }else{
            $Where  += $obj

        }

    }

    Get-OIMObject -ObjectName Person -Where $Where -full:$full -first:$first
}



