
Function Set-ADOwner($ADOwnerSAM, $ADGroupName, $cred) {
    $ADOwner = Get-ADUser -Identity $ADOwnerSAM
    $ADGroup = GEt-ADgroup -Identity $ADGroupName 
    $ADGroup | Set-ADGroup -replace @{owner = $ADEigenaar.DistinguishedName } -Credential $cred
}

