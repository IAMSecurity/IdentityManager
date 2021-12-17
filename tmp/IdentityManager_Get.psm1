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


