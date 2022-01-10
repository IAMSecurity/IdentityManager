

Function Add-OIMObjectMember{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
            $TableName,
        [Parameter(Mandatory=$true)]
            $TableColumn,
        [Parameter(Mandatory=$true)]
            $UID,
        [Parameter(Mandatory=$true)]
            [Array] $Members
    )

    $body = @{members = $Members }
    $uri = "$Script:BaseURI/api/assignments/$TableName/$TableColumn/$UID"
    Invoke-OIMRestMethod -Uri $uri  -Method Post -Body $body -WebSession $Script:WebSession
}

Function Remove-OIMObjectMember{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
            $TableName,
        [Parameter(Mandatory=$true)]
            $TableColumn,
        [Parameter(Mandatory=$true)]
            $UID,
        [Parameter(Mandatory=$true)]
            [Array] $Members
    )
    $body = @{members = $Members }
    $uri = "$Script:BaseURI/api/assignments/$TableName/$TableColumn/$UID"
    Invoke-OIMRestMethod -Uri  $uri  -Method Delete -Body $body -WebSession $Script:WebSession

}






Function Set-OIMConfigParameter{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $FullPath,
        [Parameter(Mandatory=$true)]
        $value
    )
    $obj =  Get-OIMObject -ObjectName DialogConfigParm -Where "FullPath = '$FullPath' "  -First 1
    Update-OIMObject -Object $obj -Properties @{Value = "$value"}
}

Function Get-OIMConfigParameter{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $FullPath
    )
    $ConfigParam = Get-OIMObject -ObjectName DialogConfigParm -Where "FullPath = '$FullPath' "  -First 1
    $ConfigParam.Value
}

Function Start-OIMSyncProject{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $DisplayName,
        [switch]$wait
    )
    $obj =  Get-OIMObject -ObjectName DPRProjectionStartInfo -Where "displayname = '$DisplayName'"  -First 1

    If ($null -eq $obj ){
        Write-Warning "Sync start configuration not found ($displayname)"
    }else{
        Start-OIMEvent  -Object $obj -EventName run -Parameters @{}

        if($wait){
            Wait-OIMJobQueue -JobChainName "DPR_DPRProjectionStartInfo_Run_Synchronization"
            #Param in Contains obj.uid
        }
    }
}

Function Start-OIMSchedule{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $Name,
        [switch]$wait
    )
    $obj =  Get-OIMObject -ObjectName DialogSchedule -Where "Name = '$Name'"  -First 1
    Start-OIMEvent  -Object $obj -EventName run -Parameters @{}

    if($wait){
        Wait-OIMJobQueue -JobChainName $obj.uid
    }
}







