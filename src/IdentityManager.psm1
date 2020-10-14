

<#
.SYNOPSIS
   Connect to a One Identity Manager REST API server
.DESCRIPTION
    This command starts a connection to a One Identity Manager REST API server
.EXAMPLE
    C:\PS> Connect-OIM -AppServer xxx.company.local -UseSSL
    This example starts a connection to https://xxx.company.local/Appserver
#>
Function Connect-OIM($AppServer = "localhost", $AppName = "AppServer", [PSCredential] $Credential , [switch] $useSSL) {

    # Creating URL string
    if ($useSSL ) {
        $url = "https://$AppServer/$AppName"
    }
    else {
        $url = "http://$AppServer/$AppName"
    }

    # Creating connection string
    if ($null -eq $Credential ) {
        #Single sign
        $authdata = @{AuthString = "Module=RoleBasedADSAccount" }
    }
    else {
        $user = $Credential.Username
        $Pass = $Credential.GetNetworkCredential().password
        $authdata = @{AuthString = "Module=DialogUser;User=$user;Password=$Pass" }

    }
    $authJson = ConvertTo-Json $authdata -Depth 2

    # Connecting
    Invoke-RestMethod -Uri "$url/auth/apphost" -Body $authJson.ToString() -Method Post -UseDefaultCredentials -Headers @{Accept = "application/json" } -SessionVariable session -AllowUnencryptedAuthentication | Out-Null

    Set-Variable -scope Global -name OIM_Session -Value $session
    Set-Variable -scope Global -name OIM_BaseURL -Value $url

    $session
}


Function Disconnect-OIM([WebRequestSession] $Session = (Get-Variable -scope Global -name OIM_Session)) {
    $url = Get-Variable -scope Global -name OIM_BaseURL -ErrorAction SilentlyContinue
    # Disconnect
    Invoke-RestMethod -Uri "$url/auth/logout" -WebSession $session -Method Post | Out-Null

}

Function Get-OIMObject($Object, $ObjectName, $Where, $OrderBy, $Session = (Get-Variable -scope Global -name OIM_Session), [switch]$First, [switch]$Full, $limit= 0, $offset =0) {

    # Read
    $url = Get-Variable -scope Global -name OIM_BaseURL -ErrorAction SilentlyContinue

    $dicBody = @{
        where = "$Where"
        limit = $limit
        offset = $offset
        }
    if (-not [string]::IsNullOrEmpty($OrderBy)) { $dicBody.Add("OrderBy", $OrderBy) }

    $body = $dicBody | ConvertTo-Json
    if( $null -ne $object.URI ){
        return Get-OIMObjectfromURI -uri $object.URI   -Session $Session

    }else{
        $result = Invoke-RestMethod -Uri "$url/api/entities/$($ObjectName)?loadType=ForeignDisplays" -WebSession $session -Method Post -Body $body -ContentType application/json
    }

    forEach ($item in $result) {

        if ($full) {
            $temp = Get-OIMObjectfromURI -uri $item.uri
            $temp

        }
        else {
            $temp = New-Object -TypeName PSObject -ArgumentList $item.Values
            $temp | Add-Member -Name uri -Value $item.uri -MemberType NoteProperty
            $temp | Add-Member -Name links -Value $item.Links -MemberType NoteProperty
            $temp | Add-Member -MemberType ScriptProperty -Name entity -Value {if($this.uri -match "(.*)/api/entity/(.*)/"){$matches[2]}else{"unknown"}}
            $temp | Add-Member -MemberType ScriptProperty -Name UID -Value {if($this.uri -match "(.*)/api/entity/(.*)/(.*)"){$matches[3]}else{"unknown"}}
            $temp
        }
        if ($first) { return }
    }

}

Function Get-OIMObjectfromURI($uri, $Session = (Get-Variable -scope Global -name OIM_Session)) {
    $url = Get-Variable -scope Global -name OIM_BaseURL -ErrorAction SilentlyContinue
    if ($uri -match "/(.*)(/api(.*))") {
        $uri = $matches[2]
        $itemfull = Invoke-RestMethod -Uri "$url/$uri" -WebSession $session -Method Get -ContentType application/json

        $temp = New-Object -TypeName PSObject -ArgumentList $itemfull.Values
        $temp | Add-Member -Name URI -Value $itemfull.uri -MemberType NoteProperty
        $temp | Add-Member -MemberType ScriptProperty -Name Entity -Value {if($this.uri -match "(.*)/api/entity/(.*)/(.*)"){$matches[2]}else{"unknown"}}
        $temp | Add-Member -MemberType ScriptProperty -Name UID -Value {if($this.uri -match "(.*)/api/entity/(.*)/(.*)"){$matches[3]}else{"unknown"}}
        $temp | Add-Member -Name links -Value $itemfull.Links -MemberType NoteProperty
        $temp
    }

}

Function Get-OIMURI($uri) {
    if ($uri -match "/(.*)(/api(.*))") {
        $matches[2]

    }

}



Function New-OIMObject{
    [CmdletBinding(SupportsShouldProcess)]
    param($ObjectName, [hashtable] $Properties, $Session = (Get-Variable -scope Global -name OIM_Session))

    $url = Get-Variable -scope Global -name OIM_BaseURL -ErrorAction SilentlyContinue
    # Read
    $body = @{values = $Properties } | ConvertTo-Json
    $item = Invoke-RestMethod -Uri "$url/api/entity/$($ObjectName)" -WebSession $session -Method Post -ContentType application/json  -Body $body
    if($PSCmdlet.ShouldProcess($ObjectName)){
        Get-OIMObjectfromURI -uri $item.uri -Session $session
    }
}

Function Remove-OIMObject {
    [CmdletBinding(SupportsShouldProcess)]
    Param($Object, $Session = (Get-Variable -scope Global -name OIM_Session))

    $url = Get-Variable -scope Global -name OIM_BaseURL -ErrorAction SilentlyContinue
    # Read
    $uri = Get-OIMURI $object.Uri
    if($PSCmdlet.ShouldProcess($uri)){
        Invoke-RestMethod -Uri "$url/$uri" -WebSession $session -Method Delete -ContentType application/json
    }
}

Function Update-OIMObject{
    [CmdletBinding(SupportsShouldProcess)]
    PAram($Object, [hashtable] $Properties, $Session = (Get-Variable -scope Global -name OIM_Session))

    # Read
    $url = Get-Variable -scope Global -name OIM_BaseURL -ErrorAction SilentlyContinue
    $uri = Get-OIMURI $object.Uri
    $body = @{values = $Properties } | ConvertTo-Json
    if($PSCmdlet.ShouldProcess("$url$uri")){
        Invoke-RestMethod -Uri "$url$uri" -WebSession $Session -Method Put -Body $body -ContentType application/json
    }
}

Function Add-OIMObjectMember{
    [CmdletBinding(SupportsShouldProcess)]
    PAram($TableName, $TableColumn, $UID, [Array] $Members, $Session = (Get-Variable -scope Global -name OIM_Session))

    # Read
    $body = @{members = $Members } | ConvertTo-Json
    $url = Get-Variable -scope Global -name OIM_BaseURL -ErrorAction SilentlyContinue
    if($PSCmdlet.ShouldProcess("$url$uri")){
        Invoke-RestMethod -Uri "$url/api/assignments/$TableName/$TableColumn/$UID"  -WebSession $Session -Method Post -Body $body -ContentType application/json
    }

}

Function Remove-OIMObjectMember{
    [CmdletBinding(SupportsShouldProcess)]
    PAram($TableName, $TableColumn, $UID, [Array] $Members, $Session = (Get-Variable -scope Global -name OIM_Session))

    # Read
    $body = @{members = $Members } | ConvertTo-Json
    $url = Get-Variable -scope Global -name OIM_BaseURL -ErrorAction SilentlyContinue

    if($PSCmdlet.ShouldProcess("$url$uri")){
        Invoke-RestMethod -Uri "$url/api/assignments/$TableName/$TableColumn/$UID"  -WebSession $Session -Method Delete -Body $body -ContentType application/json
    }
}


Function Start-OIMScript{
    [CmdletBinding(SupportsShouldProcess)]
    PAram($ScriptName, [array]$Parameters, $xObjectKey, $value, $Session = (Get-Variable -scope Global -name OIM_Session))

    # Read
    $body = @{parameters = $Parameters; Base = $xObjectKey; Value = $value } | ConvertTo-Json
    $body = @{parameters = $Parameters } | ConvertTo-Json
    $url = Get-Variable -scope Global -name OIM_BaseURL -ErrorAction SilentlyContinue

    if($PSCmdlet.ShouldProcess("$url$uri")){
        Invoke-RestMethod -Uri "$url/api/script/$ScriptName"  -WebSession $Session -Method put -Body $body -ContentType application/json
    }
}


Function Start-OIMMethod{
    [CmdletBinding(SupportsShouldProcess)]
    PAram($Object, $MethodName, [array]$Parameters, $Session = (Get-Variable -scope Global -name OIM_Session))

    # Read
    $url = Get-Variable -scope Global -name OIM_BaseURL -ErrorAction SilentlyContinue
    $body = @{parameters = $Parameters } | ConvertTo-Json
    $uri = Get-OIMURI $object.Uri
    if($PSCmdlet.ShouldProcess("$url$uri")){
        Invoke-RestMethod -Uri "$url/$uri/method/$MethodName"  -WebSession $Session -Method put -Body $body -ContentType application/json
    }
}


Function Start-OIMEvent{
    [CmdletBinding(SupportsShouldProcess)]
    PAram($Object, $EventName, [hashtable]$Parameters = @{}, $Session = (Get-Variable -scope Global -name OIM_Session))

    # Read
    $body = @{parameters = $Parameters } | ConvertTo-Json
    $uri = Get-OIMURI $object.Uri
    $url = Get-Variable -scope Global -name OIM_BaseURL -ErrorAction SilentlyContinue
    if($PSCmdlet.ShouldProcess("$url$uri")){
        Invoke-RestMethod -Uri "$url/$uri/event/$EventName"  -WebSession $Session -Method put -Body $body -ContentType application/json
    }
}


Function Set-OIMConfigParameter{
    [CmdletBinding(SupportsShouldProcess)]
    PAram($FullPath,$value)

    $obj =  Get-OIMObject -ObjectName DialogConfigParm -Where "FullPath = '$FullPath' "  -First -Full
    if($PSCmdlet.ShouldProcess("$url$uri")){
        Update-OIMObject -Object $obj -Properties @{Value = "$value"}
    }
}



Function Start-OIMSyncProject{
    [CmdletBinding(SupportsShouldProcess)]
    PAram($SyncName)
    #DPRShell

    $obj =  Get-OIMObject -ObjectName DialogSchedule -Where "displayname = '$SyncName "  -First -Full

    if($PSCmdlet.ShouldProcess("$url$uri")){
        Start-OIMEvent  -Object $obj -EventName run -Parameters @{}
    }
}



Function ConvertFrom-OIMDate([string] $Date){
    if(-not [string]::IsNullOrEmpty($date )){
        [datetime]::Parse( $Date)
    }
}

Function ConvertTo-OIMDate([DateTime] $Date){
    $Date.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ", [cultureinfo]::CurrentCulture)

}



<#

                $time = [datetime]::ParseExact( "2019-06-04T08:04:15.7500000Z","yyyy-MM-ddTHH:mm:ss.fff", [cultureinfo]::CurrentCulture)

ConvertTo-OIMDate(ConvertFrom-OIMDate("2019-06-04T08:04:15.7500000Z"))

200 Success
204 Success. No content returned.
401 Unauthorized. To use the One Identity Manager REST API, you first have to authenticate it against the application server.
404 Not found. The requested entity is not found.
405 Method not allowed. The HTTP request method that was specified is not the correct method for the request.
500 Internal server error. The error message is returned in the property error string of the response.
{
       "responseStatus": {
             "message": "Sample text"},
       "errorString": "Sample text",
       "exceptions": [{
             "number": 810017,
             "message": "Sample text"}
       ]
}

#>