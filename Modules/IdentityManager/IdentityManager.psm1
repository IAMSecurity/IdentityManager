
<#
.SYNOPSIS
   Connect to a One Identity Manager REST API server
.DESCRIPTION
    This command starts a connection to a One Identity Manager REST API server
.EXAMPLE
    C:\PS> Connect-OIM -AppServer xxx.company.local -UseSSL
    This example starts a connection to https://xxx.company.local/Appserver
#>
Function Connect-OIM($AppServer, $AppName = "AppServer", [PSCredential] $Cred , [switch] $useSSL) {
    
    # Creating URL string
    if ($useSSL ) {
        $Global:OIM_BaseURL = "https://$AppServer/$AppName"
    }
    else {
        $Global:OIM_BaseURL = "http://$AppServer/$AppName"
    }

    # Creating connection string
    if ($null -eq $Cred ) {
        #Single sign
        $authdata = @{AuthString = "Module=RoleBasedADSAccount" }
    }
    else {
        $user = $Cred.Username
        $Pass = $Cred.GetNetworkCredential().password
        $authdata = @{AuthString = "Module=DialogUser;User=$user;Password=$Pass" }

    }
    $authJson = ConvertTo-Json $authdata -Depth 2

    # Connecting
    Invoke-RestMethod -Uri "$Global:OIM_BaseURL/auth/apphost" -Body $authJson.ToString() -Method Post -UseDefaultCredentials -Headers @{Accept = "application/json" } -SessionVariable session | Out-Null

    $Global:OIM_Session = $session
    $session 
}


Function Disconnect-OIM([WebRequestSession] $Session = $Global:OIM_Session) {
  
    # Disconnect
    Invoke-RestMethod -Uri "$Global:OIM_BaseURL/auth/logout" -WebSession $session -Method Post | Out-Null

}

Function Get-OIMObject($ObjectName, $Where, $OrderBy, $Session = $Global:OIM_Session, [switch]$First, [switch]$Full) {

    # Read 

    $dicBody = @{where = "$Where" }
    if (-not [string]::IsNullOrEmpty($OrderBy)) { $dicBody.Add("OrderBy", $OrderBy) }

    $body = $dicBody | ConvertTo-Json
    $result = Invoke-RestMethod -Uri "$Global:OIM_BaseURL/api/entities/$($ObjectName)?loadType=ForeignDisplays" -WebSession $session -Method Post -Body $body -ContentType application/json

        
    forEach ($item in $result) {
           
        if ($full) {
            $temp = Get-OIMObjectfromURI -uri $item.uri 
            $temp
            
        }
        else {
            $temp = New-Object -TypeName PSObject -ArgumentList $item.Values 
            $temp | Add-Member -Name uri -Value $item.uri -MemberType NoteProperty
            $temp
        }
        if ($first) { return }
    }

}

Function Get-OIMObjectfromURI($uri, $Session = $Global:OIM_Session) {
    if ($uri -match "/(.*)(/api(.*))") {
        $uri = $matches[2]
        $itemfull = Invoke-RestMethod -Uri "$Global:OIM_BaseURL/$uri" -WebSession $session -Method Get -ContentType application/json  
                
        $temp = New-Object -TypeName PSObject -ArgumentList $itemfull.Values 
        $temp | Add-Member -Name uri -Value $itemfull.uri -MemberType NoteProperty
        $temp
    }

}

Function Get-OIMURI($uri) {
    if ($uri -match "/(.*)(/api(.*))") {
        $matches[2]
                    
    }

}



Function New-OIMObject($ObjectName, [hashtable] $Properties, $Session = $Global:OIM_Session) {

    # Read 
    $body = @{values = $Properties } | ConvertTo-Json     
    $item = Invoke-RestMethod -Uri "$Global:OIM_BaseURL/api/entity/$($ObjectName)" -WebSession $session -Method Post -ContentType application/json  -Body $body 
      
    Get-OIMObjectfromURI -uri $item.uri -Session $session              

}

Function Remove-OIMObject($Object, $Session = $Global:OIM_Session) {

    # Read 
    $uri = Get-OIMURI $object.Uri     
    Invoke-RestMethod -Uri "$Global:OIM_BaseURL/$uri" -WebSession $session -Method Delete -ContentType application/json 

}

Function Update-OIMObject($Object, [hashtable] $Properties, $Session = $Global:OIM_Session) {

    # Read 
    $uri = Get-OIMURI $object.Uri
    $body = @{values = $Properties } | ConvertTo-Json
    Invoke-RestMethod -Uri $uri -WebSession $Session -Method Put -Body $body -ContentType application/json

}

Function Add-OIMObjectMember($TableName, $TableColumn, $UID, [Array] $Members, $Session = $Global:OIM_Session) {

    # Read 
    $body = @{members = $Members } | ConvertTo-Json
    Invoke-RestMethod -Uri "$Global:OIM_BaseURL/api/assignments/$TableName/$TableColumn/$UID"  -WebSession $Session -Method Post -Body $body -ContentType application/json
              

}

Function Remove-OIMObjectMember($TableName, $TableColumn, $UID, [Array] $Members, $Session = $Global:OIM_Session) {

    # Read 
    $body = @{members = $Members } | ConvertTo-Json
    Invoke-RestMethod -Uri "$Global:OIM_BaseURL/api/assignments/$TableName/$TableColumn/$UID"  -WebSession $Session -Method Delete -Body $body -ContentType application/json
              

}


Function Start-OIMScript($ScriptName, [array]$Parameters, $xObjectKey, $value, $Session = $Global:OIM_Session) {

    # Read 
    $body = @{parameters = $Parameters; Base = $xObjectKey; Value = $value } | ConvertTo-Json
    $body = @{parameters = $Parameters } | ConvertTo-Json

   
    Invoke-RestMethod -Uri "$Global:OIM_BaseURL/api/script/$ScriptName"  -WebSession $Session -Method put -Body $body -ContentType application/json
              

}


Function Start-OIMMethod($Object, $MethodName, [array]$Parameters, $Session = $Global:OIM_Session) {

    # Read 
    $body = @{parameters = $Parameters } | ConvertTo-Json
    $uri = Get-OIMURI $object.Uri
    "$Global:OIM_BaseURL/ $uri /method/$MethodName"
    Invoke-RestMethod -Uri "$Global:OIM_BaseURL/$uri/method/$MethodName"  -WebSession $Session -Method put -Body $body -ContentType application/json
              

}


Function Start-OIMEvent($Object, $EventName, [hashtable]$Parameters, $Session = $Global:OIM_Session) {

    # Read 
    $body = @{parameters = $Parameters } | ConvertTo-Json
    $uri = Get-OIMURI $object.Uri
    "$Global:OIM_BaseURL/$uri/event/$EventName" 
    Invoke-RestMethod -Uri "$Global:OIM_BaseURL/$uri/event/$EventName"  -WebSession $Session -Method put -Body $body -ContentType application/json
              

}

<#
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