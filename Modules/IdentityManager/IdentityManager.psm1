

Function Connect-OIM($AppServer,$AppName = "AppServer", $Credential , [switch] $useSSL){
    
    # Creating URL string
        if($useSSL ){
            $Global:OIM_BaseURL = "https://$AppServer/$AppName"
        }else{
            $Global:OIM_BaseURL = "http://$AppServer/$AppName"
        }

    # Creating connection string
        if($Credential -eq $null){ #Single sign
            $authdata = @{AuthString="Module=RoleBasedADSAccount"}
        }else{
            $user = $Credential.Username
            $Pass = $Credential.GetNetworkCredential().password
            $authdata = @{AuthString="Module=DialogUser;User=$user;Password=$Pass"}

        }

        $authJson = ConvertTo-Json $authdata -Depth 2

    # Connecting
        Invoke-RestMethod -Uri "$Global:OIM_BaseURL/auth/apphost" -Body $authJson.ToString() -Method Post -UseDefaultCredentials -Headers @{Accept="application/json"} -SessionVariable session | Out-Null

        $Global:OIM_Session = $session
        $session 
}

Function Disconnect-OIM($Session){
  
    # Disconnect
        if($null -ne $session ){
            Invoke-RestMethod -Uri "$Global:OIM_BaseURL/auth/logout" -WebSession $session -Method Post| Out-Null

        }elseif($null -ne $Global:OIM_Session){
            Invoke-RestMethod -Uri "$Global:OIM_BaseURL/auth/logout" -WebSession $Global:OIM_Session -Method Post | Out-Null

        }

}

Function Get-OIMObject($ObjectName,$Where,$OrderBy,$Session=$Global:OIM_Session,[switch]$First,[switch]$Full){

    # Read 

        $dicBody = @{where="$Where"}
        if(-not [string]::IsNullOrEmpty($OrderBy)){$dicBody.Add("OrderBy",$OrderBy)}

        $body =  $dicBody | ConvertTo-Json
        $result = Invoke-RestMethod -Uri "$Global:OIM_BaseURL/api/entities/$($ObjectName)?loadType=ForeignDisplays" -WebSession $session -Method Post -Body $body -ContentType application/json

        
        forEach($item in $result){
           
            if($full){
                $temp = Get-OIMObjectfromURI -uri $item.uri 
                $temp
            
            }else{
                $temp = New-Object -TypeName PSObject -ArgumentList $item.Values 
                $temp | Add-Member -Name uri -Value $item.uri -MemberType NoteProperty
                $temp
            }
            if($first){return}
        }

}

Function Get-OIMObjectfromURI($uri,$Session=$Global:OIM_Session){
            if($uri-match "/(.*)(/api(.*))"){
                    $uri = $matches[2]
                    $itemfull = Invoke-RestMethod -Uri "$Global:OIM_BaseURL/$uri" -WebSession $session -Method Get -ContentType application/json  
                
                    $temp = New-Object -TypeName PSObject -ArgumentList $itemfull.Values 
                    $temp | Add-Member -Name uri -Value $itemfull.uri -MemberType NoteProperty
                    $temp
                }

}

Function Get-OIMURI($uri){
            if($uri-match "/(.*)(/api(.*))"){
                    $matches[2]
                    
                }

}



Function New-OIMObject($ObjectName, [hashtable] $Properties,$Session=$Global:OIM_Session){

    # Read 
    $body = @{values=$Properties} | ConvertTo-Json     
    $item = Invoke-RestMethod -Uri "$Global:OIM_BaseURL/api/entity/$($ObjectName)" -WebSession $session -Method Post -ContentType application/json  -Body $body 
      
    Get-OIMObjectfromURI -uri $item.uri -Session $session

   
              

}

Function Remove-OIMObject($Object, $Session=$Global:OIM_Session){

    # Read 
    $uri = Get-OIMURI $object.Uri     
    Invoke-RestMethod -Uri "$Global:OIM_BaseURL/$uri" -WebSession $session -Method Delete -ContentType application/json
     
              

}

Function Update-OIMObject($Object, [hashtable] $Properties,$Session=$Global:OIM_Session){

    # Read 
    $uri = Get-OIMURI $object.Uri
    $body=@{values=$Properties} | ConvertTo-Json
    Invoke-RestMethod -Uri $uri -WebSession $Session -Method Put -Body $body -ContentType application/json
              

}

Function Add-OIMObjectMember($TableName,$TableColumn,$UID, [Array] $Members,$Session=$Global:OIM_Session){

    # Read 
    $body=@{members=$Members} | ConvertTo-Json
    Invoke-RestMethod -Uri "$Global:OIM_BaseURL/api/assignments/$TableName/$TableColumn/$UID"  -WebSession $Session -Method Post -Body $body -ContentType application/json
              

}

Function Remove-OIMObjectMember($TableName,$TableColumn,$UID, [Array] $Members,$Session=$Global:OIM_Session){

    # Read 
    $body=@{members=$Members} | ConvertTo-Json
    Invoke-RestMethod -Uri "$Global:OIM_BaseURL/api/assignments/$TableName/$TableColumn/$UID"  -WebSession $Session -Method Delete -Body $body -ContentType application/json
              

}


Function Start-OIMScript($ScriptName,[array]$Parameters,$xObjectKey,$value,$Session=$Global:OIM_Session){

    # Read 
    $body=@{parameters=$Parameters;Base=$xObjectKey;Value=$value} | ConvertTo-Json
    $body=@{parameters=$Parameters} | ConvertTo-Json

   
  Invoke-RestMethod -Uri "$Global:OIM_BaseURL/api/script/$ScriptName"  -WebSession $Session -Method put -Body $body -ContentType application/json
              

}


Function Start-OIMMethod($Object,$MethodName,[array]$Parameters,$Session=$Global:OIM_Session){

    # Read 
    $body=@{parameters=$Parameters} | ConvertTo-Json
        $uri = Get-OIMURI $object.Uri
    "$Global:OIM_BaseURL/ $uri /method/$MethodName"
    Invoke-RestMethod -Uri "$Global:OIM_BaseURL/$uri/method/$MethodName"  -WebSession $Session -Method put -Body $body -ContentType application/json
              

}


Function Start-OIMEvent($Object,$EventName,[hashtable]$Parameters,$Session=$Global:OIM_Session){

    # Read 
    $body=@{parameters=$Parameters} | ConvertTo-Json
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


$temp = Invoke-RestMethod -Uri "$Global:OIM_BaseURL/api/script/CCC_GetHostNameFromSystem"  -WebSession $Global:OIM_Session -Method put -Body $body -ContentType application/json -InformationVariable $iv -ErrorVariable $ev -WarningVariable $wv  
 

    Start-OIMScript -ScriptName CCC_GetHostNameFromSystem -Parameters @("sasfdsa")
    $con = Connect-OIM -AppServer SBX-IAM-9001.sandbox.local -AppName D1IMAppServer -Credential $cred 
    Get-OIMObject -ObjectName Person -Where "Lastname like 'Lo%' "   -First -Full
    Get-OIMObjectfromURI -uri "/D1IMAppServer/api/entity/Person/604b2bad-a34c-4c58-a0d8-6ea86e61ba5c" 
    New-OIMObject -ObjectName Person -Properties @{Firstname="test";Lastname="test"}
    Add-OIMObjectMember -TableName PersonInOrg -TableColumn UID_Person -UID xxx-xxx -Properties @(aaaa,ssss,ddd
    Remove-OIMObjectMember -TableName PersonInOrg -TableColumn UID_Person -UID xxx-xxx -Properties @(aaaa,ssss,ddd)
    Remove-OIMObject $obj
    Start-OIMScript -ScriptName QER_GetWebBaseURL
    Start-OIMMethod -Object $obj -MethodName ExecuteTemplates -Parameters @()
    Start-OIMEvent  -Object $obj -EventName ExecuteTemplates -Parameters @{}
    Disconnect-OIM $con



    #DEV connection 
    $cred = Get-Credential
    $con = Connect-OIM -AppServer 192.168.56.101 -AppName D1IMAppServer -Credential $cred 
    $obj = Get-OIMObject -ObjectName Person -Where "centralaccount = 'test' "   -First -Full
    Start-OIMMethod -Object $obj -MethodName ExecuteTemplate -Parameters @("UID_ORg")

    Get-OIM
#>
