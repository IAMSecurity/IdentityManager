<#
.SYNOPSIS
   Connect to a One Identity Manager REST API server
.DESCRIPTION
    This command starts a connection to a One Identity Manager REST API server
.EXAMPLE
    C:\PS> Connect-OIM -AppServer xxx.company.local -UseSSL
    This example starts a connection to https://xxx.company.local/Appserver
#>
 
Function Invoke-OIMRestMethod{
    [CmdletBinding()]
    param (
        [Parameter()]
        $Uri,
        $Method = "Get",
        $body
    )
 
    $Parameters = @{
        Uri = $uri       
        Method= $Method
        websession= $Global:OIM_Session
        Headers = @{Accept = "application/json" }
    }
    if($null -ne $body){
        $JsonBody  = $body  | ConvertTo-Json
        $rawbody = [System.Text.Encoding]::UTF8.GetBytes($JsonBody)
 
        $Parameters.Add("ContentType", "application/json")
        $Parameters.Add("Body", $rawbody)
    }
    try {

        $APIResponse = Invoke-WebRequest @Parameters   -ErrorAction Stop
        ConvertFrom-Json -InputObject $APIResponse.Content
    } catch [System.UriFormatException] {

        #Catch URI Format errors. Likely $Script:BaseURI is not set; Connect-OIM should be run.
        $PSCmdlet.ThrowTerminatingError(

            [System.Management.Automation.ErrorRecord]::new(

                "$PSItem Run Connect-OIM{",
                $null,
                [System.Management.Automation.ErrorCategory]::NotSpecified,
                $PSItem

            )

        )

    } catch { 

        $ErrorID = $null
        $StatusCode = $($PSItem.Exception.Response).StatusCode.value__
        $ErrorMessage = $($PSItem.Exception.Message)

        $Response = $PSItem.Exception | Select-Object -ExpandProperty 'Response' -ErrorAction Ignore
        if ( $Response ) {

            $ErrorDetails = $($PSItem.ErrorDetails)
        }

        # Not an exception making the request or the failed request didn't have a response body.
        if ( $null -eq $ErrorDetails ) {

            throw $PSItem

        } Else {

            If (-not($StatusCode)) {

                #Generic failure message if no status code/response
                $ErrorMessage = "Error contacting $($PSItem.TargetObject.RequestUri.AbsoluteUri)"

            } ElseIf ($ErrorDetails) {

                try {

                    #Convert ErrorDetails JSON to Object
                    $Response = $ErrorDetails | ConvertFrom-Json

                    #API Error Message
                    $ErrorMessage = "[$StatusCode] $($Response.ErrorString)"

                    #API Error Code
                    $ErrorID = $Response.ErrorCode

                    #Inner error details are present
                    if ($Response.Details) {

                        #Join Inner Error Text to Error Message
                        $ErrorMessage = $ErrorMessage, $(($Response.Details | Select-Object -ExpandProperty ErrorMessage) -join ', ') -join ': '

                        #Join Inner Error Codes to ErrorID
                        $ErrorID = $ErrorID, $(($Response.Details | Select-Object -ExpandProperty ErrorCode) -join ',') -join ','

                    }

                } catch {

                    #If error converting JSON, return $ErrorDetails
                    #replace any new lines or whitespace with single spaces
                    $ErrorMessage = $ErrorDetails -replace "(`n|\W+)", ' '
                    #Use $StatusCode as ErrorID
                    $ErrorID = $StatusCode

                }
            }

        }

        #throw the error
        $PSCmdlet.ThrowTerminatingError(

            [System.Management.Automation.ErrorRecord]::new(

                $ErrorMessage,
                $ErrorID,
                [System.Management.Automation.ErrorCategory]::NotSpecified,
                $PSItem

            )

        )

    }
  
    
}
 
Function Connect-OIM{
    Param(
        [Parameter(ParameterSetName ='url',Mandatory=$true)]
        $Url,
        [PSCredential] $Credential,
        [switch] $AllowUnencryptedAuthentication

    )

    $Global:OIM_BaseURL = $url

    # Creating connection string
    $authdata = @{AuthString = "Module=RoleBasedADSAccount" }
    if ($null -ne $Credential ) {
        $user = $Credential.Username
        $Pass = $Credential.GetNetworkCredential().password
        $authdata = @{AuthString = "Module=DialogUser;User=$user;Password=$Pass" }

    }
    $authJson = ConvertTo-Json $authdata -Depth 2

    # Connecting
    Try{
        if($AllowUnencryptedAuthentication){
            Invoke-WebRequest -Uri "$url/auth/apphost" -Body $authJson.ToString() -Method Post -UseDefaultCredentials -Headers @{Accept = "application/json" } -SessionVariable session -AllowUnencryptedAuthentication 
        }else{
            Invoke-WebRequest -Uri "$url/auth/apphost" -Body $authJson.ToString() -Method Post -UseDefaultCredentials -Headers @{Accept = "application/json" } -SessionVariable session 
        }
        Set-Variable -Scope Global -Name OIM_Session -Value $session
        Set-Variable -Scope Global -Name OIM_BaseURL -Value $url
        
    }catch{
        Write-Error "OIM: Connection failed $($Global:OIM_BaseURL): $($_.Exception.Message)"
    }
}

Function Disconnect-OIM() { 

    Write-Verbose "Disconnect-OIM"
    $uri = "$Global:OIM_BaseURL/auth/logout" 
    Try{
        Invoke-OIMRestMethod -Uri $uri -Method Post | Out-Null
        $Global:OIM_Session = $null
        $Global:OIM_BaseURL = $null
    }catch{
        Write-Error "OIM: Disconnect failed: $($_.Exception.Message) ($uri)"
    }
}

Function Set-OIMGlobalVariable{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $Name,
        [Parameter(Mandatory=$true)]
        $Value
    )

    Write-Verbose "Set-OIMGlobalVariable: $Name Value: $Value"
    $body =  @{"Value"= $Value}   
    $uri = "$Global:OIM_BaseURL/appserver/variable/$Name"
    Try{
        Invoke-OIMRestMethod -Uri  $uri  -Method PUT -Body $body | Out-Null
    }catch{
        Write-Error "OIM Set GlobalVariable  failed: $($_.Exception.Message) ($uri)"
    }

}

Function Get-OIMObject{
    [CmdletBinding(SupportsPaging=$true,
        HelpUri="https://support.oneidentity.com/technical-documents/identity-manager/8.1/rest-api-reference-guide/7#TOPIC-1134467")]
    Param(
        [Parameter(Position=0,
            ParameterSetName="Object",
            Mandatory=$true,
            ValueFromPipeline=$true)]
        [Alias("Type","ObjectName")]
        $Object,        
        $Uid,
        [ValidateSet("Default","BulkReadOnly", "Slim", "ForeignDisplays","ForeignDisplaysForAllColumns")]
        $LoadType = "BulkReadOnly",
        $Where,
        $OrderBy,
        $displayColumns
    )
    Begin{
        $limit = $PSCmdlet.PagingParameters.First
        $offset = $PSCmdlet.PagingParameters.Skip
    }
    Process{
        If($object -is [array]){
            ForEach($item in $object){
                If(-not [string]::IsNullOrEmpty( $item.entity) -AND  -not [string]::IsNullOrEmpty($item.UID) ){
                    GEt-OIMObject -Object $item.entity -UID $item.UID  
                }
            }
        }

        If($null -ne $object.entity -and $null -ne $object.uid){
            $uid = $object.uid
            $object = $object.entity
        }
        if($object -is [string]  -and -not [string]::IsNullOrEmpty($object)){
            if(-not [string]::isnullOrEmpty($uid)){

                $Parameters = @{
                    uri = "$Global:OIM_BaseURL/api/entity/$($object)/$($Uid)"  
                }

            }else{

                $body = @{
                    where = $Where                
                    displayColumns = $displayColumns
                    }
                if (-not [string]::IsNullOrEmpty($OrderBy)) { $body.Add("OrderBy", $OrderBy) }   

                $Parameters = @{
                    uri = "$Global:OIM_BaseURL/api/entities/$($object)?loadType=$LoadType&limit=$limit&offset=$offset"

                    body = $body
                    Method="POST"
                }
            }

            Try{
                $OIMResponse = Invoke-OIMRestMethod @Parameters
            }catch{
                Write-Error "OIM: Get Object failed $uri - $($_.Exception.Message)"
            }


            ForEach ($item in $OIMResponse) {    
                If($item.uri -match "(.*)/api/entity/(.*)/(.*)"){
                    $uri = "$Global:OIM_BaseURL/api/entity/$($matches[2])/$($matches[3])"
                    $tmp = New-Object -TypeName PSObject -ArgumentList $item.Values
                    $tmp | Add-Member -Name uri     -Value $uri         -MemberType NoteProperty
                    $tmp | Add-Member -Name entity  -Value $matches[2]  -MemberType NoteProperty
                    $tmp | Add-Member -Name UID     -Value $matches[3]  -MemberType NoteProperty
                    $tmp
                }else{
                    Write-Warning "OIM: Uri not as expected "
                }
            }
        }
    }
    End{}
}

Function New-OIMObject{
    [CmdletBinding(SupportsShouldProcess)]
    Param(
        [Parameter(Mandatory=$true)]
        [Alias("Type","Object")]
            $ObjectName,
        [Parameter(Mandatory=$true)]
            [hashtable] $Properties,
            [switch] $checkexists)
    Begin{}
    Process{
        $Uri = "$Global:OIM_BaseURL/api/entity/$($ObjectName)"
        $body = @{values = $Properties }

        If($checkexists -and $Properties.Containskey("Ident_$objectname")){
            $OIMResponse = GEt-OIMObject -object $objectname -where "Ident_$objectname = '$($Properties["Ident_$objectname"])'"
            if($null -ne  $OIMResponse){
                return $OIMResponse
            }
        }
        Try{
            if ($PSCmdlet.ShouldProcess($Object.uri , "Create object")) {  
                $OIMResponse = Invoke-OIMRestMethod -Uri $Uri  -Method Post -Body $body
                if($null -ne  $OIMResponse){
                    Get-OIMObject -Object $ObjectName -uid  $OIMResponse.uid
                }
            }else{
                Write-Warning "Should create object $ObjectName Value = $($Properties | ConvertTo-json )"
            }
        }catch{
            Write-Error "failed $ObjectName - $($Properties.Values -join ";") - $($_.Exception.Message)"
        }
    }
    End{}
}

Function Remove-OIMObject{
    [CmdletBinding(SupportsShouldProcess)]
    Param(
        [Parameter(ValueFromPipeline=$true, Mandatory=$true)]
            $Object

        )
    Begin{}
    Process{
        # Read
        if($null -eq $object ){Return }
        if($object -is [System.Array]){

            ForEach($ChildObject in $Object){
                if ($PSCmdlet.ShouldProcess($ChildObject.uri , "removing object")) {
                    if($null -ne  $ChildObject.uri){
                        Invoke-OIMRestMethod -Uri $ChildObject.uri -Method Delete  | Out-Null
                    }
                }
            }

        }else{
            if ($PSCmdlet.ShouldProcess($Object.uri , "removing object")) {
                if($null -ne  $Object.uri){
                    Invoke-OIMRestMethod -Uri $Object.uri -Method Delete  | Out-Null
                }
            }else{
                Write-Warning "Should run removing object: $($Object.uri)"
            }
        }
    }
    End{}
}

Function Update-OIMObject{
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
            $Object,
        [Parameter(Mandatory=$true)]
            [hashtable] $Properties
    )

    Begin{}
    Process{
        $body = @{values = $Properties }
        if(-not ($object -is [array])){
            $list = @($object)
        }
        ForEach($item in $object){
            if ($PSCmdlet.ShouldProcess($item.uri , "Update Object")) {  
                Invoke-OIMRestMethod -Uri $item.Uri -Method Put -Body $body  |Out-Null
                Get-OIMObject -Object $item.entity -Uid  $item.UID 
            }else{
                Write-Warning "Should update object $($item.Uri ) Value = $($Properties | ConvertTo-json )"
            }
 
        }
    }
    End{}
}

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
    $uri = "$Global:OIM_BaseURL/api/assignments/$TableName/$TableColumn/$UID"
    Invoke-OIMRestMethod -Uri $uri  -Method Post -Body $body
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
    $uri = "$Global:OIM_BaseURL/api/assignments/$TableName/$TableColumn/$UID"
    Invoke-OIMRestMethod -Uri  $uri  -Method Delete -Body $body
 
}
 
Function Start-OIMScript{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]  
            $ScriptName,
        [array]$Parameters
    )
    $body = @{parameters = $Parameters }
    $uri = "$Global:OIM_BaseURL/api/script/$ScriptName"
    Invoke-OIMRestMethod -Uri $uri  -Method put -Body $body  | out-Null
}  
    
Function Start-OIMMethod{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true,Mandatory=$true)]  
            $Object,
        [Parameter(Mandatory=$true)]  
            $MethodName,
            [array]$Parameters = @()
    )
    # Read
    $body = @{parameters = $Parameters }
    
    if($null -ne $Object.Uri -and $null -ne $MethodName){
        $uri = "$($Object.Uri)/method/$MethodName"
        Invoke-OIMRestMethod -Uri $uri -Method put -Body $body  | out-Null
    }
}
 
Function Start-OIMEvent{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true,Mandatory=$true)]  
            $Object,
        [Parameter(Mandatory=$true)]  
            $EventName,
            [hashtable]$Parameters = @{},
            [switch]$wait
    )
    # Read
    Begin{}
    Process{
        Write-VErbose "Start-OIMEvent Start "
        if($null -eq $object){return}
        If($object.uri -is [array]){
 
            Write-VErbose "Start-OIMEvent Array "
            ForEach($item in $object){
                Start-OIMEvent -Object $item -eventname $EventName -parameters $Parameters
            }
        }else{
            Write-VErbose "Start-OIMEvent object. "
            $body = @{parameters = $Parameters }
            if($null -ne $Object.Uri -and $null -ne $EventName){
                $uri = "$($Object.Uri)/event/$EventName"
                Invoke-OIMRestMethod -Uri   $uri -Method put -Body $body | out-Null
            }
        }
        if($wait){
            Wait-OIMJobQueue -JobChainName $EventName
        }
    }
    End{}
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
 
 
<# Helper functions #>
Function ConvertFrom-OIMDate{
    [CmdletBinding()]
    param (
        [Parameter()]
            [string] $Date
    )
 
    if(-not [string]::IsNullOrEmpty($date )){
        [datetime]::Parse( $Date)
    }
}
 
Function ConvertTo-OIMDate{
    [CmdletBinding()]
    param (
        [Parameter()]
            [DateTime] $Date
    )
    $Date.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ", [cultureinfo]::CurrentCulture)
 
}
 
Function Wait-OIMJobQueue{
    [CmdletBinding()]
    param (
        [Parameter()]
        $JobChainName,
        $timeout = 300
    )
    $jobs = ""
    $remainingtime = $timeout
 
    Write-Progress "Wait jobqueue with jobchainname: '$JobChainName'"
    $where = "JobChainName LIKE '%" + $JobChainName +"%' AND Ready2EXE <> 'HISTORY' AND Ready2exe <> 'FINISHED'"
 
    While($jobs -ne $null -and $remainingtime -ge 0){
        Start-Sleep -Seconds 3
        $remainingtime = $remainingtime -3
        
    Write-Progress "Wait jobqueue with jobchainname:'$JobChainName' jobscount:$( $jobs.count) seconds remaining:$remainingtime" -PercentComplete ((($timeout-$remainingtime)/$timeout )*100)
        $jobs =  Get-OIMObject -ObjectName "JobQueue" -Where  $where
   
    }
}