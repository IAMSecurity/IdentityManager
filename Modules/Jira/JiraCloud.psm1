<#
$cred = Get-Credential 
$JiraURL = "jumbo-supermarkten.atlassian.net"
 Connect-JC -Server $JiraURL -Credential $cred
 Get-JCIssueObject -ID IAMD-1189 


#>



<#
.SYNOPSIS
   Connect to a Jira Cloud REST API server
.DESCRIPTION
    This command starts a connection to a One Identity Manager REST API server
.EXAMPLE
    C:\PS> Connect-OIM -AppServer xxx.company.local -UseSSL

#>
Function Connect-JC($Server, [PSCredential] $Credential ) {
    
    # Creating URL string
        $Global:JC_BaseURL = "https://$Server"
        $Headers =  @{
                    Accept = "application/json"
                }

    # Creating connection string
    if ($null -eq $Cred ) {
        #Single sign
        
    }
    else {        

        $AuthString = $username + ":" +  $password
        $Bytes = [System.Text.Encoding]::UTF8.GetBytes($AuthString)
        $EncodedText = [Convert]::ToBase64String($Bytes)
        $Headers.Add("Authorization","Basic $($EncodedText)")
    }
    # Connecting
    Invoke-RestMethod -Uri "$Global:JC_BaseURL/rest/api/3/myself" -Method GET -UseDefaultCredentials -Headers $Headers -SessionVariable session | Out-Null

    $Global:JC_Session = $session
    $session 
}


Function Get-JCIssueObject($ID, $Session = $global:JC_Session) {

    # Read 

      Invoke-RestMethod -Uri "$Global:JC_BaseURL/rest/api/2/issue/$ID" -WebSession $session -Method GET  
   

}

