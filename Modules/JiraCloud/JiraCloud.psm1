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
Function Connect-Jira($Server, [PSCredential] $Credential ) {
    
    # Creating URL string
        $Global:JC_BaseURL = "https://$Server"
        $Headers =  @{
                    Accept = "application/json"
                }

    # Creating connection string
    if ($null -eq $Credential ) {
        #Single sign
        
    }
    else {        

        $AuthString = $Credential.UserName + ":" +  $Credential.GetNetworkCredential().password
        $Bytes = [System.Text.Encoding]::UTF8.GetBytes($AuthString)
        $EncodedText = [Convert]::ToBase64String($Bytes)
        $Headers.Add("Authorization","Basic $($EncodedText)")
    }
    # Connecting
    Invoke-RestMethod -Uri "$Global:JC_BaseURL/rest/api/3/myself" -Method GET -UseDefaultCredentials -Headers $Headers -SessionVariable session  | Out-Null

    $Global:JC_Session = $session
    $session 
}

Function Find-JiraIssue($JQL, $Session = $global:JC_Session) {

    # Read 

      Invoke-RestMethod -Uri "$Global:JC_BaseURL/rest/api/2/search?jql=$JQL" -WebSession $session -Method GET  -AllowUnencryptedAuthentication
   

}

Function Get-JiraIssue($ID, $Session = $global:JC_Session) {

    # Read 

      Invoke-RestMethod -Uri "$Global:JC_BaseURL/rest/api/2/issue/$ID" -WebSession $session -Method GET  -AllowUnencryptedAuthentication
   

}

Function Get-JiraComments($ID, $Session = $global:JC_Session) {

    # Read 

      Invoke-RestMethod -Uri "$Global:JC_BaseURL/rest/api/2/issue/$ID/comment" -WebSession $session -Method GET  -AllowUnencryptedAuthentication
   

}



<# Confluence download PDF 
 $Jira_IAMD1189  = Get-JCIssueObject -ID IAMD-1189 
    $URI_JiraCreatePDF = "https://jumbo-supermarkten.atlassian.net/wiki/spaces/flyingpdf/pdfpageexport.action?pageId=882934558"
    $URI_JiraWaitPDF = "https://jumbo-supermarkten.atlassian.net/wiki/spaces/flyingpdf/runningtaskxml.action?taskId=924811880"



    $a = Invoke-WebRequest -Uri $URI_JiraWaitPDF-WebSession $Global:JC_session -Method GET  


    $a.Content  -match "([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})"
    $matches
    $URI_JIRAPDF = "https://jumbo-supermarkten.atlassian.net/wiki/download/temp/filestore/{}"
    Invoke-WebRequest -Uri $URI_JIRAPDF -WebSession $Global:JC_session -OutFile test.pdf

#>

