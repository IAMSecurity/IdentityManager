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

Function Get-ConfluenceSpaces($Session = $global:JC_Session){

    $results = Invoke-RestMethod -Uri "$Global:JC_BaseURL/wiki/rest/api/space" -WebSession $session -Method GET  -AllowUnencryptedAuthentication
   
    $results.Results
}
Function Get-ConfluencePage($title, $spaceId, $Session = $global:JC_Session){
    $dicbody = @{
        Title = $title
        Type = "page"
        SpaceKey = $spaceId
        
        Status = "current"
       
    }
    $body = ConvertTo-Json $dicbody 
    $body.toString()
  Invoke-RestMethod -Uri "$Global:JC_BaseURL/wiki/rest/api/content" -Body $body.toString() -WebSession $session  -Method GET  -AllowUnencryptedAuthentication


  
 
}

Function New-ConfluencePage ($id = "",$spaceId,$ancestorsID, $Title, $Session = $global:JC_Session) {


    # Read 
        $dicbody = @{
            id = "5200"
            title = $title
            type = "page"
            space = @{
                key = $spaceId
            }
            status = "current"
            ancestors =@(@{
                
                id = $ancestorsID
            })
            body = @{
                view = @{
                    value = "<P>test</P>"
                    representation = "view"
                }
            }
        }
        $headers =  @{
            Accept = "application/json" 
        "Content-Type" ="application/json" }

        $body = ConvertTo-Json $dicbody 
        $body.toString()
      Invoke-RestMethod -Uri "$Global:JC_BaseURL/wiki/rest/api/content" -Body $body.toString() -WebSession $session -Headers $headers -Method POST  -AllowUnencryptedAuthentication
   

}

Function Add-ConfluenceAttachement($id,$file, $Session = $global:JC_Session){

    $results = Invoke-RestMethod -Uri "$Global:JC_BaseURL/wiki/rest/api/content/$id/child/attachment" -WebSession $session -Method POST -Body [convert]::ToBase64String((get-content  $file -Encoding  ))   -AllowUnencryptedAuthentication
   
    $results.Results
}
<#
$id = 963871283
$file = "\\jumbo.local\HKT-UsrFolders$\roblooman\Desktop\Firewall\FirewallAanpassing_IAM-ABX_SAP-ABXv2.xlsx"
Add-ConfluenceAttachement -Id 963871283 -File $file

963871283
New-ConfluencePage -SpaceID I -AncestorsID 923861135 -title "Test"

#>


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

