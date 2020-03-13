Import-Module Jira
$cred = Get-Credential 
$JiraURL = "jumbo-supermarkten.atlassian.net"
 Connect-JC -Server $JiraURL -credential $cred
 Get-JCIssueObject -ID IAMD-1189 
 #rob.looman@jumbo.com


$issues = Invoke-RestMethod -Uri "$Global:JC_BaseURL/rest/api/2/search?jql=issuetype in (Bug, Story) AND project = IAMD AND assignee not in (EMPTY) and Sprint in openSprints()"-WebSession $Global:JC_session -Method GET  
ForEach($issue in $issues.issues){
    $issue.key
    #New-OIMObject -ObjectName DialogTag -Properties @{Ident_DialogTag=$issue.key;TagType="CHANGE";description=$issue.fields.summary}

}
 $Global:JC_Headers =  @{               Accept = "application/json"               }

$issues = Invoke-WebRequest -Uri "https://jumbo-supermarkten.atlassian.net/wiki/spaces/flyingpdf/pdfpageexport.action?pageId=882934558"-WebSession $Global:JC_session -Method GET  -Headers @{               Accept = "application/json"               }

https://jumbo-supermarkten.atlassian.net/wiki/spaces/flyingpdf/pdfpageexport.action?pageId=882934558

Invoke-WebRequest -Uri "https://jumbo-supermarkten.atlassian.net/wiki/spaces/flyingpdf/runningtaskxml.action?taskId=924811880"-WebSession $Global:JC_session -Method GET  -Headers @{               Accept = "application/json"               }
$a = Invoke-WebRequest -Uri "https://jumbo-supermarkten.atlassian.net/wiki/spaces/flyingpdf/runningtaskxml.action?taskId=924811880"-WebSession $Global:JC_session -Method GET  -Headers @{               Accept = "application/json"               }

 $a.Content  -match "([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})"
 $matches

Invoke-WebRequest -Uri "https://jumbo-supermarkten.atlassian.net/wiki/download/temp/filestore/5e303eea-0dd4-4ca0-9f3b-d30b291249a4" -WebSession $Global:JC_session -OutFile test.pdf

spaces/flyingpdf/pdfpageexport.action?pageId=882934558
