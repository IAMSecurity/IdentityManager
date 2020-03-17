Import-Module JiraCloud -force

If($credJira -eq $null){
    $credJira = Get-Credential 
    $JiraURL = "jumbo-supermarkten.atlassian.net"
    Connect-Jira -Server $JiraURL -credential $credJira
}

$URI_JiraCurrentStory = "$Global:JC_BaseURL/rest/api/2/search?jql=issuetype in (Bug, Story) AND project = IAMD AND assignee not in (EMPTY) and Sprint in openSprints() AND label = Release"
$issues = Invoke-RestMethod -Uri  $URI_JiraCurrentStory  -WebSession $Global:JC_session -Method GET  


$issues = Find-JiraIssue -JQL "issuetype in (Bug, Story) AND project = IAMD AND assignee not in (EMPTY) and Sprint in openSprints()"


ForEach($issue in $issues.issues){
    $issue.key
    $issue.fields.Status.Name
    New-OIMObject -ObjectName DialogTag -Properties @{Ident_DialogTag=$issue.key;TagType="CHANGE";description=$issue.fields.summary}

}

$issue = Get-JiraIssue -id IAMD-1249

([a-zA-Z]:(\\w+)*\\[a-zA-Z0_9]+)?.xlsx
([a-zA-Z]:(\\w+)*\\[a-zA-Z0_9]+)?.xlsx

