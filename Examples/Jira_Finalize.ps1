# Initialize
Import-Module JiraCloud -force
Import-Module rExcel -force
# Variables
$SourceFolder = "O:\Map G-L\IAM - Project\07_Development"
$TemplatePath = Join-Path $sourceFolder "_Template\"

# Connect to Jira
If($credJira -eq $null){
    $credJira = Get-Credential 
    $JiraURL = "jumbo-supermarkten.atlassian.net"
    Connect-Jira -Server $JiraURL -credential $credJira | Out-Null
}

# Select issue
$issues = Find-JiraIssue -JQL "issuetype in (Bug, Story) AND project = IAMD AND assignee not in (EMPTY) and Sprint in openSprints()"
$selectedissues = $issues.issues | Select-Object key, {$_.Fields.Summary},id | Sort-Object -property id  | Out-GridView  -OutputMode Multiple 

ForEach($issue in $issues.issues){

    if($issue.key -in $selectedissues.key){
    # Get Values
        $currentissue = $issue
        $ID = $issue.key
        $name = $issue.fields.summary
        $folder = $id + " " + $name
        $folder = $folder -replace ":",""
        $description = $issue.fields.description            
        $destinationfolder = Join-Path $SourceFolder $folder 
        $documentationfolder = Join-Path  $destinationfolder "Documentation"
        $comments= Get-JiraComments $ID
        ForEach($comment in $comments.comments.body){
            if( $comment -match "(O:\\(.*).xlsx)"){

            }
            if( $comment -match "(O:\\(.*).docx)"){
                $file  = $Matches[1].Replace("\\","\")
                $
                if(Test-PAth $file){
                    Copy-Item -path $file -Destination  $documentationfolder     
                }else{
                    Write-warning "file not found: $file"
                }
               
            }

        }
        
    # Create folder
       # Check all documentatien
    
    }

}


