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
    $issues = Find-JiraIssue -JQL "issuetype in (Bug, Story) AND project = IAMD and Sprint in openSprints()"

    $selectedissues = $issues.issues | Select-Object key, {$_.Fields.Summary},id | Sort-Object -property id  | Out-GridView  -OutputMode Multiple 

    ForEach($issue in $issues.issues){
    
        if($issue.key -in $selectedissues.key){
        # Get Values
            $ID = $issue.key
            $name = $issue.fields.summary
            $folder = $id + " " + $name
            $folder = $folder -replace ":",""
            $description = $issue.fields.description            
            $destinationfolder = Join-Path $SourceFolder $folder 
        # Create folder
            Copy-Item -Path $TemplatePath -Destination   $destinationfolder -Container -Recurse
            REname-Item -Path $destinationfolder\IAMD-XXX.xlsx -NewName "$ID.xlsx" 
        
        # Update excel document
            $workbook = Open-rExcelWorkBook "$destinationfolder\$ID.xlsx" 
            Set-rExcelRangeField -Excel $global:excel -RangeName IAMD_ID -RangeValue $ID
            Set-rExcelRangeField -Excel $global:excel -RangeName IAMD_name -RangeValue $name
            Set-rExcelRangeField -Excel $global:excel -RangeName IAMD_description -RangeValue $description
            Save-rExcelWorkbook  $workbook -close

        # Create change label 
            #New-OIMObject -ObjectName DialogTag -Properties @{Ident_DialogTag=$issue.key;TagType="CHANGE";description=$issue.fields.summary}

        
        }

    }
    

# Close 
    Close-rExcel
