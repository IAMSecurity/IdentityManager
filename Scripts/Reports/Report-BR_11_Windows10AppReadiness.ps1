    . ..\..\data\SNSBankNV_WPS.ps1



$csv = Import-Csv $OutputPathRESObjects -UseCulture
$applicationlist = $csv |?{$_.RESTypeName -eq "application"}

$dicGLobalGroups = @{}
$i = 1 
$total = $applicationlist.Count

ForEach($application in $applicationlist  ){
    Write-Progress -Activity "Processing application list" -PercentComplete (($i /$total)*100)
    $i++
   if($application.ac_grouplist -eq $null){
    continue
   }
   <#
   
    #>
    $dicGLobalGroupsv2 = Get-ADGroup -filter {SamAccountName -like "GAP*"} -Properties info | Group-Object -AsHashTable -Property NAme

    ForEach($DLgroup in  $application.ac_grouplist.Split("#")){
           if($DLgroup -eq $null ){
            continue
           }
           #$application.Name
            #$application.Description
            $arDLGroup = $DLgroup.split("\")
            
            if($arDLGroup[1].Length  -le 7 ){
            #$DLgroup
           # continue
           }
            $SAM = $arDLGroup[1] 

             if([string]::IsNullOrEmpty( $SAM) ){
            #$DLgroup
                continue
           }
            ForEach($GGroup in (Get-ADGroup  -filter {SamAccountName -eq  $SAM }  -Properties members -ErrorAction SilentlyContinue).Members ){
                if($GGroup -Match "CN=(.*?),(.*)"){
                    if($dicGLobalGroups.ContainsKey($matches[1])){
                        $tmp = $dicGLobalGroups[$matches[1]]
                        if(-not $tmp.Contains($application)){
                            $tmp += $application
                        }
                        $dicGLobalGroups[$matches[1]] =  $tmp 
                    }elsE{
                        $tmp = @()
                        $tmp += $application
                        $dicGLobalGroups.Add($matches[1],$tmp)
                    }
                }
                

            }
    }

}


$users = Get-ADUser -filter * -SearchBase "OU=Users,OU=Organisatie,DC=VERz,DC=local" -SearchScope OneLevel -Properties memberof,businesscategory
#$users = Get-ADUser -Identity 9404463 -Properties memberof,businesscategory
$listResult = @()

$i = 1 
$total = $users.Count


ForEach($user in $users){

    Write-Progress -Activity "Processing User list" -Status "$($user.Samaccountname)" -PercentComplete (($i /$total)*100)
    $i++
    ForEach($group in $user.MemberOf){
        $found = $false 
        if($group -Match "CN=(.*?),(.*)"){
            if($dicGLobalGroups.ContainsKey($matches[1])){
               #$dicGLobalGroups[$matches[1]].administrativenote
               $GLGroup = $matches[1]
               ForEach($DLGroup in $dicGLobalGroups[$matches[1]]){
                   if([string]::IsNullOrEmpty($DLGroup.administrativenote)){
                    $adminnote = ""
                   }else{
                    $adminnote = $DLGroup.administrativenote.ToUpper()
                   }

                    $found = $true
                    $ready = "False"
                    $status = "-"
                    if($adminnote.Contains("W10:OK")){
                        $Ready = "True"
                    }
                    if($adminnote.Contains("W10:NVT")){
                        $Ready = "True"
                    }
                    if($adminnote.Contains("W10:NEW")){
                        $Ready = "True"
                    }

                    if($adminnote -match "W10:(.*?)( |$)" ){
                        $Status =  $matches[1]

                    } 

                     if($dicGLobalGroupsv2.ContainsKey($GLGroup)){
                            $deVolkbankAppName =  $dicGLobalGroupsv2[$GLGroup].Info       
                         } 

                   
                    if([string]::IsNullOrEmpty( $deVolkbankAppName)){
                                      $deVolkbankAppName = $DLGroup.sns_naam
                    }else{
                    
                            $deVolkbankAppName =  $deVolkbankAppName.Split(" ")[0]  
                    }
                 


                      $listResult += New-Object -TypeName PSObject -Property @{
                       SamAccountname = $user.SamAccountName
                       Name = $user.Name
                       Afdeling = $user.businesscategory.Value
                       Ready = $ready 
                       Status = $STatus
                       GAP = $GLGroup 
                       DAP = $DLGroup.ac_grouplist
                       Title = $DLGroup.c_title 
                       SNSNaam =  $deVolkbankAppName}
                      # "---"

               }
          
            }
        }
    }
}



        $Select = @("Afdeling","SamAccountName","Name","Ready","Status", "SNSNaam","Title","GAP"
                    "DAP"
                     )
            		

        $SortOrder = @("Afdeling", "Name", "SNSNaam","Title")

        #SDE_Beheergroep,VM_Folder,	Naam, Team, SDE_Omschrijving, SDE_Domain,OS,SDE_OS,AD_OS,VM_OS, AD_LastLogon, SDE_IsActief, SDE_IsVirtueel, AD_IsActief, VM_IsActief,Collectie, AD_CN 
        $resultsorted = $listResult | Select-Object 	$Select

        $workbook = New-SNSExcelFromTemplate -TemplatePath  $OuputPath_BR11_tpl -ExcelPath  $OuputPath_BR11_xlsx 
        Save-SNSADObjectToExcelRange -InputObject $resultsorted -WorkBook $workbook -Sort $SortOrder -SheetName "Bron" -EntireRow $false -verbose -tempPath $OuputPath_BR11_csv  
        Save-ExcelWorkbook -workbook $workbook -close
        Close-Excel | Out-Null





