# Configuration
. ..\..\Data\SNSBankNV_WPS.ps1

$LogFile = "D:\ScriptRepository\Log\BR_05_SCCMPAtchRondes.log"
"" | Out-file  $LogFile 
$currentLocation = Get-Location
    #Invoke-WebREquest  -Uri $url -OutFile $temp -Credential $cred 
    $csv = Import-CSV $OuputPath_MR01_csv -UseCulture 
    $list = @{"ALG - Werkplek - Update Ronde 1 ( Bal Update )"="GAP_SCCM_WerkplekUpdateRonde1" }

    $listFriendly = @{"ALG - Werkplek - Update Ronde 1 ( Bal Update )"="Linking Pin testgroep (2 weken test)" }


    cd P01:
    $CSVREsult = @()
    Foreach($key in $list.Keys){
        
       
        $SCCMDevices = Get-CMDevice -CollectionName $key 
        $GroupMembers = Get-ADGroupMember -Identity     $list[$key]

        $dicUsers = $csv |Group-Object -AsHashTable -Property RES_GebruikersCode
        $dicSCCMDevices = $SCCMDevices |Group-Object -AsHashTable -Property Name
        if($dicSCCMDevices -eq $null){$dicSCCMDevices =@{}}

        ForEach($member in $GroupMembers){
             
             $Userkey = $member.SamAccountName
                $Userkey 
            if($dicUsers.Contains($Userkey)){
                ForEach($ComputerName in $dicUsers[$Userkey].Name){
                    $CSVREsult += New-Object -TypeName PSObject -Property @{
                        SAMACcountName = $member.SamAccountName
                        Name =$member.name
                        Ronde = $key
                        Beschrijving = $listFriendly[$key]
                        Laptop = $ComputerName
                     }
                    if($dicSCCMDevices.ContainsKey($ComputerName)){
                        "Already containing $computername in collecntion $key $($member.SamAccountName)" | Out-file  $LogFile  -Append
                        $dicSCCMDevices[$ComputerName] = $null
                    }else{
                        "Adding $ComputerName to collection $key  $($member.SamAccountName)" | Out-file  $LogFile  -Append
                        $device = Get-CMDevice -Name $ComputerName 
                        if($device.ResourceID -ne $null){
                            
                            Add-CMDeviceCollectionDirectMembershipRule -CollectionName $key -ResourceId $device.ResourceID -ErrorAction SilentlyContinue
                        }else{
                            "FAILED: Adding $ComputerName to collection $key  $($member.SamAccountName)" | Out-file  $LogFile  -Append
                        }
                    }
                }
           
            }Else{
                 "User $($member.SamAccountName) not found  "| Out-file  $LogFile  -Append
            }
        
        }


        ForEach($keyDevice in $dicSCCMDevices.Keys){
            if($dicSCCMDevices[$keyDevice] -ne $null){

                 "Computername not found: $keyDevice" | Out-file  $LogFile  -Append
                Remove-CMDeviceCollectionDirectMembershipRule -CollectionName $key -ResourceName $keyDevice -Force
            }
        }
    }

    
      
    
    $sortORder = @("SAMACcountName","Name","Beschrijving","Ronde","Laptop")
    $resultSort =  $CSVREsult | Select-Object  $sortORder|Sort-Object     Ronde, Name

 

    $workbook = New-SNSExcelFromTemplate -TemplatePath  $OuputPath_BR05_tpl -ExcelPath  $OuputPath_BR05_xlsx 
    Save-SNSADObjectToExcelRange -InputObject $resultSort -WorkBook $workbook -Sort Ronde,Name -SheetName "Servers" -EntireRow $false -verbose 
    Save-ExcelWorkbook -workbook $workbook -close
    Close-Excel | Out-Null
    




     $CSVREsult | Sort-Object Ronde,Name | Export-Csv -path $OuputPath_BR05_csv -UseCulture -NoTypeInformation
     $currentLocation | Set-Location