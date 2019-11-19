<#
    File   : Process_ADComputer.ps1
    Author : Rob Looman
    Date   : 28-01-2016
    Company: SNS Bank N.V.
    Version: 1.0

    Description:
        Haalt AD Computer objecten op
    Change Log:
        RL-15012016:  
#>

# Config 

     


# Init 
    $list = @()
    $Managers = @{}
    $dicObjects = @{}
    Import-Module SNSExcel -force
    . ..\..\data\SNSBankNV_WPS.ps1

#Process

    $dicDept2Mgr = @{}


    $AD = Import-Csv $OuputPath_ADComputerAll -UseCulture 
    #$VMWare = Import-Csv $OuputPath_VMWare -UseCulture 
    $SCCM = Import-Csv $OuputPath_SCCMDevices -UseCulture 
    $WSUS = Import-Csv $OuputPath_WSUS -UseCulture 
    $Server = Import-Csv $OuputPath_ServerAssignment -Delimiter ","
    $Sophos = Import-CSV $OuputPath_Sophos -UseCulture
    $SNOW = Import-Csv $OuputPath_SNOWServers -UseCulture 
    $SNOWPC = Import-Csv $OuputPath_SNOWPC -UseCulture 

    $list = $AD+ $SCCM +$WSUS+$Server +   $Sophos  +  $SNOW + $SNOWPC 
    #Custom result object
        $properties = @{}
        $properties.Add("OS","")
        $properties.Add("Manager","")
        $properties.Add("DMZ_ISActief","False")
        ForEach($item in $AD |Get-Member -MemberType NoteProperty){
           if( -not $properties.ContainsKey( $item.Name)){
             $properties.Add($item.Name,"")
           }
        }
        <#
        ForEach($item in $VMWare |Get-Member -MemberType NoteProperty){
           if( -not $properties.ContainsKey( $item.Name)){
             $properties.Add($item.Name,"")
           }
        }
        #>
        ForEach($item in $SCCM |Get-Member -MemberType NoteProperty){
           if( -not $properties.ContainsKey( $item.Name)){
             $properties.Add($item.Name,"")
           }
        }
        ForEach($item in $WSUS |Get-Member -MemberType NoteProperty){
           if( -not $properties.ContainsKey( $item.Name)){
             $properties.Add($item.Name,"")
           }
        }
        ForEach($item in $Sophos |Get-Member -MemberType NoteProperty){
           if( -not $properties.ContainsKey( $item.Name)){
             $properties.Add($item.Name,"")
           }
        }
        ForEach($item in $Server |Get-Member -MemberType NoteProperty){
           if( -not $properties.ContainsKey( $item.Name)){
             $properties.Add($item.Name,"")
           }
        }
        
        ForEach($item in $SNOW |Get-Member -MemberType NoteProperty){
           if( -not $properties.ContainsKey( $item.Name)){
             $properties.Add($item.Name,"")
           }
        }
        
        ForEach($item in $SNOWPC |Get-Member -MemberType NoteProperty){
           if( -not $properties.ContainsKey( $item.Name)){
             $properties.Add($item.Name,"")
           }
        }
        $managers = Get-ADUser -filter {St -gt 0 -and Title -gt 0} -Properties st,title,department,businessCategory ,mail
        $Managers | Add-Member -MemberType ScriptProperty -Name Afdeling -Value {$this.businessCategory} -Force 
     
      $dicManagers =   $managers| Group-Object Afdeling -AsHashTable -AsString
   



    Foreach($item in $list){
    
        $NAAM = $null
        if(-not [string]::IsNullOrEmpty($item.Name)){
            $naam = $item.Name.ToUpper()
        } 
        if(-not [string]::IsNullOrEmpty($item.Naam)){
            $naam = $item.Naam.ToUpper()
        }

        $item | Add-Member -NotePropertyName Name -NotePropertyValue $NAAM -force

        if(!$dicObjects.ContainsKey( $naam )){
            $temp = New-Object -TypeName PSObject -Property $properties 
            $dicObjects.Add($naam,$temp)
         }else{
            $temp =   $dicObjects[ $naam ] 
         }

         #Determine if is DMZ system 
      
        ForEach($member in $item | Get-Member -MemberType NoteProperty ){       
               $membername = $member.Name
               $temp.$membername =  $item.$membername
               #Fill Manager 
               if($membername -eq "SNOW_EigenaarAfdeling"){
                    if( $dicManagers.ContainsKey($item.$membername) ){
                        $value = $dicManagers[$item.$membername].Mail
                        if($value -ne $null){
                            $temp.Manager = [string]::Join(",", $value )
                         }
                     }else{
                        #Write-Warning "Manager not found for costcenter $($item.$membername)" 
                     }
                 
               }
              #Combine OS 
              if($membername.Contains("_OS")){
                if(![string]::IsNullOrEmpty($item.$membername )){
                    $temp.OS =  $item.$membername

                }
              }
              
        }
        
        if(-not[string]::IsNullOrEmpty($temp.SNOW_IPAddress)){
            if( $temp.SNOW_IPAddress.Contains("10.208.") -or $temp.SNOW_IPAddress.Contains("10.192.")){
                $temp.DMZ_ISActief = "True"
            }
        }


        $dicObjects[ $naam ] = $temp


    }

    $list = @()
    $dicObjects.Keys | %{
        $list += $dicObjects[$_]

    } 
    
    
    
    # Totaal lijst             
        $Select = @("SNOW_EigenaarAfdeling","SNOW_Beheergroep","Manager","Name",
                    "SNOW_Omschrijving", "OS", 
                    "SNOW_IsActief", "AD_IsActief", "DMZ_IsActief"
                    "AD_LastLogon", "WSUS_LastSyncTime","SCCM_Hardwarescan",
                    "Sophos_VersionSAV",  "WSUS_NeededCount",   "WSUS_FailedCount", 
                    "SCCM_Collectie",  "WSUS_LastReportedStatusTime", "WSUS_Server",
                    "SCCM_Kostenplaats","SCCM_REsourceID",	
                    "AD_OS","AD_CN",
                    "SNOW_Type" ,"SNOW_Location", "SNOW_Toelichting",                   
                    "Sophos_Domain" ,"Sophos_Managed","Sophos_InstalledAU","Sophos_InstalledSAV", "Sophos_InstalledOnAccess","Sophos_InstalledWeb",
                    "Sophos_VersionSoftware","Sophos_VersionEnging",
                    "Sophos_VersionVirusData","Sophos_VersionAgent","Sophos_TimeChangedOnEP","Sophos_TimeLastUpdate"
                     )
            		

        $SortOrder = @("SNOW_EigenaarGroep", "AD_CN", "SNOW_IsActief")

        #SDE_Beheergroep,VM_Folder,	Naam, Team, SDE_Omschrijving, SDE_Domain,OS,SDE_OS,AD_OS,VM_OS, AD_LastLogon, SDE_IsActief, SDE_IsVirtueel, AD_IsActief, VM_IsActief,Collectie, AD_CN 
        $resultsorted = $list | Select-Object 	$Select

        $workbook = New-SNSExcelFromTemplate -TemplatePath  $OuputPath_BR01Server_tpl -ExcelPath  $OuputPath_BR01All_xlsx 
        Save-SNSADObjectToExcelRange -InputObject $resultsorted -WorkBook $workbook -Sort $SortOrder -SheetName "Servers" -EntireRow $false -verbose -tempPath $OuputPath_BR01All_csv  
        Save-ExcelWorkbook -workbook $workbook -close
     Close-Excel | Out-Null
    # Werkplek lijst    
    
    


 		
        #SDE_Beheergroep,VM_Folder,	Naam, Team, SDE_Omschrijving, SDE_Domain,OS,SDE_OS,AD_OS,VM_OS, AD_LastLogon, SDE_IsActief, SDE_IsVirtueel, AD_IsActief, VM_IsActief,Collectie, AD_CN 
        $resultsorted =  $list | Where-Object {$_.Name.StartsWith("ONV")}  | Select-Object 	$Select


        $workbook = New-SNSExcelFromTemplate -TemplatePath  $OuputPath_BR01Server_tpl -ExcelPath  $OuputPath_BR01Werkplek_xlsx 
        Save-SNSADObjectToExcelRange -InputObject $resultsorted -WorkBook $workbook -Sort $SortOrder -SheetName "Servers" -EntireRow $false -verbose -tempPath $OuputPath_BR01Werkplek_csv  
         $list | Where-Object {$_.Naam.StartsWith("ONV")} | Export-csv  $OuputPath_BR01Werkplek_csv  -UseCulture -NoTypeInformation
        Save-ExcelWorkbook -workbook $workbook -close
       Close-Excel | Out-Null  
       # Server lijst             


        #SDE_Beheergroep,VM_Folder,	Naam, Team, SDE_Omschrijving, SDE_Domain,OS,SDE_OS,AD_OS,VM_OS, AD_LastLogon, SDE_IsActief, SDE_IsVirtueel, AD_IsActief, VM_IsActief,Collectie, AD_CN 
        $resultsorted =  $list | Where-Object {$_.Name.StartsWith("S")} | Select-Object 	$Select


        $workbook = New-SNSExcelFromTemplate -TemplatePath  $OuputPath_BR01Server_tpl -ExcelPath  $OuputPath_BR01Server_xlsx 
        Save-SNSADObjectToExcelRange -InputObject $resultsorted -WorkBook $workbook -Sort $SortOrder -SheetName "Servers" -EntireRow $false -verbose -tempPath $OuputPath_BR01Server_csv  
        Save-ExcelWorkbook -workbook $workbook -close
    



    #####
    # Eigenaar	Beheergroep	Manager	Naam	Omschrijving	SNOW_IsActief	AD_IsActief	DMZ_IsActief	Sophos Version	WSUS Missing	SCCM MW


    $Select = @("SNOW_EigenaarAfdeling","SNOW_Beheergroep","Manager","Name",
                    "SNOW_Omschrijving",  
                    "SNOW_IsActief", "AD_IsActief", "DMZ_IsActief",                    
                    "Sophos_VersionSAV",  "WSUS_NeededCount",  "SCCM_Collectie",
                    "APP_Adobe Flash Player ActiveX Latest",
                    "APP_Adobe Flash Player Plugin Latest",
                    "APP_Adobe Reader XI Latest",
                    "APP_Adobe Shockwave Player Latest",
                    "APP_Java 8 Latest","AD_OS"
                     )
            		

        $SortOrder = @("SNOW_EigenaarAfdeling", "SNOW_EigenaarGroep", "Naam")
  
        #SDE_Beheergroep,VM_Folder,	Naam, Team, SDE_Omschrijving, SDE_Domain,OS,SDE_OS,AD_OS,VM_OS, AD_LastLogon, SDE_IsActief, SDE_IsVirtueel, AD_IsActief, VM_IsActief,Collectie, AD_CN 
         $resultsorted =  $list | Where-Object {$_.Name.StartsWith("S") -and ($_.SNOW_ISActief -eq $true -or $_.AD_ISActief -eq $true  ) } | Select-Object 	$Select

        $workbook = New-SNSExcelFromTemplate -TemplatePath  $OuputPath_BR01ServerAfd_tpl -ExcelPath  $OuputPath_BR01ServerAfd_xlsx 
        Save-SNSADObjectToExcelRange -InputObject $resultsorted -WorkBook $workbook -Sort $SortOrder -SheetName "Servers" -EntireRow $false -verbose -tempPath $OuputPath_BR01Server_csv  
        Save-ExcelWorkbook -workbook $workbook -close
  
         $resultsorted =  $list | Where-Object {$_.Name.StartsWith("O") -and ($_.SNOW_ISActief -eq $true -or $_.AD_ISActief -eq $true  )} | Select-Object 	$Select
        $workbook = New-SNSExcelFromTemplate -TemplatePath  $OuputPath_BR01ServerAfd_tpl -ExcelPath  $OuputPath_BR01LaptopsAfd_xlsx 
        Save-SNSADObjectToExcelRange -InputObject $resultsorted -WorkBook $workbook -Sort $SortOrder -SheetName "Servers" -EntireRow $false -verbose -tempPath $OuputPath_BR01Server_csv  
        Save-ExcelWorkbook -workbook $workbook -close

        
         $resultsorted =  $list | Where-Object {$_.Name.StartsWith("V") -and ($_.SNOW_ISActief -eq $true -or $_.AD_ISActief -eq $true  )} | Select-Object 	$Select
        $workbook = New-SNSExcelFromTemplate -TemplatePath  $OuputPath_BR01ServerAfd_tpl -ExcelPath  $OuputPath_BR01VDIAfd_xlsx 
        Save-SNSADObjectToExcelRange -InputObject $resultsorted -WorkBook $workbook -Sort $SortOrder -SheetName "Servers" -EntireRow $false -verbose -tempPath $OuputPath_BR01Server_csv  
        Save-ExcelWorkbook -workbook $workbook -close

     Close-Excel | Out-Null
