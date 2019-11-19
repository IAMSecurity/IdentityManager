
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

     
# WPS Init
    $ScriptDirectory = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
    Set-Location $ScriptDirectory
    . ..\..\data\SNSBankNV_WPS.ps1

#End WPS Init



# Init 
    $list = @()
    $result = @()

#Process
    $list = Import-Csv  $OuputPath_BR01Werkplek_csv -UseCulture 
    
    #determine BAL
    $propertieBAL = $list | Get-Member |? {$_.Name -like "APP_BAL 20*"}
    

    Foreach($item in $list){

        $SNOW_LastSeen = 0
        $SNOW_Status   = "-"
        $SNOW_LastSeenDate = "-"
        $AD_Status    = "-"
        $AD_LastSeen  = 0
        $AD_LastSeenDate = "-"
        $Sophos_LastSeen = -1
        $Sophos_LastSeenDate = "-"
        $WSUS_LastSeen = -1
        $WSUS_LastSeenDate = "-"

        $DaysSNOW
        $Checked = "-"
        $WSUS_NeededCount = ""
        $Sophos_VersionSAV = ""
        $Sophos_TimeChangedOnEP =""          

              
        $WSUS_NeededCount = $item.WSUS_NeededCount
        $Sophos_VersionSAV = $item.Sophos_VersionSAV

        $BAL = $item.$($propertieBAL.Name)
         
        
         if($item.SNOW_Toelichting.Contains("WPS CFG")){
            $Checked = "Yes"
         }
         
        #checking date
            #AD
            if(-not [string]::IsNullOrEmpty($item.AD_LastLogon)){
                $timeinfo = $item.AD_LastLogon
                $template = 'd-M-yyyy HH:mm:ss'
                $AD_LastSeenDate = [DateTime]::ParseExact($timeinfo, $template, $null) 
                $AD_LastSeen  = ($currentDateTime - $AD_LastSeenDate).Days
                  
            }
            #SDE
            if(-not [string]::IsNullOrEmpty($item.SNOW_WhenModified)){
                $timeinfo = $item.SNOW_WhenModified
                $template = 'd-M-yyyy HH:mm:ss'
                $SNOW_LastSeenDate = [DateTime]::ParseExact($timeinfo, $template, $null) 
                $SNOW_LastSeen  = ($currentDateTime - $SNOW_LastSeenDate).Days
                  
            }
            #Sophos
            if(-not [string]::IsNullOrEmpty($item.Sophos_TimeChangedOnEP)){
                $timeinfo = $item.Sophos_TimeChangedOnEP
                $template = 'd-M-yyyy HH:mm:ss'
                $Sophos_LastSeenDate = [DateTime]::ParseExact($timeinfo, $template, $null) 
                $Sophos_LastSeen  = ($currentDateTime - $Sophos_LastSeenDate).Days
                  
            }
            #WSUS
            if(-not [string]::IsNullOrEmpty($item.WSUS_LastReportedStatusTime)){
                $timeinfo = $item.WSUS_LastReportedStatusTime
                $template = 'd-M-yyyy HH:mm:ss'
                $WSUS_LastSeenDate = [DateTime]::ParseExact($timeinfo, $template, $null) 
                $WSUS_LastSeen  = ($currentDateTime - $WSUS_LastSeenDate).Days
                  
            }


         if($item.SNOW_Status -notin  $SNOW_OKStatus){
             $SNOW_Status = "NOT OK"
         }
         
         if( $item.SNOW_Status -in  $SNOW_OKStatus){
             $SNOW_Status = "OK"
         }
         if([string]::IsNullOrEmpty($item.SNOW_Status)){
             $SNOW_Status = "-"            
         }

         

         if($AD_LastSeen -gt 90 ){
             $AD_Status = "NOT OK"
         }
         
         if($AD_LastSeen -lt 90 ){
             $AD_Status = "OK"
         }
         if([string]::IsNullOrEmpty($item.AD_LastLogon)){
             $AD_Status = "-"            
         }
         $Actie = "Geen Actie"
         if($SNOW_Status -eq "-" -and $AD_Status -eq "OK" -and $Checked -ne "Yes"){
            $Actie = "Actie uitzoeken ServiceNow"
         }
         if($SNOW_Status -eq "-" -and $AD_Status -eq "NOT OK" -and $Checked -ne "Yes"){
            $Actie = "Actie opruimen"
         }
         if($SNOW_Status -eq "OK" -and $AD_Status -eq "NOT OK" -and $Checked -ne "Yes"){
            $Actie = "Actie uitzoeken ServiceNow"
         }
         if($SNOW_Status -eq "OK" -and $AD_Status -eq "NOT OK" -and $Checked -eq "Yes"){
            $Actie = "Geen Actie 1"
         }
         if($SNOW_Status -eq "NOT OK" -and $AD_Status -eq "OK" -and $Checked -ne "Yes"){
            $Actie = "Actie uitzoeken ServiceNow"
         }
         if($SNOW_Status -eq "NOT OK" -and $AD_Status -eq "OK" -and $Checked -eq "Yes"){
            $Actie = "Geen Actie 2"
         }
         if($SNOW_Status -eq "NOT OK" -and $AD_Status -eq "NOT OK" -and $Checked -ne "Yes"){
            $Actie = "Actie opruimen"
         }

         $properties = @{
            Name = $item.Naam
            SNOW_Status = $SNOW_Status 
            AD_Status = $AD_Status  
            Checked = $Checked  
            Actie = $Actie
            Verschil_SNOW_AD = $SNOW_LastSeen  - $AD_LastSeen
            SNOW_LastSeen = $SNOW_LastSeen 
            SNOW_LastSeenDate = $SNOW_LastSeenDate 
            SNOW_Status1=$item.SNOW_Status
            AD_LastSeen = $AD_LastSeen 
            AD_LastSeenDate =  $AD_LastSeenDate
            SNOW_Beheergroep = $item.SNOW_Beheergroep
            SNOW_Toelichting = $item.SNOW_Toelichting
            SNOW_Type = $item.SNOW_Type
            SNOW_Location = $item.SNOW_Location
            Sophos_LastLogonUser = $item.Sophos_LastLogonUser
            Sophos_LastSeen =$Sophos_LastSeen
            Sophos_LastSeenDate= $Sophos_LastSeenDate 
            WSUS_LastSeen = $WSUS_LastSeen 
            WSUS_LastSeenDate = $WSUS_LastSeenDate 
            WSUS_NeededCount = $WSUS_NeededCount
            Sophos_VersionSAV = $Sophos_VersionSAV 
            Sophos_TimeChangedOnEP = $Sophos_TimeChangedOnEP      
            BALVersion = $BAL

        }
        $result += New-Object -TypeName PSObject -Property $properties
    }

    $result |Select-Object Name,SNOW_Status, AD_Status , Checked ,Actie,  SNOW_Status1, Verschil_SNOW_AD ,SNOW_LastSeen , SNOW_LastSeenDate , AD_LastSeen,AD_LastSeenDate,
            SNOW_Beheergroep ,SNOW_Type,SNOW_Location, Sophos_LastLogonUser,
            Sophos_LastSeen ,   Sophos_LastSeenDate,Sophos_VersionSAV, Sophos_TimeChangedOnEP,
            WSUS_LastSeen ,           WSUS_LastSeenDate ,WSUS_NeededCount, BALVersion,SNOW_Toelichting  |  
        Export-Csv -UseCulture -NoTypeInformation -Path $OuputPath_BR03_csv 

      