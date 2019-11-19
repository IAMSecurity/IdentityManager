
# WPS Init
    $ScriptDirectory = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
    Set-Location $ScriptDirectory
    . ..\..\data\SNSBankNV_WPS.ps1

#End WPS Init


$SNS_KC_Objects = @()
$Managers = @{}

$CSV = Import-Csv  $OuputPath_LocalAdmin -UseCulture
#$SDEServers = Import-Csv $OuputPath_SDEServers -UseCulture 


ForEach($manager in Get-ADUser -filter {St -gt 0 -and Title -gt 0} -Properties  $ADUserProperties){
    if(!$Managers.ContainsKey($manager.st)){$Managers.Add($manager.st,$manager)}
}

$dicSDE = @{}
ForEach($sde in $SDEServers){
    if(!$dicSDE.ContainsKey($sde.Naam.ToUpper())){
        $dicSDE.Add($sde.Naam.ToUpper(),$sde)
    }
}

 if($glocal:GroupCache -eq $null){$global:GroupCache = @{}}
$result = @()

$PRoperties = @{
    Applicatie="" 
    Server=""
    BeheergroepID = ""
    Beheergroep = ""
    BeheergroepManager = ""
    DN=""
    OK="Error"

    Direct_MemberDomain=""                                                                                           
    Direct_MemberType=""
    Direct_MemberName=""                                                                                            
    Direct_Status=""                                                                                                     
    Direct_LocalGroupName=""

    Indirect_MemberDomain=""                                                                                           
    Indirect_MemberType=""
    Indirect_MemberName=""                                                                                              
    Indirect_Status=""                                                                                              
    Indirect_AccountType=""

    Template_Name=""
    Template_Eigenaar=""      
}


 ForEach($computer in $CSV){
    $ComputerName = $computer.server 
    $BeheergroepID= $computer.BeheergroepID
    $computer.OK = "OK"
    if($dicSDE.ContainsKey( $computer.server.ToUpper() )){
        $computer.Beheergroep = $dicSDE[ $computer.server.ToUpper()].SDE_Beheergroep
    }
    try{
        $ADComputer = Get-ADComputer $ComputerName -Properties CanonicalName
        $computer.DN = $ADComputer.CanonicalName
    }catch{
        Write-Warning "Could not find computer $ComputerName"
    }
    If($computer.Direct_Status -eq "Offline"){$computer.OK = "FOUT: De server kan niet benaderd worden (Offline)"}
    If($computer.Direct_Status -eq "Failed"){$computer.OK = "FOUT: De server kan niet benaderd worden (Access Denied)"}
    If($computer.Direct_MemberType -eq "DomainUser"){  $computer.OK = "FOUT: Een account is direct lid van de local admin groep"    }
    If($computer.Direct_MemberName.StartsWith("S-1")){$computer.OK = "FOUT: Een oud account is lid van de local admin groep"} 
   
   
    If($computer.Direct_MemberName -eq "GAD_AdminAll" -or $computer.Direct_MemberName -eq "DAD_AdminAll" -or $computer.Direct_MemberType -ne "DomainGroup" ){
        $result += $computer
        continue
     }
    
    $FullName = "$($computer.Direct_MemberDomain )\$($computer.Direct_MemberName)"
    $members = $null
    if($global:GroupCache.ContainsKey($FullName)){
        $members = $global:GroupCache[$FullName]
    }elsE{
        Try{
            $members = Get-ADGroupMember -Server $computer.Direct_MemberDomain -Identity $computer.Direct_MemberName
        }catch{
            $global:GroupCache.Add($FullName,$members)
            $computer.OK = "FOUT: De groep  $FullName in de local admin groep is niet gevonden in Active Directory ";Continue
        }

        $global:GroupCache.Add($FullName,$members)
    }


    ForEach($member in $members){
        $Newcomputer = New-Object -TypeName PSObject -Property $PRoperties
        
        $Newcomputer.Applicatie= $computer.Applicatie
        $Newcomputer.Server= $computer.Server
        $Newcomputer.BeheergroepID= $computer.BeheergroepID
        $Newcomputer.Beheergroep = $computer.Beheergroep
        $Newcomputer.DN = $computer.DN
        $Newcomputer.BeheergroepManager = $computer.BeheergroepManager
        $Newcomputer.OK="OK"

        $Newcomputer.Direct_MemberDomain= $computer.Direct_MemberDomain                                                                             
        $Newcomputer.Direct_MemberType= $computer.Direct_MemberType
        $Newcomputer.Direct_MemberName= $computer.Direct_MemberName
        $Newcomputer.Direct_Status= $computer.Direct_Status
        $Newcomputer.Direct_LocalGroupName= $computer.Direct_LocalGroupName    
          


        $Newcomputer.Indirect_MemberName = $member.name
        $Newcomputer.Indirect_MemberType = $member.objectClass
        
        if($member -match "(.*),DC=(.*),DC=(.*)"){$Newcomputer.Indirect_MemberDomain = $Matches[2].ToUpper()}

        If($Newcomputer.Indirect_MemberType -eq "user"){                        
            $Newcomputer.Indirect_AccountType = "Onbekend"
            if($member.distinguishedName.Contains( "OU=D-accounts,OU=Admin,OU=Organisatie")) {$Newcomputer.Indirect_AccountType = "D-Account"}        
            if($member.distinguishedName.Contains( "OU=A-accounts,OU=Admin,OU=Organisatie")) {$Newcomputer.Indirect_AccountType = "A-Account"}
            if($member.distinguishedName.Contains( "OU=Service Accounts")) {$Newcomputer.Indirect_AccountType = "Service Account"}
            if($member.distinguishedName.Contains( "OU=Users,OU=Organisatie")) {$Newcomputer.Indirect_AccountType = "Personeels account"}
       
        
            If($Newcomputer.Indirect_AccountType -eq "Service Account"){
                $Newcomputer.Indirect_Status = "OK"
            }else{
                $Newcomputer.Indirect_Status = "OK"
                $Newcomputer.OK = "FOUT: Alleen een TPL groep of Service account mag indirect lid zijn"
            }
         }
       
        If($Newcomputer.Indirect_MemberType -eq "computer"){
            $Newcomputer.Indirect_AccountType = "Computer"
            $Newcomputer.OK = "OK"
        }
        
        If($Newcomputer.Indirect_MemberType -eq "group"){
            $Newcomputer.Indirect_AccountType = "Groep"
            $Newcomputer.OK = "FOUT: Alleen een TPL groep of Service account mag indirect lid zijn"
        }
        
            If($Newcomputer.Indirect_MemberName.StartsWith("TPL_")){$Newcomputer.OK = "OK"}        
            If($Newcomputer.Direct_MemberName -eq "GAD_AdminAll"){$Newcomputer.OK = "OK"}
            If($Newcomputer.Direct_MemberName -eq "DAD_AdminAll"){$Newcomputer.OK = "OK"}
            If($Newcomputer.Direct_MemberName -eq "Domain Admins"){$Newcomputer.OK = "OK"}
            If($Newcomputer.Direct_MemberName -eq "Enterprise Admins"){$Newcomputer.OK = "OK"}     
            If($Newcomputer.Indirect_MemberName -eq "GAD_AdminAll"){$Newcomputer.OK = "OK"}
            If($Newcomputer.Indirect_MemberName -eq "DAD_AdminAll"){$Newcomputer.OK = "OK"}
            If($Newcomputer.Indirect_MemberName -eq "Domain Admins"){$Newcomputer.OK = "OK"}
            If($Newcomputer.Indirect_MemberName -eq "Enterprise Admins"){$Newcomputer.OK = "OK"}

        $result += $Newcomputer

    }

 }
  $sortORder = @("Applicatie,Server","DN","Beheergroep","BeheergroepId","BeheergroepManager", "OK", "Direct_MemberDomain",  "Direct_MemberType", 
                "Direct_MemberName",  "Direct_Status",      "Indirect_MemberDomain",  
                "Indirect_MemberType",    "Indirect_MemberName",     "Indirect_Status","Indirect_AccountType",    "Template_Name",    "Template_Eigenaar")
    $resultSort =  $result | Select-Object  $sortORder|Sort-Object     $sortORder

 

    $workbook = New-SNSExcelFromTemplate -TemplatePath  $OuputPath_BR02_tpl -ExcelPath  $OuputPath_BR02_xlsx 
    Save-SNSADObjectToExcelRange -InputObject $resultSort -WorkBook $workbook -Sort $SortOrder -SheetName "Servers" -EntireRow $false -verbose 
    Save-ExcelWorkbook -workbook $workbook -close
    Close-Excel | Out-Null

 
 $result | Select-Object Applicatie,Server,DN,Beheergroep,BeheergroepId,BeheergroepManager, OK, Direct_MemberDomain,  Direct_MemberType, 
                Direct_MemberName,  Direct_Status,      Indirect_MemberDomain,  
                Indirect_MemberType,    Indirect_MemberName,     Indirect_Status,Indirect_AccountType,    Template_Name,    Template_Eigenaar | Export-Csv  $OuputPath_BR02_csv  -UseCulture -NoTypeInformation


