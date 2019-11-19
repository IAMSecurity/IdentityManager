    . ..\..\data\SNSBankNV_WPS.ps1

Import-Module MsOnline, ActiveDirectory

$AzureADUser = "SA_AzureAD_Scripts@devolksbank.nl"
$pwdFileName = "SecureString_SA-AzureAD-Scripts_SA-ADScript"

#Inloggegevens
#Versleuteld wachtwoord is middels \\verz.local\groepsdata$\Werkplek Services\Script_Repository\Scripts_PowerShell\Office365\Include\ExportPasswordwithAES.ps1 gegenereerd
#Let op, het script moet draaien onder het zelfde windows account als waarmee de versleutelde wachtwoord file is gemaakt!
$AESKey = Get-Content "\\verz.local\groepsdata$\Werkplek Services\Script_Repository\Scripts_PowerShell\Office365\Include\AESKey.txt"
$pwdTxt = Get-Content "\\verz.local\groepsdata$\Werkplek Services\Script_Repository\Scripts_PowerShell\Office365\Include\$pwdFileName.txt"
$securePwd = $pwdTxt | ConvertTo-SecureString -Key $AESKey
$Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $AzureADUser, $securePwd

#Eerst een sessie opzetten naar de SysProxy 
$SetProxy=Invoke-WebRequest -Proxy http://sysproxy.verz.local:80 -ProxyUseDefaultCredentials -uri https://login.microsoftonline.com -UseBasicParsing

#Verbind naar Office 365
Connect-MsolService -Credential $Credential 

    function Get-UserGroupMembershipRecursive {
[CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [String[]]$UserName
    )
    begin {
        # introduce two lookup hashtables. First will contain cached AD groups,
        # second will contain user groups. We will reuse it for each user.
        # format: Key = group distinguished name, Value = ADGroup object
        if($Global:ADGroupCache -eq $null){
            $Global:ADGroupCache= @{}
        } 
        $UserGroups = @{}
        # define recursive function to recursively process groups.
        function __findPath ([string]$currentGroup) {
            Write-Verbose "Processing group: $currentGroup"
            # we must do processing only if the group is not already processed.
            # otherwise we will get an infinity loop
            if (!$UserGroups.ContainsKey($currentGroup)) {
                # retrieve group object, either, from cache (if is already cached)
                # or from Active Directory
                $groupObject = if ($Global:ADGroupCache.ContainsKey($currentGroup)) {
                    Write-Verbose "Found group in cache: $currentGroup"
                    $Global:ADGroupCache[$currentGroup]
                } else {
                    Write-Verbose "Group: $currentGroup is not presented in cache. Retrieve and cache."
                    $g = Get-ADGroup -Identity $currentGroup -Property "MemberOf" -ErrorAction SilentlyContinue -server verz.local:3268
                    # immediately add group to local cache:
                    $Global:ADGroupCache.Add($g.DistinguishedName, $g)
                    $g
                }
                # add current group to user groups
                $UserGroups.Add($currentGroup, $groupObject)
                Write-Verbose "Member of: $currentGroup"
                foreach ($p in $groupObject.MemberOf) {
                    __findPath $p
                }
            } else {Write-Verbose "Closed walk or duplicate on '$currentGroup'. Skipping."}
        }
    }
    process {
        foreach ($user in $UserName) {
            Write-Verbose "========== $user =========="
            # clear group membership prior to each user processing
            $UserObject = Get-ADUser -Identity $user -Property "MemberOf"
            $UserObject.MemberOf | ForEach-Object {__findPath $_}
            New-Object psobject -Property @{
                UserName = $UserObject.Name;
                MemberOf = $UserGroups.Values | % {$_}; # groups are added in no particular order
            }
            $UserGroups.Clear()
        }
    }
}
			


# Setting variables 
    $SamAccountNameGroup = "DAP_OFFICE365_ProPlus"
    $datum = [system.datetime]::Now.ToString("yyyyMMdd")
    $datum2 = [system.datetime]::Now.ToString("yyyy-MM")

# Versie van TPM
    $TPMOK = @("3.17","3.19","4.43","6.43","7.62","7.63")
    $TPMNotOK = @("4.40","6.40","6.41","7.40","7.41")

#Lijst inladen met te testen gebruikers
    $gebruikers = get-aduser -Filter * -SearchBase "OU=Users,OU=Organisatie,DC=VERZ,DC=Local" -SearchScope OneLevel  -Properties userPrincipalName, EmailAddress , businesscategory, telephoneNumber

    $MaandRap = Import-csv $OuputPath_MR01_csv -UseCulture | Select Name, AD_OS, RES_GebruikersCode, RES_Time,SNOW_Type
    $RES     =  Import-CSV $OuputPath_RESUserToComputer -UseCulture
    $Batch   =  Import-CSV $OuputPath_Windows10Batch  -UseCulture | Group-Object SamAccountName -AsHashTable
    $SCCM   =  Import-CSV $OuputPath_SCCMDevices  -UseCulture | Group-Object Naam -AsHashTable
    $dicHello   = Get-ADGroupMember -Identity "Windows Hello For Business Users" | Group-Object -Property SamAccountName -AsHashTable
   $dicResUser =  $MaandRap | Where-Object {-not [string]::IsNullOrEmpty($_.RES_GebruikersCode)}  | Group-Object RES_GebruikersCode -AsHashTable
 
# Exception lijst met gebruikers die niet gecontroleerd hoeven te worden (schoonmaak etc)
    $Exception = @()
    $exception = Get-Content $OuputPath_Windows10Exception

# Lees de applicaties in RES in een array
    $RESapplicaties = import-csv $OutputPathRESObjects  -UseCulture |Where-Object{$_.RESTypeName -eq "application"}
    $RESapplicaties | Add-Member -MemberType AliasProperty -Name ApplicationName -Value c_title
    $RESapplicaties | Add-Member -MemberType AliasProperty -Name Groups -Value ac_grouplist
    $RESapplicaties | Add-Member -MemberType AliasProperty -Name AdminNote -Value c_administrativenote

    $dicGroupToRES = @{}
    ForEach($Shortcut in $RESapplicaties){
        if([string]::IsNullOrEmpty($Shortcut.Groups)){
        continue 

        }
        $arShortcutSplit = $Shortcut.Groups.Split("#")

        ForEach($item in $arShortcutSplit ){
            $item = $item.ToUpper()
            $item = $item.Replace(" ","")
            $item = $item.Replace("VERZ\","")
            if($dicGroupToRES.ContainsKey($item)){
                $tmp = $dicGroupToRES[$item]
            }else {
                $tmp = @()
            }

            $tmp += $Shortcut
        
            $dicGroupToRES[$item] = $tmp
        }
    

    }
# Init
    $body=""
    $array = @()
    $Prop = @{
	    Gebruiker = ""
	    Informatie = ""
	    Applicaties = ""
    }

#----------------------------------------------------------------------------------------
    $URI = "https://id.snsreaal.nl/sites/011473/Beheer%20Rapportages/BR_10_Windows10Readiness.xlsx"
    $tempXLSX = "C:\Windows\TEmp\temp.xlsx"
    $webclient = New-Object System.Net.WebClient

    $webclient.UseDefaultCredentials = $true
       
  
    $webclient.DownloadFile( $URI , $tempXLSX)
    
    $SqlConnection = New-Object System.Data.OleDb.OleDbConnection

    # Option 1
    # $SqlConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=""$Path""; Extended Properties=""Excel 12.0; HDR=NO; IMEX=1; ReadOnly=True"" "

    # Option 2
    $SqlConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=""$tempXLSX""; Extended Properties=""Excel 12.0; HDR=YES; IMEX=1; ReadOnly=True"" "

    $SqlCmd = New-Object System.Data.OleDb.OleDbCommand
    $SqlCmd.CommandText = "select * From [Bron$]" 
    $SqlCmd.Connection = $SqlConnection

    $SqlAdapter = New-Object System.Data.OleDb.OleDbDataAdapter
    $SqlAdapter.SelectCommand = $SqlCmd

    $DataSet = New-Object System.Data.DataSet
    $nRecs = $SqlAdapter.Fill($DataSet)
    $nRecs | Out-Null

    # Populate Hash Table
    $objTable = $DataSet.Tables[0]
    $dicOpmerking = @{}
    ForEach($row in $objTable){
        $dicOpmerking.Add($row.SamAccountName.ToString(),    $row.Opmerking)
    }



$object = New-Object -TypeName PSObject -Property $Prop
$resultList = @()
$i = 0
$itotal = $gebruikers.Count
$W10Ready = 0
$W10Laptop = 0
ForEach($ADgebruiker in $gebruikers){
    $i++
    Write-Progress -CurrentOperation "$($CurrentUser.SamAccountName) ($i of $itotal)" -Activity "Processing user" -PercentComplete (($i / $itotal)*100)
    $afdeling = ""
    if($ADgebruiker.businesscategory -ne $null){
        $afdeling = $ADgebruiker.businesscategory.Value
    }
    $CurrentUser = New-object -TypeName PSObject -Property @{
        Name = $ADgebruiker.Name 
        Afdeling = $afdeling
        Hello = $dicHello.ContainsKey($ADgebruiker.SamAccountName)
        SamAccountName = $ADgebruiker.SamAccountName
        UPN = $ADgebruiker.UserPrincipalName
        Tel = $ADgebruiker.telephoneNumber
      
        Remark = ""
        Mail = $ADgebruiker.EmailAddress
        App_OK = 0
        App_NotOk = 0 
        Status = "Unknown"
        W10_Laptop = $false
        Error = ""
        LaptopCount = $dicResUser[$ADgebruiker.SamAccountName].Count
        LaptopOS = ""
        LaptopName = ""
        LaptopType = ""
        BatchDatum = ""
        MFA= ""
        TPM=""
    }

    if( $dicOpmerking.ContainsKey( $CurrentUser.SamAccountName.ToString())){
           $CurrentUser.Remark =  $dicOpmerking[$CurrentUser.SamAccountName.ToString()]
    }
    if($CurrentUser.SamAccountName -notmatch "^[\d\.]+$"){  
        continue
    }

    if($Batch.ContainsKey($CurrentUser.SamAccountName)){
        $CurrentUser.BatchDatum = $Batch[$CurrentUser.SamAccountName].Datum
    }
   
    if ($ADgebruiker.UserPrincipalName -notlike "*.nl"){        
        $CurrentUser.Error = "UPN Error" 
    }
    if($exception -contains $CurrentUser.SamAccountName){
        $CurrentUser.Error = "Exception" 
    }
    if($CurrentUser.LaptopCount -eq 0){
        #$CurrentUser.Error = "No Laptop Found" 
    }
        
    # Get Laptop OS        
	if($CurrentUser.LaptopCount -gt 0){
            $CurrentUser.LaptopOS = ($dicResUser[$CurrentUser.SamAccountName] | Sort-Object RES_Time -Descending)[0].AD_OS
            $CurrentUser.LaptopName = ($dicResUser[$CurrentUser.SamAccountName] | Sort-Object RES_Time -Descending)[0].Name
             $CurrentUser.LaptopType =($dicResUser[$CurrentUser.SamAccountName] | Sort-Object RES_Time -Descending)[0].SNOW_Type  
            If ($SCCM.Containskey($CurrentUser.LaptopName)){
                $CurrentUser.TPM = $SCCM[$CurrentUser.LaptopName].TPM_Version
            }
            if( $CurrentUser.TPM -in $TPMNotOK){$CurrentUser.TPM = "Not OK"}
            if( $CurrentUser.TPM -in $TPMOK){$CurrentUser.TPM = "OK"}
    } 

        
	if($CurrentUser.LaptopOS -eq "Windows 10 Enterprise"){
        $CurrentUser.W10_Laptop = $true
	}

    if((-not $CurrentUser.W10_Laptop -or (1 -eq 1)) ){
        $GebruikerLidvanGroepen = (Get-UserGroupMembershipRecursive $CurrentUser.SamAccountName).Memberof 	
   
		ForEach($groep in $GebruikerLidvanGroepen){
            
            if($dicGroupToRES.ContainsKey($groep.Name)){
                ForEach($AppGroep in $dicGroupToRES[$groep.Name]){#$appgroep.title;$AppGroep.AdminNote;}}}
                   
                    if ($appgroep.Groups -match "$($groep.name)$"){
                     
					    if ($AppGroep.AdminNote.ToUpper() -match "W10:OK" -or $AppGroep.AdminNote.ToUpper() -match "W10:FTOK"){
						    $CurrentUser.App_OK++
					    }
					    else{
						    if ($AppGroep.AdminNote.ToUpper() -notmatch "W10:NVT"){							
							    $CurrentUser.App_NotOk++
                                #$appgroep
						    }
					    }
				    }	
                }

            }		
				
			
		}	
    }
		
    if($CurrentUser.App_NotOk -gt 0 ){
        $CurrentUser.Error = "App not ok " 
    }


    if( $CurrentUser.Error  -eq ""){
        $W10Ready++
        $CurrentUser.Status = "Ready"
    }else{
        $CurrentUser.Status = "Not Ready"
    }

    if( $CurrentUser.BatchDatum -ne ""){
        $CurrentUser.Status  = "In Progress"
    }
    
    if( $CurrentUser.MFA -eq ""){
        #$CurrentMFAState =  (Get-MsolUser -UserPrincipalName $CurrentUser.UPN |select -ExpandProperty StrongAuthenticationRequirements).state 
        #MFA is ingesteld, niets doen
        if($CurrentMFAState -eq "Enforced"){
            $CurrentUser.MFA = "Ingesteld"
        #MFA is nog niet ingesteld, instellen
        }elseif($CurrentMFAState -eq "Enabled"){
            $CurrentUser.MFA = "Nog in te stellen door user"
        }Else{
            $CurrentUser.MFA = "Nog toe te wijzen"
        }
    }
    
    if( $CurrentUser.TPM  -eq "Not OK"){
          $CurrentUser.Status = "TPM not OK"
        }   
     
    if ( $CurrentUser.W10_Laptop){
        $W10Laptop++        
        $CurrentUser.Status = "Done"
    }

    $resultList += $CurrentUser

}

    $W10Ready 
    $W10Laptop 
        $resultList  | Export-Csv $OuputPath_BR10_csv -UseCulture -NoTypeInformation
				

  


        $Select = @("Status","Hello","Remark","SamAccountName","Afdeling","Name","Mail","Tel","LaptopName","BatchDatum"
                    "Error","App_OK", "App_NotOk", 
                    "W10_Laptop",  "LaptopCount", "LaptopType", "LaptopOS","UPN","MFA","TPM"
                     )
            		

        $SortOrder = @("Status", "BatchDatum", "Name")

        #SDE_Beheergroep,VM_Folder,	Naam, Team, SDE_Omschrijving, SDE_Domain,OS,SDE_OS,AD_OS,VM_OS, AD_LastLogon, SDE_IsActief, SDE_IsVirtueel, AD_IsActief, VM_IsActief,Collectie, AD_CN 
        $resultsorted = $resultList | Select-Object 	$Select

        $workbook = New-SNSExcelFromTemplate -TemplatePath  $OuputPath_BR10_tpl -ExcelPath  $OuputPath_BR10_xlsx 
        Save-SNSADObjectToExcelRange -InputObject $resultsorted -WorkBook $workbook -Sort $SortOrder -SheetName "Bron" -EntireRow $false -verbose -tempPath $OuputPath_BR10_csv  
        Save-ExcelWorkbook -workbook $workbook -close
        Close-Excel | Out-Null



