
. ..\..\Data\SNSBankNV_WPS.ps1

$RESItems = Import-csv -Path $OutputPathRESObjects -UseCulture
$list = @()

###### 

    $webclient = New-Object System.Net.WebClient;
    $webclient.UseDefaultCredentials = $true

    $temppath = "C:\Windows\Temp\Applicatielijst.xlsx"
    $webclient.DownloadFile($RESApplicatieLijst,$temppath)


    $connection = New-Object System.Data.OleDb.OleDbConnection

    $connectstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$temppath;Extended Properties='Excel 12.0 Xml;HDR=YES;'"
    $connection.ConnectionString = $connectstring



    $connection.open()


    #$connection.GetSchema("columns") | ft TABLE_NAME, COLUMN_NAME
    $cmdObject = New-Object System.Data.OleDb.OleDbCommand

 
    $query = "Select * from [Bron`$]"
    $cmdObject.CommandText = $query
    $cmdObject.CommandType = "Text"
    $cmdObject.Connection = $connection

    $dataReader = $cmdObject.ExecuteReader()

    $dicApplicatielijst = @{}
    While ($dataReader.Read()) {
        $temp = $datareader
  
        $shortname = $datareader.Item("ShortName")
        $afdeling = $datareader.Item("Afdeling Eigenaar")
        $contact = $datareader.Item("Eigenaar")
        $dicApplicatielijst.Add($shortname ,@{Afdeling=$afdeling;Contact=$contact})

    }
    $dataReader.Close()

    $connection.Close()



######
$dicGAP = @{}
ForEach($group in  Get-ADGroup -filter {Info -like "* -*"} -Properties info){
    if($group.info -match "(.*) -(.*)"){
     $dicGAP.Add($group.Name.ToUppeR() , $Matches[1])
    }
}

ForEach($item in $RESItems){
    
    # Skip if it is a PAT application
    if($item.RESFolderName.Contains("\PAT")){continue}


    
    $RESDLGroup =  ""
    $RESName =  $item.name
    $RESTitle =  $item.title
    $RESType =  $item.RESTypeName
    $arGroups = @()
    if(-not [string]::IsNullOrEmpty($item.af_authorizedgroup)){
        $arGroups = $item.af_authorizedgroup.Split("#") 
    }
    
    if(-not [string]::IsNullOrEmpty($item.ac_grouplist)){
        $arGroups += $item.ac_grouplist.Split("#") 
    
    }

    if($arGroups.Count -eq 0 ){
        continue
    }
    
     

    ForEach($group in $arGroups){
        if($group -match "VERZ\\(D.*)"){
            $RESDLGroup = $Matches[1]
            if([string]::IsNullOrEmpty($RESDLGroup)){continue}

            ForEach($globalgroup in Get-ADGroupMember $RESDLGroup){
            $APP = $Afdeling= $contact =""
                $name = $globalgroup.Name.ToUpper()
                if($dicGAP.ContainsKey( $name)){
                    $APP = $dicGAP[$name]
                }
                
                if($dicApplicatielijst.ContainsKey( $APP)){
                    $Afdeling = $dicApplicatielijst[$APP].Afdeling
                    $contact  = $dicApplicatielijst[$APP].Contact
                }

                $list += New-Object -TypeName PSObject -Property @{
                    Global = $globalgroup.Name
                    DomainLocal = $RESDLGroup
                    RESName = $RESName
                    RESTitle = $RESTitle
                    RESType = $RESType
                    APP = $APP
                    AFdeling= $afdeling
                    Contact = $contact
                }
            }
        }else{
            if(-not $group.Contains("{")){
                Write-warning $group
            }
        }
    }

    
   

}


  

  #### Excel

    $sortORder = @("RESTitle","AFdeling","Global","RESName","Contact","RESType","APP","DomainLocal")



    $export = $list| Select-Object  $sortORder

    $resultSort =  $export  |Sort-Object     $sortORder
    $workbook = New-SNSExcelFromTemplate -TemplatePath $OuputPath_BR04_tpl -ExcelPath $OuputPath_BR04_xlsx 
    Save-SNSADObjectToExcelRange -InputObject $resultSort  -WorkBook $workbook -Sort $SortOrder -SheetName "Bron" 
#Pivot Settings

    $XlPivotTableSourceType = [Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase 
    $XlPivotTableVersionList = [Microsoft.Office.Interop.Excel.XlPivotTableVersionList]::xlPivotTableVersion14
    $XlPivotTableSource = "Bron"
    $xlValueEquals = 7 
    $PivotTableCache = $workbook.PivotCaches().Create($XlPivotTableSourceType,$XlPivotTableSource ,$XlPivotTableVersionList)
    $PivotTableRESApplicatie = $PivotTableCache.CreatePivotTable("RESApplicatie") 
    

#Pivot HNW Laptops

    $fieldsHNW = $PivotTableRESApplicatie.PivotFields()
    $fieldsHNW.Item("Afdeling").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
    $fieldsHNW.Item("APP").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
    $fieldsHNW.Item("Global").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
    
    $fieldsHNW.Item("DomainLocal").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
    #$fieldsHNW.Item("AD_RecentlyUsed").PivotFilters.Add($xlValueEquals,"True")
    $fieldsHNW.Item("RESType").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlColumnField
    $fieldsHNW.Item("APP").ShowDetail = $false
#Pivot SNS Laptops
   
    

    $workbook.Sheets.Item("Info").Select()

    #RapportDatum RapportDatumMaand
     Set-RangeField -Excel $global:excel -RangeName "RapportDatum" -RangeValue $RapportDatum   
     Set-RangeField -Excel $global:excel -RangeName "RapportDatumMaand" -RangeValue $RapportDatumMaand   
     


    Save-ExcelWorkbook -workbook $workbook  -close
    Close-Excel

<#
  $xlValueEquals = 7 
        $field.PivotFilters.Add($xlValueEquals,"True")
#Set-alignmentColumn -Excel $global:excel -WorkBook $workbook -SheetName "Beheer Gebruikers"  -Column "F" -Alignment "-4108"
#Set-alignmentColumn -Excel $global:excel -WorkBook $workbook -SheetName "Beheer Gebruikers"  -Column "H" -Alignment "-4108"

$result.Values | Select-Object Naam,AfdelingAll,AD_OS,AD_Actief,AD_SAM,AD_ObjectGUID,
            AD_ObjectClass,AD_SID,AD_LastLogon,AD_LastLogonMonth,AD_DN,AD_CN,AD_DNS,AD_IsActief,SDE_IsActief,
            SDE_Categorie,SDE_Omgeving,SDE_SubCategroie,SDE_AssetTag,SDE_ObjectSerieNR,
                SDE_WhenCreated,SDE_Status,SDE_Beheergroep,SDE_Toelichting,
                SDE_WhenModified,SDE_IsVirtueel,SDE_Locatie,SDE_ObjectCode,
                SDE_ObjectOmschrijving, SDE_Afdelingsnaam,    SDE_Afdelingscode , SDE_Gebruikerscode,SDE_Gebruiker,
            AD_UserNummer,AD_UserName	,AD_UserEnabled ,AD_UserAfdeling,AD_UserAfdeling01,AD_UserAfdeling02	,
                AD_UserAfdeling03	,AD_UserAfdeling04	,AD_UserAfdeling05	,AD_UserLastlogon ,AD_UserLastlogonMonth ,AD_UserRecentlyUsed  ,   
            GO_Afdeling ,GO_Naam,GO_PersoneelsCode,GO_Datum ,
            RES_GebruikersCode ,    RES_Naam, RES_Afdeling,    RES_Time | Export-Csv D:\Automating\Output\KeyControl\MR_01_Laptops.csv -UseCulture -NoTypeInformation



#>