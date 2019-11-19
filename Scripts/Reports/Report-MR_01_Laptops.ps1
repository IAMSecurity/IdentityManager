#Init
    . ..\..\data\SNSBankNV_WPS.ps1

    $list = @() 
    $arUser = @{}
    $result = @{}
#Importeer Computer gegevens

    $AD =   Import-CSV $OuputPath_ADComputerWorkstations -UseCulture
    #$SDE =  Import-CSV $OuputPath_SDEPC -UseCulture
    #$GO =   Import-CSV $OuputPath_GO -UseCulture
    $RES =  Import-CSV $OuputPath_RESUserToComputer -UseCulture
    $SCCM = Import-CSV $OuputPath_SCCMUsers -UseCulture
    $Sophos = Import-CSV $OuputPath_Sophos -UseCulture
    $SNOWPC = Import-Csv $OuputPath_SNOWPC -UseCulture 

    $list +=$AD
    $list += $SNOWPC
    #$list += $GO
    $list += $RES
    $list += $Sophos

    $list = $list | Where-Object {$_.Name -like "ON*" -or  $_.Name -like "DV*"}

#Importeer User gegevens
    $ADUser = Import-CSV  $OuputPath_ADUsers -UseCulture
    ForEach($item in $ADUser){
        $arUser.Add($item.AD_UserNummer,$item)    | Out-Null
    }


#Custom result object
    $properties = @{}
    $properties.Add("AfdelingAll","")
    ForEach($item in $AD |Get-Member -MemberType NoteProperty){
       if( -not $properties.ContainsKey( $item.Name)){
         $properties.Add($item.Name,"")
       }
    }
   
    ForEach($item in $SNOWPC |Get-Member -MemberType NoteProperty){
       if( -not $properties.ContainsKey( $item.Name)){
         $properties.Add($item.Name,"")
       }
    }
    <#
    
    ForEach($item in $GO |Get-Member -MemberType NoteProperty){8
       if( -not $properties.ContainsKey( $item.Name)){
         $properties.Add($item.Name,"")
       }
    }
    #>
    ForEach($item in $Sophos |Get-Member -MemberType NoteProperty){
       if( -not $properties.ContainsKey( $item.Name)){
         $properties.Add($item.Name,"")
       }
    }
    ForEach($item in $RES |Get-Member -MemberType NoteProperty){
       if( -not $properties.ContainsKey( $item.Name)){
         $properties.Add($item.Name,"")
       }
    }
    ForEach($item in $ADUser |Get-Member -MemberType NoteProperty){
       if( -not $properties.ContainsKey( $item.Name)){
         $properties.Add($item.Name,"")
       }
    }
    
#Nalopen van elke computer 
    ForEach($item in $list){
        $NAAM = $null
        if(-not [string]::IsNullOrEmpty($item.Name)){
            $naam = $item.Name.ToUpper()
        } 
        if(-not [string]::IsNullOrEmpty($item.Naam)){
            $naam = $item.Naam.ToUpper()
        }

        $item | Add-Member -NotePropertyName Name -NotePropertyValue $NAAM -force

        if(!$result.ContainsKey($item.Name)){        
            $temp = New-Object -TypeName PSObject -Property $properties
            $result.Add($item.Name,$temp)
        }else{
             $temp = $result[$item.Name]
        }

        ForEach($prop in $item |Get-Member -MemberType NoteProperty){
            $temp.$($prop.Name) =$item.$($prop.Name)
        }

    
         $result[$item.Name] = $temp 
    }

#unieke laptops ophalen
    $keys = @()
    ForEach($key in $result.Keys){
        $keys += $key.ToString()
    }

#Alle gegevens opslaan
    ForEach($key in $keys ){
        $temp = $result[$key]

        $temp.AfdelingAll =  $temp.SNOW_EigenaarAfdeling
        if([string]::IsNullOrEmpty($temp.AfdelingAll)){
            $temp.AfdelingAll =  $temp.GO_Afdeling
        }
        if([string]::IsNullOrEmpty($temp.AfdelingAll)){
            $temp.AfdelingAll =  $temp.AD_UserAfdeling
        }

    
        if([string]::IsNullOrEmpty($temp.AfdelingAll)){
            $temp.AfdelingAll =  $temp.RES_Afdeling
        }



    
        if(![string]::IsNullOrEmpty($temp.SDE_GebruikersCode) -and $arUser.ContainsKey($temp.SDE_GebruikersCode) ){
            $item = $arUser[$temp.SDE_GebruikersCode]
            ForEach($prop in $item |Get-Member -MemberType NoteProperty){
      
                $temp.$($prop.Name) =$item.$($prop.Name)
            }
        }
        $result[$key] = $temp 
    }

  


    $sortORder = @("Name","AfdelingAll","AD_OS","AD_Actief","AD_SAM","AD_ObjectGUID",
            "AD_ObjectClass","AD_SID","AD_LastLogon","AD_LastLogonMonth","AD_RecentlyUsed","AD_DN","AD_CN","AD_DNS","AD_IsActief","SNOW_IsActief",
            "SNOW_ObjectCode",
            "SNOW_Status","SNOW_EigenaarGroep","SNOW_Toelichting",
            "SNOW_WhenModified","SNOW_Location","SNOW_Type",
            "AD_UserNummer","AD_UserName","AD_UserEnabled","AD_UserAfdeling","AD_UserAfdeling01","AD_UserAfdeling02",
            "AD_UserAfdeling03","AD_UserAfdeling04","AD_UserAfdeling05","AD_UserLastlogon","AD_UserLastlogonMonth","AD_UserRecentlyUsed",
            "GO_Afdeling","GO_Naam","GO_PersoneelsCode","GO_Datum","GO_Akkoord",
            "RES_GebruikersCode","RES_Naam","RES_Afdeling","RES_Time",
            "Sophos_Domain","Sophos_ServicePack","Sophos_Description","Sophos_Managed","Sophos_LastLogonUser",
            "Sophos_InstalledAU","Sophos_InstalledSAV","Sophos_InstalledOnAccess","Sophos_InstalledWeb",
            "Sophos_VersionSoftware","Sophos_VersionSAV","Sophos_VersionEnging","Sophos_VersionVirusData",
            "Sophos_VersionAgent","Sophos_PrimaryLocation","Sophos_TimeInstalled","Sophos_TimeLastScan",
            "Sophos_TimeChangedOnEP","Sophos_TimeLastUpdate"
)

    $export = $result.Values | Select-Object  $sortORder

    $resultSort =  $export  |Sort-Object     $sortORder

    $resultSort| Export-Csv -LiteralPath $OuputPath_MR01_csv  -UseCulture -NoTypeInformation
    $workbook = New-SNSExcelFromTemplate -TemplatePath $OuputPath_MR01_tpl -ExcelPath $OuputPath_MR01_xlsx 
    Save-SNSADObjectToExcelRange -InputObject $resultSort  -WorkBook $workbook -Sort $SortOrder -SheetName "Bron" 
#Pivot Settings

    $XlPivotTableSourceType = [Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase 
    $XlPivotTableVersionList = [Microsoft.Office.Interop.Excel.XlPivotTableVersionList]::xlPivotTableVersion14
    $XlPivotTableSource = "Bron"
    $xlValueEquals = 7 
    $PivotTableCache = $workbook.PivotCaches().Create($XlPivotTableSourceType,$XlPivotTableSource ,$XlPivotTableVersionList)
    $PivotTableHNWLaptops = $PivotTableCache.CreatePivotTable("HNWLaptops") 
    $PivotTableSNSLaptops = $PivotTableCache.CreatePivotTable("SNSLaptops") 
    

#Pivot HNW Laptops

    $fieldsHNW = $PivotTableHNWLaptops.PivotFields()
    $fieldsHNW.Item("AD_LAstLogonMonth").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
    $fieldsHNW.Item("AD_CN").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
    #$fieldsHNW.Item("AD_RecentlyUsed").PivotFilters.Add($xlValueEquals,"True")
    $fieldsHNW.Item("AD_RecentlyUsed").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlPageField

    $fieldsHNW.Item("AD_RecentlyUsed").PivotItems("FAlse").Visible = $false
    $fieldsHNW.Item("AD_RecentlyUsed").PivotItems("(blank)").Visible = $false
#Pivot SNS Laptops

    $fieldsSNS = $PivotTableSNSLaptops.PivotFields()
    $fieldsSNS.Item("AfdelingAll").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
    $fieldsSNS.Item("AD_CN").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField

    ForEach($value in  $fieldsSNS.Item("AfdelingAll").PivotItems()){
        if( $value.Caption.Contains("SNS Winkel") -or   $value.Caption.Contains("SNS Retail")){

        }else{
            $value.Visible = $false
        }
    }
    $fieldsSNS.Item("AD_RecentlyUsed").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlPageField

    $fieldsSNS.Item("AD_RecentlyUsed").PivotItems("FAlse").Visible = $false
    $fieldsSNS.Item("AD_RecentlyUsed").PivotItems("(blank)").Visible = $false
    

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