#Init
    . ..\..\data\SNSBankNV_WPS.ps1
    $list = @() 
    $arUser = @{}
    $result = @{}

#Importeer Computer gegevens
  # $dicSDE = Import-Csv $OuputPath_SDEPC -UseCulture  | Group-Object -Property SDE_Gebruikerscode -AsHashTable

    $AD = Import-CSV  $OuputPath_ADUsers -UseCulture    
    $AD | Add-Member -MemberType NoteProperty -Name SCCM_UserLogonMinutes -Value ""
    $AD | Add-Member -MemberType NoteProperty -Name SCCM_LastLogon -Value ""
    $AD | Add-Member -MemberType NoteProperty -Name SCCM_LogonTimes -Value ""
    $AD | Add-Member -MemberType NoteProperty -Name SDE_Naam -Value ""
    $AD | Add-Member -MemberType NoteProperty -Name SDE_InstallatieDatum -Value ""
    $AD | Add-Member -MemberType NoteProperty -Name SDE_Status -Value ""
        
        
        
    $SCCMUsage  = Import-CSV  $OuputPath_SCCMUsers -UseCulture  | Group-Object -Property SAMAccountName -AsHashTable

    ForEach($item in $AD){
        if($SCCMUsage.ContainsKey($item.AD_UserNummer)){
            $item.SCCM_UserLogonMinutes =  $SCCMUsage[$item.AD_UserNummer].CombinedTotalUserConsoleMinutes
            $item.SCCM_LastLogon =  $SCCMUsage[$item.AD_UserNummer].CombinedLastConsoleUse
            $item.SCCM_LogonTimes =  $SCCMUsage[$item.AD_UserNummer].CombinedNumberOfConsoleLogons

        }
       
    }
    # AD_UserNummer

    # SAMAccountName

    $sortORder = @("AD_UserName","AD_UserNummer","AD_USerLastlogon","AD_UserLastlogonMonth",
            "AD_UserAfdeling01","AD_UserAfdeling02","AD_UserAfdeling03","AD_UserAfdeling04","AD_UserAfdeling05",
            "AD_UserRecentlyUsed","AD_UserEnabled","AD_UserKostenPlaats","AD_UserKostenOmschrijving",
            "AD_UserOECode","AD_UserOEOmschrijving","Quest_ContractID","Quest_ContractName","Quest_UserAfdeling01",
            "Quest_UserAfdeling02","Quest_UserAfdeling03","Quest_UserAfdeling04","Quest_UserAfdeling05",
            "Quest_UserAfdeling06","Quest_UserAfdeling07","Quest_UserAfdeling08", 
            "SCCM_UserLogonMinutes" ,"SCCM_LastLogon","SCCM_LogonTimes","SDE_Naam","SDE_InstallatieDatum","SDE_Status")
    $resultSort =  $AD | Select-Object  $sortORder|Sort-Object     $sortORder
    
    $resultsort | Export-Csv -UseCulture -NoTypeInformation -Path $OuputPath_MR02Month_csv -Force 

    Write-Verbose "Creating Excel file $PathOutput"
    $workbook = New-SNSExcelFromTemplate -TemplatePath $OuputPath_MR02_tpl -ExcelPath $OuputPath_MR02_xlsx 
    Save-SNSADObjectToExcelRange -InputObject $resultSort  -WorkBook $workbook -Sort $SortOrder -SheetName "Bron" 
    






#Pivot Settings

    $XlPivotTableSourceType = [Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase 
    $XlPivotTableVersionList = [Microsoft.Office.Interop.Excel.XlPivotTableVersionList]::xlPivotTableVersion14
    $XlPivotTableSource = "Bron"
    $xlValueEquals = 7 
    $PivotTableCache = $workbook.PivotCaches().Create($XlPivotTableSourceType,$XlPivotTableSource ,$XlPivotTableVersionList)
    $PivotTableHNWGebruikers = $PivotTableCache.CreatePivotTable("HNWGebruikers") 
    $PivotTableHNWContract = $PivotTableCache.CreatePivotTable("HNWContract") 
    $PivotTableHNWQuest = $PivotTableCache.CreatePivotTable("HNWQuest") 
   

#Pivot HNW Gebruikers

    $fieldsHNW = $PivotTableHNWGebruikers.PivotFields()
    $fieldsHNW.Item("AD_UserLastlogonMonth").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
    $fieldsHNW.Item("AD_UserName").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
    #$fieldsHNW.Item("AD_RecentlyUsed").PivotFilters.Add($xlValueEquals,"True")
    $fieldsHNW.Item("AD_UserRecentlyUsed").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlPageField
    $fieldsHNW.Item("AD_UserRecentlyUsed").PivotItems("FAlse").Visible = $false

#Pivot HNW Contract

    $fieldsHNW = $PivotTableHNWContract.PivotFields()
    $fieldsHNW.Item("AD_UserAfdeling04").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
    $fieldsHNW.Item("AD_UserOEOmschrijving").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
    $fieldsHNW.Item("AD_Username").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
    $fieldsHNW.Item("Quest_ContractName").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlColumnField
    #$fieldsHNW.Item("AD_RecentlyUsed").PivotFilters.Add($xlValueEquals,"True")
    $fieldsHNW.Item("AD_UserRecentlyUsed").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlPageField
    $fieldsHNW.Item("AD_UserRecentlyUsed").PivotItems("FAlse").Visible = $false
        
#Pivot HNW Quest

    $fieldsHNW = $PivotTableHNWQuest.PivotFields()
    $fieldsHNW.Item("Quest_UserAfdeling03").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
    $fieldsHNW.Item("Quest_UserAfdeling04").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
    $fieldsHNW.Item("AD_UserName").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
    #$fieldsHNW.Item("AD_RecentlyUsed").PivotFilters.Add($xlValueEquals,"True")
    $fieldsHNW.Item("AD_UserRecentlyUsed").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlPageField
    $fieldsHNW.Item("AD_UserRecentlyUsed").PivotItems("FAlse").Visible = $false
    

    $workbook.Sheets.Item("Info").Select()

    #RapportDatum RapportDatumMaand
     Set-RangeField -Excel $global:excel -RangeName "RapportDatum" -RangeValue $RapportDatum   
     Set-RangeField -Excel $global:excel -RangeName "RapportDatumMaand" -RangeValue $RapportDatumMaand   



    Save-ExcelWorkbook -workbook $workbook  -close
    Close-Excel
