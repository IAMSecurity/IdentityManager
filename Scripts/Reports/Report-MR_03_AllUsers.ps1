
#Init


    . ..\..\data\SNSBankNV_WPS.ps1





    Import-Module SNSExcel
    
    $list = @() 
    $arUser = @{}
    $result = @{}
#Process 


    $list = @()
    ForEach($file in Get-ChildItem $OutputPath_History -filter "MR_02_Users_$CurrentYear-*.csv"){
    
        $month = "unknown"
        if($file.Name -match "MR_02_Users_$CurrentYear-(.*)\.csv"){
            $month = $Matches[1]
        }
        $temp = Import-Csv $file.FullName -UseCulture 
        $temp | Add-Member -MemberType NoteProperty -Name Month -Value $month -Force
     
        $list += $temp
    }
    $sortORder = @("AD_UserName","AD_UserNummer","AD_USerLastlogon","AD_UserLastlogonMonth","AD_UserAfdeling01","AD_UserAfdeling02","AD_UserAfdeling03","AD_UserAfdeling04","AD_UserAfdeling05","AD_UserRecentlyUsed","AD_UserEnabled","AD_UserKostenPlaats","AD_UserKostenOmschrijving","AD_UserOECode","AD_UserOEOmschrijving","Quest_ContractID","Quest_ContractName","Month")
    
    $resultSort =  $list | Select-Object  $sortORder|Sort-Object     $sortORder
    $resultSort | Export-Csv -UseCulture -NoTypeInformation $OuputPath_MR02Year_csv -Force


#


#Importeer Computer gegevens
  


    $workbook = New-SNSExcelFromTemplate -TemplatePath $OuputPath_MR03_tpl -ExcelPath $OuputPath_MR03_xlsx 

    Save-SNSADObjectToExcelRange -InputObject $resultSort  -WorkBook $workbook -Sort $SortOrder -SheetName "Bron" 
#Pivot Settings

    $XlPivotTableSourceType = [Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase 
    $XlPivotTableVersionList = [Microsoft.Office.Interop.Excel.XlPivotTableVersionList]::xlPivotTableVersion14
    $XlPivotTableSource = "Bron"
    $xlValueEquals = 7 
    $PivotTableCache = $workbook.PivotCaches().Create($XlPivotTableSourceType,$XlPivotTableSource ,$XlPivotTableVersionList)
    #$PivotTableHNWVerschil = $PivotTableCache.CreatePivotTable("HNWVerschil") 
    #$PivotTableHNWVerloop = $PivotTableCache.CreatePivotTable("HNWVerloop") 

     $PivotTableHNWVerschil =  $workbook.Sheets.Item("HNW Verschil").PivotTables("DraaitabelVerschil")
    $PivotTableHNWVerloop =  $workbook.Sheets.Item("HNW Verloop").PivotTables("DraaitabelVerloop")


     $PivotTableHNWVerschil.RefreshTable()

     $PivotTableHNWVerloop.RefreshTable()


#Pivot HNW VErschil
    $fieldsHNW = $PivotTableHNWVerschil.PivotFields()
    $fieldsHNW.Item("AD_UserAfdeling04").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
    $fieldsHNW.Item("AD_UserOEOmschrijving").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
    #$fieldsHNW.Item("Month").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
    #$fieldsHNW.Item("AD_RecentlyUsed").PivotFilters.Add($xlValueEquals,"True")
    $fieldsHNW.Item("Quest_ContractName").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlColumnField
    $fieldsHNW.Item("AD_UserRecentlyUsed").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlPageField
    $fieldsHNW.Item("AD_UserRecentlyUsed").PivotItems("FAlse").Visible = $false


    #$Sum = $PivotTableHNWVerschil.AddDataField($fieldsHNW.Item("Month"),"Verschil",[Microsoft.Office.Interop.Excel.XlConsolidationFunction]::xlCount)
    $fieldsHNW.Item("Month").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlColumnField
    
    $fieldsHNW = $PivotTableHNWVerschil.PivotFields()
    $fieldsHNW.Item("VErschil").calculation = [Microsoft.Office.Interop.Excel.XlPivotFieldCalculation]::xlDifferenceFrom
    $fieldsHNW.Item("VErschil").BaseField = "Month"
    $fieldsHNW.Item("VErschil").BaseItem = "1"


    
#Pivot HNW Verloop
    $fieldsHNW = $PivotTableHNWVerloop.PivotFields()
    $fieldsHNW.Item("AD_UserAfdeling04").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
    $fieldsHNW.Item("AD_UserOEOmschrijving").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
    #$fieldsHNW.Item("Month").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
    #$fieldsHNW.Item("AD_RecentlyUsed").PivotFilters.Add($xlValueEquals,"True")
    $fieldsHNW.Item("Quest_ContractName").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlColumnField
    $fieldsHNW.Item("AD_UserRecentlyUsed").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlPageField
    $fieldsHNW.Item("AD_UserRecentlyUsed").PivotItems("FAlse").Visible = $false


    $PivotTableHNWVerloop.AddDataField($fieldsHNW.Item("Month"),"Verloop",[Microsoft.Office.Interop.Excel.XlConsolidationFunction]::xlCount)
    $fieldsHNW.Item("Month").Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlColumnField
   





   
    

    $workbook.Sheets.Item("Info").Select()

    #CurrentDate CurrentDateMaand
     Set-RangeField -Excel $global:excel -RangeName "CurrentDate" -RangeValue $CurrentDate   
     Set-RangeField -Excel $global:excel -RangeName "CurrentDateMaand" -RangeValue $CurrentDateMaand   
     


    Save-ExcelWorkbook -workbook $workbook  -close
    Close-Excel
