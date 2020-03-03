#requires -version 4
<#
.SYNOPSIS
  Close
.DESCRIPTION
  <Brief description of script>
.PARAMETER <Parameter_Name>
  <Brief description of parameter input required. Repeat this attribute if required>
.INPUTS
  <Inputs if any, otherwise state None>
.OUTPUTS
  <Outputs if any, otherwise state None>
.NOTES
  Version:        1.0
  Author:         <Name>
  Creation Date:  <Date>
  Purpose/Change: Initial script development
.EXAMPLE
  <Example explanation goes here>
  
  <Example goes here. Repeat this attribute for more than one example>
#>


Function New-rExcelWorkBook($path,$sheetname= "Default"){
    if($null -eq   $Global:excel){$Global:excel = New-Object -ComObject Excel.Application }

    $excel.Visible = $true
    $workbook = $excel.Workbooks.add()
    $sheet1 = $workbook.worksheets.Item(1)
    $sheet1.name = $sheetname
    $workbook.SaveAs($path)
    return $workbook
}


 Function New-rExcelFromTemplate{
    Param(
           $TemplatePath,
           $ExcelPath, 
           $Title,
           $Subtitle,
           $Author,
           $Contracter,
           $Owner,
           $Project,
           $ChangedDate,
           $Information,
           [switch]
           $ExcelNotVisible
       )
       Begin{        
           [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Excel") | Out-Null
           if($Global:excel.Application -eq $null){
               $Global:excel = New-Object -comobject Excel.Application
               if(! $ExcelNotVisible){$Global:excel.Visible = $true}
               $Global:excel.DisplayAlerts = $False
           }
       }
   
       Process{
   
   
       Copy-Item -Path $TemplatePath -Destination $ExcelPath -Force |out-Null
       $Workbook = $Global:excel.Workbooks.Open($ExcelPath)
   
       if(![string]::isNullorEmpty($Title)){Set-RangeField -Excel $Global:excel -RangeName "Title" -RangeValue $Title} 
       if(![string]::isNullorEmpty($Subtitle)){Set-RangeField -Excel $Global:excel -RangeName "Subtitle" -RangeValue $Subtitle} 
       if(![string]::isNullorEmpty($Author)){Set-RangeField -Excel $Global:excel -RangeName "Author" -RangeValue $Author} 
       if(![string]::isNullorEmpty($Contracter)){Set-RangeField -Excel $Global:excel -RangeName "Contracter" -RangeValue $Contracter} 
       if(![string]::isNullorEmpty($Project)){Set-RangeField -Excel $Global:excel -RangeName "Project" -RangeValue $Project} 
       if(![string]::isNullorEmpty($Owner)){Set-RangeField -Excel $Global:excel -RangeName "Owner" -RangeValue $Owner} 
       if(![string]::isNullorEmpty($ChangedDate)){Set-RangeField -Excel $Global:excel -RangeName "ChangedDate" -RangeValue $ChangedDate} 
       if(![string]::isNullorEmpty($Information)){Set-RangeField -Excel $Global:excel -RangeName "Information" -RangeValue $Information} 
       return $Workbook
   
       }#End Process
   
   
   
   }
   
Function Close-rExcel(){
    Try{
        $Global:excel.Quit() | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Global:excel)
        $Global:excel = $null
        }catch{
        Write-Warning "Did not close excel"
        }
}

function Save-rExcelObjectToPivotTable {
    [CmdletBinding()]
    param
    (
    [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
        HelpMessage='Specify an object?')]
    $InputObject,
    $WorkBook,
    $SheetName = "Gegevens",       
    $PivotDestination = "PivotDestination",           
    $RangeName = "SourceData",    
    $RangeTableName = "SourceTable",
    $tempPath = "C:\Windows\Temp\tmp_Save-rExcelObjectToPivotTable.csv",
    $Sort ="",
    [string[]] $PivotColumns,
    [string[]] $PivotRows,
    [string[]] $PivotData,
    [string[]] $PivotCollapse
        
    )
Process{
    
    Write-Host "Saving Object To Excel Pivot"
        if(-not [string]::IsNullOrEmpty($sort)){
            $InputObject = $InputObject | Select-Object $Sort 
        }
        
        $InputObject | Export-Csv -Path $tempPath -UseCulture -NoTypeInformation 
        

        
    Write-Host "Copy to workbook" 
        $csv = $global:excel.Workbooks.Open($tempPath)   
        $range = $WorkBook.Sheets.Item($SheetName).Range( $RangeName ) 
        $rangeTable = $WorkBook.Sheets.Item($SheetName).Range( $RangeTableName ) 
        $csv.ActiveSheet.UsedRange.Select()
        $csv.ActiveSheet.UsedRange.Copy($range)
        $csv.Close()
        
    Write-Host "Closing sheet" 
        #Remove Header             
        $range.EntireColumn.AutoFit()
        $range.Rows.Item(1).EntireRow.Delete()

        $rangeTable.EntireColumn.AutoFit()



    Write-Host "Create Pivot table"
        $XlPivotTableSourceType = [Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase 
        $XlPivotTableVersionList = [Microsoft.Office.Interop.Excel.XlPivotTableVersionList]::xlPivotTableVersion14
        $XlPivotTableSource = $RangeTableName

        $PivotTableCache = $workbook.PivotCaches().Create($XlPivotTableSourceType,$XlPivotTableSource ,$XlPivotTableVersionList)
        $PivotTable = $PivotTableCache.CreatePivotTable($PivotDestination ) 


        ForEach($field in $PivotTable.PivotFields()){
            if($field.Name -in $PivotRows){
                $field.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
            }
            if($field.Name -in $PivotData){
                $field.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
            }
            if($field.Name -in $PivotColumns){
                $field.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlColumnField
            }
            
            
        }
        
        
        ForEach($field in $PivotTable.PivotFields()){
            if($field.Name -in $PivotCollapse){
                $field.ShowDetail = $false
            }        
        }
            

    
    #Layou t
        $PivotTable.ShowTableStyleRowStripes = $true

    return $PivotTable
}#End Process
}

Function Save-rExcelWorkbook{
    Param(
        $workbook,
        [switch]   $close 
    )
    Process{
        if($null -ne $workbook){
            $workbook.Save()
            if($close){
                $workbook.Close()                
            }
        }
    }

}

Function Save-rExcelObjectRange {
    Param (
        $InputObject,
        $WorkBook,
        $SheetName = "Gegevens",      
        $Sort ="",  
        $RangeName = "SourceData",    
        $RangeTableName = "",
        $tempPath = "C:\Windows\Temp\tmp_Save-rExcelObjectRange.csv",
        $noHeader= $true,
        $EntireRow = $true
    )
    Process{
            Write-Verbose "Saving Object To Excel Pivot"
            if($sort -ne ""){
                $InputObject |Select-Object $Sort | Export-Csv -Path $tempPath -UseCulture -NoTypeInformation 
            }else{
                $InputObject | Export-Csv -Path $tempPath -UseCulture -NoTypeInformation 
            }


            
            $csv = $global:excel.Workbooks.Open($tempPath)      
            Write-Verbose "Copy to workbook" 
            Try{
                
                $sheet = $WorkBook.Sheets.Item($SheetName)
                
                $range = $sheet.Range( $RangeName ) 
                if($range -ne $null){
                    $csv.ActiveSheet.UsedRange.Select()| out-Null
                    $csv.ActiveSheet.UsedRange.Copy($range)| out-Null
                    if($noHeader -and $EntireRow ){ $range.Rows.Item(1).EntireRow.Delete() | out-Null}
                    if($noHeader -and !($EntireRow)){ $range.Rows.Item(1).Delete() | out-Null}
                }else{
                    Write-Warning "Range not found: $range"
                }
                
                $csv.Close() | out-Null

            }catch{
                $csv.Close() | out-Null
                Write-Warning "Sheet not found: $SheetName $RangeName"
            }
            
            
            Try{            
                if(![string]::isNullOrEmpty( $RangeTableName )){
                    $rangeTable = $sheet.Range( $RangeTableName ) 
                    $rangeTable.EntireColumn.AutoFit() | Out-Null
                }
            }catch{
                Write-Warning "Could not autofit rangetable : $RangeTableName "

            }
    }

}

Function Set-rExcelRangeField {
    Param(
        $Excel,
        $RangeName,
        $RangeValue

    )
    Process{
        Try{
            $Excel.Range($RangeName).value2 = $RangeValue
        }catch{
            Write-Warning "Could not set value: $RangeName range: $RangeValue"
        }


    }

}

Function Set-rExcelAlignColumn {
#PvH 29-08-2014 toegevoegd voor de alignement van een Column.
    Param(
        $Excel, 
        $WorkBook,
        $SheetName,
        $Column,
        $Alignment
    )
    Process{
        Try{
            $ws = $Workbook.Sheets.Item($Sheetname)
            $Range = $ws.Cells.Item(1,$Column).EntireColumn
            If($Alignment -eq "-4108"){
                #Center
                    Write-Verbose "Column Center"
                $Range.HorizontalAlignment = -4108
            }
            If($Alignment -eq "-4152"){
                #Right
                    Write-Verbose "Column Right"
                $Range.HorizontalAlignment = -4152
            }
            If($Alignment -eq "-4131"){
                #Left
                    Write-Verbose "Column Left"
                $Range.HorizontalAlignment = -4131
            }
        }catch{
            Write-Warning "Could not Allign: $Column"
        }

    }

}

Function Set-rExcelAlignCell {
    Param(
        $Excel, 
        $WorkBook,
        $SheetName,
        $Column,
        $Cell,
        $Alignment
    )
    Process{
        Try{
            $ws = $Workbook.Sheets.Item($Sheetname)
            $Range = $ws.Cells.Item($cell,$Column)
            If($Alignment -eq "-4108"){
                #Center
                    Write-Verbose "Cell Center"
                $Range.HorizontalAlignment = -4108
            }
            If($Alignment -eq "-4152"){
                #Right
                    Write-Verbose "Cell Right"
                $Range.HorizontalAlignment = -4152
            }
            If($Alignment -eq "-4131"){
                #Left
                    Write-Verbose "Cell Left"
                $Range.HorizontalAlignment = -4131
            }
        }catch{
            Write-Warning "Could not Allign: $Cell,$Column"
        }

    }

}

Function Copy-rExcelWorksheet {
    Param(
        $Workbook,
        $SourceName ="Template",
        $DestinationName
    )
    Process{
        #Copieer een werkblad in de huidige sheet en geef het een naam.
        $Workbook.WorkSheets.Item($SourceName).Copy($Workbook.WorkSheets.Item(1))
        $ws = $Workbook.worksheets.Item(1)
        $ws.Name = $DestinationName
    }
}