function Install-OIMFunction
{
    Param(
        $Namespace  = @("Person","Eset","Org","ADS","AAD","UNS","TSB")
    )
    $dicFunction = @{}
    ForEach($NamespaceItem in $Namespace ){
        ForEach($Table in Get-OIMObject DialogTable -Where "UID_DialogTable like '%$NameSpaceItem%'"){
            $TableName  = ($table.UID_DialogTable -split "-")[2]
            $columns = @()
            ForEach($column in Get-OIMObject DialogColumn -Where "UID_DialogTable = '$($Table.UID_DialogTable)'"){
                if(-not $column.ColumnName.StartsWith("X") -or $column.ColumnName -eq "XIsInEffect"){
                    $columns += $column.ColumnName
                }
            }
            $dicFunction[$tableName ] =       $columns

        }
    }

   # creates a new function dynamically
   ForEach($key in $dicFunction.Keys){
    # GET
        $Name = "Get-OIM$Key"
        $Param =  "`$"  + ($dicFunction[$key] -join ",`$")
        $Code = {
            [CmdletBinding()]
        Param($TEMP)
        $whereParam = $PSBoundParameters | Get-OIMParameter | ConvertTo-SQLWhereString -NoEscape

        Get-OIMObject -ObjectName $Key -where $whereParam

        }
        $code = $code -replace "\$`TEMP",$param
        $code = $code -replace "\$`Key",$Key
        # create new function in function: drive and set scope to "script:"
        $null = New-Item -Path function: -Name "global:$Name" -Value $Code -Force
    # Set
        $Name = "Set-OIM$Key"
        $Param =  "$"  + ($dicFunction[$key] -join ",`$")
        $Code = {
            [CmdletBinding()]
            Param(
                [parameter(
                    Mandatory = $false,
                    ValueFromPipeline = $true
                )]
                $Object,
                $TEMP)
            $Object | Set-OIMObject -Properties  $PSBoundParameters

        }
        $code = $code -replace "\$`TEMP",$Param
        # create new function in function: drive and set scope to "script:"
        $null = New-Item -Path function: -Name "global:$Name" -Value $Code -Force
    # New
        $Name = "New-OIM$Key"
        $Param =  "$"  + ($dicFunction[$key] -join ",`$")
        $Code = {
            [CmdletBinding()]
            Param($TEMP)
            New-OIMObject -ObjectName $key -Properties  $PSBoundParameters

        }
        $code = $code -replace "\$`TEMP",$Param
        $code = $code -replace "\$`Key",$Key
        # create new function in function: drive and set scope to "script:"
        $null = New-Item -Path function: -Name "global:$Name" -Value $Code -Force

   }
}