Function ConvertFrom-OIMDate{
    [CmdletBinding()]
    param (
        [Parameter()]
            [string] $Date
    )
 
    if(-not [string]::IsNullOrEmpty($date )){
        [datetime]::Parse( $Date)
    }
}