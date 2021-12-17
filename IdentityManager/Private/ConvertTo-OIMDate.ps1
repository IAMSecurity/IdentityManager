Function ConvertTo-OIMDate{
    [CmdletBinding()]
    param (
        [Parameter()]
            [DateTime] $Date
    )
    $Date.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ", [cultureinfo]::CurrentCulture)
 
}