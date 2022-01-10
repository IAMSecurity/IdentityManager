Function Start-OIMScript{
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory=$true)]
            $ScriptName,
        [array]$Parameters
    )
    $body = @{parameters = $Parameters }
    $uri = "$Script:BaseURI/api/script/$ScriptName"
    if ($PSCmdlet.ShouldProcess($uri, "Run Script$ScriptName")) {
        Invoke-OIMRestMethod -Uri $uri  -Method put -Body $body -WebSession $Script:WebSession | out-Null

    }
}