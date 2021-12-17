Function Set-OIMObject{
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
            $Object,
        [Parameter(Mandatory=$true)]
            [hashtable] $Properties
    )

    Begin{}
    Process{
        $body = @{values = $Properties }
		ForEach($xObjectkey in $Object.xObjectKey){

			$xmlXObjectKey = 	[xml] $xObjectkey

            $URI = "$Script:BaseURI/api/entity/$($xmlXObjectKey.key.T))/$($xmlXObjectKey.key.P))"
			if ($PSCmdlet.ShouldProcess($item.uri , "Update Object")) {
                Invoke-OIMRestMethod -Uri $URI -Method Put -Body $body  |Out-Null
			}
		}

    }
    End{}
}