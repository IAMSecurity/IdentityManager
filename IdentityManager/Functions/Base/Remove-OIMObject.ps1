Function Remove-OIMObject{
    [CmdletBinding(SupportsShouldProcess)]
    Param(
        [Parameter(ValueFromPipeline=$true, Mandatory=$true)]
            $Object,
            [int]
            $max = 100

        )
    Begin{
		#Assert-VersionRequirement -RequiredVersion 8.1
}
    Process{

		ForEach($xObjectkey in $Object.xObjectKey){

			    $xmlXObjectKey = 	[xml] $xObjectkey

            $URI = "$Script:BaseURI/api/entity/$($xmlXObjectKey.key.T))/$($xmlXObjectKey.key.P))"
            if ($max -gt 0 -and $PSCmdlet.ShouldProcess($xmlXObjectKey , "removing object")) {
                $max--
                Invoke-OIMRestMethod -Uri  $URI -Method Delete
            }
		}
    }
    End{}
}