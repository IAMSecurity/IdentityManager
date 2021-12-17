Function Start-OIMMethod{
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
            $Object,
        [Parameter(Mandatory=$true)]
            $MethodName,
            [array]$Parameters = @()
    )
    BEGIN {	}#begin

	PROCESS {
        $body = @{parameters = $Parameters }
        ForEach($xObjectkey in $Object.xObjectKey){

			$xmlXObjectKey = 	[xml] $xObjectkey
		    $URI = "$Script:BaseURI/api/entity/$($xmlXObjectKey.key.T)/$($xmlXObjectKey.key.P)/method/$methodname"

            If($PSCmdlet.ShouldProcess($xmlXObjectKey , "Start-OIMMethod $MethodName")){
                Invoke-OIMRestMethod -Uri $uri -Method put -Body $body  | out-Null

            }
        }


    }
}