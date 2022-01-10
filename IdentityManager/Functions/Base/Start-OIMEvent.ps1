Function Start-OIMEvent{
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
            $Object,
        [Parameter(Mandatory=$true)]
            $EventName,
            [hashtable]$Parameters = @{},
            [switch]$wait
    )
    BEGIN {	}#begin

	PROCESS {
        $body = @{parameters = $Parameters }
        ForEach($xObjectkey in $Object.xObjectKey){

			$xmlXObjectKey = 	[xml] $xObjectkey
		    $URI = "$Script:BaseURI/api/entity/$($xmlXObjectKey.key.T)/$($xmlXObjectKey.key.P)/event/$EventName"

            If($PSCmdlet.ShouldProcess($xmlXObjectKey , "Start-OIMEvent $EventName")){
                Invoke-OIMRestMethod -Uri $uri -Method put -Body $body -WebSession $Script:WebSession | out-Null

            }
        }

        if($wait){
            Wait-OIMJobQueue -JobChainName $EventName
        }


    }
}