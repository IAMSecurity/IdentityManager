Function Start-OIMMethod{
    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = "object")]
    Param(
        [Parameter(
            ParameterSetName = "Object",
            ValueFromPipeline=$true,
            Mandatory=$true)]
            $Object,
        [Parameter(Mandatory=$true)]
            $MethodName,
            [Parameter(
                ParameterSetName = "ObjectSingle",
                Mandatory=$true)]
                $ObjectTable,

            [Parameter(
                ParameterSetName = "ObjectSingle",
                Mandatory=$true)]
                $ObjectKey,
        [array]$Parameters = @()

    )
    BEGIN {	}#begin

	PROCESS {
        $body = @{parameters = $Parameters }

        switch ($PSCmdlet.ParameterSetName) {

			'Object' {

                ForEach($xObjectkey in $Object.xObjectKey){

                    $xmlXObjectKey = 	[xml] $xObjectkey
                    $URI = "$Script:BaseURI/api/entity/$($xmlXObjectKey.key.T)/$($xmlXObjectKey.key.P)/method/$methodname"

                    If($PSCmdlet.ShouldProcess($xmlXObjectKey , "Start-OIMMethod $MethodName")){
                        Invoke-OIMRestMethod -Uri $uri -Method put -Body $body  -WebSession $Script:WebSession | out-Null

                    }
                }
            }
            'ObjectSingle'{
                $URI = "$Script:BaseURI/api/entity/$ObjectTable/$ObjectKey/method/$methodname"

                If($PSCmdlet.ShouldProcess($xmlXObjectKey , "Start-OIMMethod $MethodName")){
                    Invoke-OIMRestMethod -Uri $uri -Method put -Body $body  -WebSession $Script:WebSession | out-Null

                }

            }

        }
    }
}