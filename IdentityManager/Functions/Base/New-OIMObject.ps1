function New-OIMObject {
	[CmdletBinding(SupportsShouldProcess)]
	param(
		[Parameter(Mandatory=$true)]
        [Alias("Type","Object")]
            $ObjectName,
        [Parameter(Mandatory=$true)]
            [hashtable] $Properties,
            [switch] $checkexists
	)

	BEGIN {	}#begin

	PROCESS {

		$URI = "$Script:BaseURI/api/entity/$($ObjectName)"
		#Get request parameters
		$body = @{values = $Properties }

        If($checkexists -and $Properties.Containskey("Ident_$objectname")){
            $OIMResponse = GEt-OIMObject -objectname $objectname -where "Ident_$objectname = '$($Properties["Ident_$objectname"])'"
            if($null -ne  $OIMResponse){
                return $OIMResponse
            }
        }else{
			if ($PSCmdlet.ShouldProcess($ObjectName, 'Create Object')) {

				#send request to web service
				$result = Invoke-OIMRestMethod -Uri $URI -Method POST -Body $Body -WebSession $Script:WebSession -ContentType 'application/json'

				If ($null -ne $result) {

					Get-OIMObject -Objectname $ObjectName -id $result.uid

				}

			}
		}






	}#process

	END { }#end

}