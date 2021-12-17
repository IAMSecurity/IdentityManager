function Set-OIMGlobalVariable {
	[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Gen2')]
	param(

		[parameter(
			Mandatory = $true,
			ValueFromPipelinebyPropertyName = $true
		)]
		[string]$Name,
		[parameter(
			Mandatory = $true,
			ValueFromPipelinebyPropertyName = $true
		)]
		[string]$Value
	)

	BEGIN {
		#Assert-VersionRequirement -RequiredVersion 8.1
	}#begin

	PROCESS {

		$URI = "$Script:BaseURI/appserver/variable/$Name"
		$body =  @{"Value"= $Value}
		#Get request parameters
		if ($PSCmdlet.ShouldProcess("uri", "Set-OIMGlobalVariable $Name")) {
			Invoke-OIMRestMethod -Uri  $uri  -Method PUT -Body $body
		}



	}#process

	END { }#end

}