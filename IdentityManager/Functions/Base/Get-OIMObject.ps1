function Get-OIMObject {
	[CmdletBinding(SupportsPaging = $true, DefaultParameterSetName = "ObjectSearch")]
	param(
		[Parameter(Position = 0,
			ParameterSetName = "Object",
			Mandatory = $true,
			ValueFromPipeline = $true)]
		$Object,
		[Parameter(ParameterSetName = "ObjectSingle", Mandatory = $true)]
		[Parameter(ParameterSetName = "ObjectSearch", Mandatory = $true)]
		[Alias("Type")]
		$ObjectName,
		[Parameter(ParameterSetName = "ObjectSingle", Mandatory = $true)]
		[Alias("uid")]
		$id,
		[Parameter(ParameterSetName = "ObjectSearch")]
		[ValidateSet("Default", "BulkReadOnly", "Slim", "ForeignDisplays", "ForeignDisplaysForAllColumns")]
		$LoadType = "BulkReadOnly",
		[Parameter(ParameterSetName = "ObjectSearch")]
		$Where,
		[Parameter(ParameterSetName = "ObjectSearch")]
		$OrderBy,
		[Parameter(ParameterSetName = "ObjectSearch")]
		$displayColumns

	)

	BEGIN {
		$limit = $PSCmdlet.PagingParameters.First
		$SelectPAram = $null
		if ( $PSBoundParameters.Keys.contains("First")) {	$SelectPAram = @{First = $limit } }

	}#begin

	PROCESS {
		#Assert-VersionRequirement -RequiredVersion 8.1
		switch ($PSCmdlet.ParameterSetName) {

			'Object' {
				ForEach ($item in $object) {

					$xmlXObjectKey = [xml] $item.xObjectKey
					GEt-OIMObject -ObjectName $xmlXObjectKey.key.T -id $xmlXObjectKey.key.P
				}
			}

			'ObjectSingle' {
				$URI = "$Script:BaseURI/api/entity/$ObjectName/$id"
				$result = Invoke-OIMRestMethod -Uri $URI -Method GET -WebSession $Script:WebSession

			}


			'ObjectSearch' {
				$URI = "$Script:BaseURI/api/entity/$ObjectName"
				$queryString = $PSBoundParameters | Get-OIMParameter -ParametersToKeep  loadType, limit, offset | ConvertTo-QueryString
				$body = $PSBoundParameters | Get-OIMParameter -ParametersToKeep  where, displayColumns | ConvertTo-QueryString

				if ($null -ne $queryString) {
					#Build URL from base URL
					$URI = "$URI`?$queryString"

				}
				$result = Invoke-OIMRestMethod -Uri $URI -Method POST -Body $body -WebSession $Script:WebSession
			}

		}



		#send request to web service


		If ($null -ne $result) {

			#$result | Add-ObjectDetail -typename $TypeName
			$result | Select-Object @SelectPAram
		}

	}#process

	END { }#end

}