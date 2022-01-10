function Get-OIMObject {
	[CmdletBinding(SupportsPaging = $true, DefaultParameterSetName = "ObjectSearch")]
	param(
		[Parameter(
			ParameterSetName = "Object",
			Mandatory = $true,
			ValueFromPipeline = $true)]
		$Object,
		[Parameter(Position = 0,ParameterSetName = "ObjectSingle", Mandatory = $true)]
		[Parameter(Position = 0,ParameterSetName = "ObjectSearch", Mandatory = $true)]
		[Alias("Type")]
		[string]
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
				$URI = "$Script:BaseURI/api/entities/$ObjectName`?LoadType=$LoadType"
				$queryString = $PSBoundParameters | Get-OIMParameter -ParametersToKeep  LoadType, limit, offset | ConvertTo-QueryString
				Write-Verbose ($PSBoundParameters.Keys -Join "-")
				if ($null -ne $queryString) {
					#Build URL from base URL
					$URI = "$URI&$queryString"

				}

				$body = $PSBoundParameters | Get-OIMParameter -ParametersToKeep  where, displayColumns | ConvertTo-Json


				if($null -eq $body){
					$result = Invoke-OIMRestMethod -Uri $URI -WebSession $Script:WebSession
				}else{
					$result = Invoke-OIMRestMethod -Uri $URI -Method POST -Body $body -ContentType 'application/json' -WebSession $Script:WebSession

				}
			}

		}



		#send request to web service


		If ($null -ne $result) {
			#$result | Add-ObjectDetail -typename $TypeName
			$result.Values | Select-Object @SelectPAram
		}

	}#process

	END { }#end

}