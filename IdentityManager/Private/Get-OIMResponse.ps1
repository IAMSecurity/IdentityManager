function Get-OIMResponse {
	<#
	.SYNOPSIS
	Receives and returns the content of the web response from the CyberArk API

	.DESCRIPTION
	Accepts a WebResponseObject.
	By default returns the Content property OIMsed in the output of Invoke-OIMRestMethod.
	Processes the API response as required depending on the format of the response, and
	the format required by the functions which initiated the request.

	.PARAMETER APIResponse
	A WebResponseObject, as returned from the OIM API using Invoke-WebRequest

	.EXAMPLE
	$WebResponseObject | Get-OIMResponse

	Parses, if required, and returns, the required properties of $WebResponseObject

	#>
	[CmdletBinding()]
	[OutputType('System.Object')]
	param(
		[parameter(
			Position = 0,
			Mandatory = $true,
			ValueFromPipeline = $true)]
		[ValidateNotNullOrEmpty()]
		[Microsoft.PowerShell.Commands.WebResponseObject]$APIResponse

	)

	BEGIN {	}#begin

	PROCESS {

		if ($APIResponse.Content) {

			#Default Response - Return Content
			$OIMResponse = $APIResponse.Content

			#get response content type
			$ContentType = $APIResponse.Headers["Content-Type"]

			#handle content type
			switch ($ContentType) {

				'text/html; charset=utf-8' {

					If ($OIMResponse -match '<HTML>') {

						#Fail if HTML received from API

						$PSCmdlet.ThrowTerminatingError(

							[System.Management.Automation.ErrorRecord]::new(

								"Guru Meditation - HTML Response Received",
								$StatusCode,
								[System.Management.Automation.ErrorCategory]::NotSpecified,
								$APIResponse

							)

						)

					}

				}

				'application/json; charset=utf-8' {

					#application/json content expected for most responses.

					#Create Return Object from Returned JSON
					$OIMResponse = ConvertFrom-Json -InputObject $APIResponse.Content

				}

				default {

					# Byte Array expected for files to be saved
					if ($($OIMResponse | Get-Member | Select-Object -ExpandProperty typename) -eq "System.Byte" ) {

						#return content and headers
						$OIMResponse = $APIResponse | Select-Object Content, Headers

						#! to be OIMsed to `Out-OIMFile`

					}

				}

			}

			#Return OIMResponse
			$OIMResponse

		}

	}#process

	END {	}#end

}