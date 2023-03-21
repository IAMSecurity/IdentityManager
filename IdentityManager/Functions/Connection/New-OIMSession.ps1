function New-OIMSession {
	[CmdletBinding(SupportsShouldProcess)]
	param(
		[PSCredential]$Credential,

		[ValidateSet('DialogUser', 'RoleBasedADSAccount')]
		[string]$Module = 'RoleBasedADSAccount',
		[parameter(
			Mandatory = $true,
			ValueFromPipeline = $false,
			ValueFromPipelinebyPropertyName = $true
		)]
		[string]$BaseURI,

		[parameter(
			Mandatory = $false,
			ValueFromPipeline = $false,
			ValueFromPipelinebyPropertyName = $true
		)]
		[string]$AppName = 'AppServer',

		[Parameter(
			Mandatory = $false,
			ValueFromPipeline = $false,
			ValueFromPipelinebyPropertyName = $false
		)]
		[pscredential]$IISCredential,


		[Parameter(
			Mandatory = $false,
			ValueFromPipeline = $false,
			ValueFromPipelinebyPropertyName = $false
		)]
		[switch]$SkipVersionCheck,

		[parameter(
			Mandatory = $false,
			ValueFromPipeline = $false,
			ValueFromPipelinebyPropertyName = $true
		)]
		[switch]$SkipCertificateCheck

	)

	BEGIN {

		$Uri = "$baseURI/$AppName"
		#Hashtable to hold Logon Request
		$LogonRequest = @{ }

		#Define Logon Request Parameters
		$LogonRequest['Method'] = 'POST'
		$LogonRequest['SessionVariable'] = 'WebSession'
		if($null -eq $IISCredential){

			$LogonRequest['UseDefaultCredentials'] = $true
		}
		$LogonRequest['SkipCertificateCheck'] = $SkipCertificateCheck.IsPresent






	}#begin


	# Connecting


	PROCESS {

		$authdata = @{AuthString = "Module=$Module" }
		if ($null -ne $Credential ) {
			$authdata = @{AuthString = "Module=$Module;User=$($Credential.Username);Password=$($Credential.GetNetworkCredential().password)" }
		}
		$authJson = ConvertTo-Json $authdata -Depth 2

		$LogonRequest['Uri'] = "$Uri/auth/apphost"  #hardcode Windows for integrated auth
		$LogonRequest['Body'] = $authJson.ToString()
		
		if ($null -ne $IISCredential ) {
			$LogonRequest['Credential'] = $IISCredential
			Write-Warning "Using IIS credentials"
		}
		if ($PSCmdlet.ShouldProcess($LogonRequest['Uri'], 'Logon')) {

			try {
				#Send Logon Request
				$OIMSession = Invoke-OIMRestMethod @LogonRequest
			}
			catch {

				#Throw all errors not related to ITATS542I
				throw $PSItem


			}
			finally {

				#If Logon Result
				If ($OIMSession) {

					#BaseURI set in Module Scope
					Set-Variable -Name BaseURI -Value $Uri -Scope Script
					Set-Variable -Name WebSession -Value $WebSession -Scope Script
					Set-Variable -Name IISCredential -Value $IISCredential -Scope Script

					#Initial Value for Version variable
					[System.Version]$Version = '0.0'

					if ( -not ($SkipVersionCheck)) {

						Try {

							#Get CyberArk ExternalVersion number.
							[System.Version]$Version = Get-OIMObject DialogDatabase -ErrorAction Stop |
							Select-Object -ExpandProperty EditionVersion

						}
						Catch { [System.Version]$Version = '0.0' }

					}

					#Version information available in module scope.
					Set-Variable -Name ExternalVersion -Value $Version -Scope Script

				}

			}

		}

	}#process

	END { }#end

}