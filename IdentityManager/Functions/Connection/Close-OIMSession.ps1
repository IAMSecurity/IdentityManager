function Close-OIMSession {
	[CmdletBinding()]
	param()
	BEGIN {
		$URI = "$Script:BaseURI/auth/logout"

	}#begin

	PROCESS {

		#Send Logoff Request
		Invoke-OIMRestMethod -Uri $URI -Method POST -WebSession $Script:WebSession | Out-Null

	}#process

	END {

		#Set ExternalVersion to 0.0
		[System.Version]$Version = "0.0"
		Set-Variable -Name ExternalVersion -Value $Version -Scope Script -ErrorAction SilentlyContinue

		#Clear Module scope variables on logoff
		Clear-Variable -Name BaseURI -Scope Script -ErrorAction SilentlyContinue
		Clear-Variable -Name WebSession -Scope Script -ErrorAction SilentlyContinue

	}#end
}