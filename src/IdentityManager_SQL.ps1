
Function Connect-OIMSQL($servername, $database, [PSCredential] $Cred ) {
    
   
    $ConnectionSTring =  "server='$servername';database='$database';"

    # Creating connection string
    if ($null -eq $Cred ) {
        #Single sign
        $ConnectionSTring  = $ConnectionSTring + "integrated security=true"
    }else {

        $user = $Cred.Username
        $Pass = $Cred.GetNetworkCredential().password
        
        $ConnectionSTring  = $ConnectionSTring + "trusted_connection=false; user= '$user'; Password = '$Pass'; integrated security='False'"
    }

    $Global:OIM_SQLConnection = New-Object System.Data.SQLClient.SQLConnection
    $Global:OIM_SQLConnection.ConnectionString = $ConnectionString 
    $Global:OIM_SQLConnection.Open()

    
    $Global:OIM_SQLConnection 
}


# function that executes sql commands against an existing Connection object; In pur case
# the connection object is saved by the ConnectToDB function as a global variable
function Invoke-OIMSQLQuery {
    # define parameters
    param(
     
        [string]
        $sqlquery,
        $connection = $Global:OIM_SQLConnection
    
    )
    
    Begin {
        If (!$Connection) {
            Throw "No connection to the database detected. Run command Connect-OIMSQL first."
        }
        elseif ($Connection.State -eq 'Closed') {
            Write-Verbose 'Connection to the database is closed. Re-opening connection...'
            try {
                # if connection was closed (by an error in the previous script) then try reopen it for this query
                $Connection.Open()
            }
            catch {
                Write-Verbose "Error re-opening connection. Removing connection variable."
                Remove-Variable -Scope Global -Name Connection
                throw "Unable to re-open connection to the database. Please reconnect using the ConnectToDB commandlet. Error is $($_.exception)."
            }
        }
    }
    
    Process {
        #$Command = New-Object System.Data.SQLClient.SQLCommand
        $command = $Connection.CreateCommand()
        $command.CommandText = $sqlquery
    
        Write-Verbose "Running SQL query '$sqlquery'"
        try {
            $result = $command.ExecuteReader()      
        }
        catch {
            $Connection.Close()
        }
        $Datatable = New-Object "System.Data.Datatable"
        $Datatable.Load($result)
        return $Datatable          
    }
    End {
        Write-Verbose "Finished running SQL query."
    }
}

<#
$cred = Get-Credential 
Connect-OIMSQL -
Invoke-OIMSQLQuery -sqlquery "select top 1 * from person"

#>