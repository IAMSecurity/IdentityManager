Function Wait-OIMJobQueue{
    [CmdletBinding()]
    param (
        [Parameter()]
        $JobChainName,
        $timeout = 300,
        $sleep = 3
    )
    $jobs = ""
    $remainingtime = $timeout

    $where = "JobChainName LIKE '%" + $JobChainName +"%' AND Ready2EXE <> 'HISTORY' AND Ready2exe <> 'FINISHED'"

    While($null -ne $jobs -and $remainingtime -ge 0){
        Write-Progress "Wait jobqueue with jobchainname:'$JobChainName' jobscount:$( $jobs.count) seconds remaining:$remainingtime" -PercentComplete ((($timeout-$remainingtime)/$timeout )*100)

        Start-Sleep -Seconds $sleep
        $remainingtime -=  $sleep

        $jobs =  Get-OIMObject -ObjectName JobQueue -Where  $where

    }
}