Function Wait-OIMJobQueue{
    [CmdletBinding()]
    param (
        [Parameter()]
        $id,
        $JobChainName,
        $timeout = 300,
        $sleep = 3
    )
    $jobs = ""
    $remainingtime = $timeout
    if($null -eq $id){
        $where = "JobChainName LIKE '%" + $JobChainName +"%' AND Ready2EXE <> 'HISTORY' AND Ready2exe <> 'FINISHED'"

    }else{
        $where = "UID_JobQueue ='$id'"
    }

    While($null -ne $jobs -and $remainingtime -ge 0){
        Write-Progress "Wait jobqueue with jobchainname:'$JobChainName' jobscount:$( $jobs.count) seconds remaining:$remainingtime" -PercentComplete ((($timeout-$remainingtime)/$timeout )*100)

        Start-Sleep -Seconds $sleep
        $remainingtime -=  $sleep

        $jobs =  Get-OIMObject -ObjectName JobQueue -Where  $where

    }
}