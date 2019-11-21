

$OIM_EnvironmentVaribles = @{
    "DEV" = @{
        AppServer     = "SBX-IAM-9001.sandbox.local"
        AppServerName = "D1IMAppServer"
        DBServer      = "SBX-IAMDB-9001.sandbox.local"
        DBName        = "D2IMv7"
    }
    
    "ACC" = @{
        AppServer     = "ABX-IAMDB-8001.accbox.local"
        AppServerName = "D1IMAppServer"
        DBServer      = "ABX-IAM-8001.accbox.local"
        DBName        = "D2IMv7"
    }

    
    "PRD" = @{
        AppServer     = "HKT-IAM-0001.jumbo.local"
        AppServerName = "D1IMAppServer"
        DBServer      = "HKT-IAMDB-0001.jumbo.local"
        DBName        = "D2IMv7"
    }



}


