<# 
    Name:         Run_MigrateUser_Prodv2.ps1
    Author:       Rob Looman
    Date:         15-11-2019
    Version:      1.0
    
    Description:  Script to migrate users from Quest to One Identity 

    PreReq: 
        - OIM_Module file
        - One Idenity System User
        - AD user 
        
    $csvPersonMigrate  = file with users to migrate
    $csvPersonInOrg    = Export of Table PersonInOrg from Quest
    $OIMServer         = One Identity manager server   
#>


# Initialize 
    # Login with system user 
        if($credAdmin  -eq $null){
            $credAdmin = Get-Credential -Message "Admin user" -UserName viAdmin 
        }

        . C:\IAM\ScriptRepository\Scripts\OneIM\Migratie\OIM_Module.ps1
    #Configuration parameters
        $csvPersonMigrate = "C:\IAM\ScriptRepository\Scripts\OneIM\Migratie\Input.csv"
        $csvPersonInOrg = "F:\applic\OneIM\Quest\Import\PersonInOrg.csv"
        $OIMServer = "svpv26236.verz.local"

    #OneIM Prod UID's
        $Global:UID_QERResourceType = "329b72d1-11cd-4d36-b8a2-7b7fda34216b"
        $Global:ACCProductGroupSPR = "d0f45e28-9339-46bb-9cd4-eb79bc1b044d" 
        $Global:ACCProductGroupTR = "678f7e00-1ce6-47e2-aa62-0896a589bcf9"  
        $Global:ACCProductGroupLR =  "72b7672e-df90-45eb-b342-fe9d5aed4b0b"      
        $Global:ACCProductGroupOld = "20c03a6f-30a4-4e22-af86-894ff5a3c631"
        $Global:ACCProductGroupRR = "30019b88-f8c7-4e62-8b05-ed7ad0ad7385"
        $Global:ITShopOrgTR =  "8e24a769-a6b5-4ffd-99f1-08f5dca566e1"
        $Global:ITShopOrgLR =  "cc663791-a49f-497b-9807-0310dcc40743"
        $Global:ITShopOrgOld = "975ede13-b980-44f9-81f3-3bb79e9ec4b6"
        $Global:ITShopOrgRR  ="ff88c2d2-3400-473e-be5e-17c4369e0d96"
        $Global:ITShopOrgSPR = "2d11bb9c-7641-4bfc-a1e5-69256d741dc5"
    
    #Connect
        $conAdmin = Connect-OIM -AppServer $OIMServer -useSSL -Credential $credAdmin  
        $con =      Connect-OIM -AppServer $OIMServer -useSSL

    # Load files 
        $MigrateUsers = Get-Content $csvPersonMigrate 
        $csv = Import-Csv $csvPersonInOrg -UseCulture

# Function definition 
    Function Add-O1IMServiceItem($ident_org,$ident_orgRoot) {        
            $ITShopOrg = ""
        # Business Role
            $OneIM_OrgFull  = Get-OIMObject -ObjectName Org -Where "ident_org  = '$($ident_org)'"   -First -Full

        # Assignment Resource
            $OneIM_QerASsign  = Get-OIMObject -ObjectName QERAssign -Where "ObjectKeyAssignTarget  = '$($OneIM_OrgFull.xObjectKey)'"  -First -Full

            if(-not [string]::IsNullOrEmpty($OneIM_QerASsign.uri) ){
                Write-Verbose ([string]::Format("{0}: Already Exists ",$ident_org) )
                return
            }
        
        # Determain ITShopOrg and ACCproduct depending on type
            switch ($ident_orgRoot)
            {
                 "Specialization Role" {
                    Write-Warning "Specialization Role"
                    $ACCProductGroup =  $ACCProductGroupSPR
                    $ITShopOrg = $ITShopOrgSPR 
                    
                }
                "TeamRole" {
                    $ACCProduct = $ident_org
                    $ACCProductGroup =  $ACCProductGroupTR
                    $ITShopOrg =  $ITShopOrgTR 
                }
                "License Role" {                    
                    $ACCProduct = $ident_org
                    $ACCProduct = $ACCProduct.Substring(13,$ACCProduct.Length-13)               
                    $ACCProductGroup = $ACCProductGroupLR 
                    $ITShopOrg =  $ITShopOrgLR 
                }
                "functionalRole"{
                    $ACCProduct = $ident_org
                    $ACCProductGroup = $ACCProductGroupOld 
                    $ITShopOrg = $ITShopOrgOld 
                }
                "Retail Role"{
                    $ACCProduct = $ident_org
                    $ACCProduct = $ACCProduct.Substring(11,$ACCProduct.Length-11)     
                    $ACCProductGroup =  $ACCProductGroupRR 
                    $ITShopOrg =  $ITShopOrgRR  
                }
                "projectRole"{
                    $ACCProduct = $ident_org
                    $ACCProductGroup =  $ACCProductGroupOld 
                    $ITShopOrg = $ITShopOrg
                }
                "afdelingsbasisrol"{
                    $ACCProduct = $ident_org
                    $ACCProductGroup =  $ACCProductGroupOld 
                    $ITShopOrg = $ITShopOrg
                }
                "SNS Approval Workflow"{
                    Write-Verbose "Invalid Org Root SNS Approval Workflow"
                    return 
                }else{
        
                    Write-Warning "Invalid Org Root ($($ident_orgRoot))"
                    return 
                }
        
            }

            if($ident_orgRoot -eq "Specialization Role" ){continue}

            $OneIM_Ident_Org = $OneIM_OrgFull.Ident_Org
            $OneIM_Ident_xObjectKey = $OneIM_OrgFull.xObjectKey
            $Description = $OneIM_OrgFull.FullPath
 

        # Create ACCProduct
                Write-Host ([string]::Format("{0}: Create ACCProcut",$ident_org) )   
            $ArticleCode = $OneIM_OrgFull.Ident_Org
            if( $ArticleCode.Length  -gt 64){$ArticleCode =  $ArticleCode.Substring(0,63)}
            
            $values = @{ 
                ArticleCode =$ArticleCode;
                Description= "$Description";
                Ident_ACCProduct=$ACCProduct;
                UID_ACCProductGroup=$ACCProductGroup}    

            $newACCProductFull = New-OIMObject -ObjectName ACCProduct -Properties $values

    
        # Create QERAssign
            $Ident_QERAssign = $OneIM_Ident_Org
            if( $Ident_QERAssign.Length  -gt 64){$Ident_QERAssign  = $Ident_QERAssign.Substring(0,63)}    
                Write-Host ([string]::Format("{0}: Create QERAssign",$ident_org) )

            $values = @{ 
                Ident_QERAssign = $Ident_QERAssign
                isForITSHop= "True";
                isITShopOnly= "True";
                ObjectKeyAssignTarget=$OneIM_Ident_xObjectKey;
                ACCProductGroup=$ACCProductGroup
                UID_AccPRoduct = $newACCProductFull.UID_AccProduct
                UID_DialogTableAssignTarget = "RMB-T-Org"
                UID_QERResourceType = $UID_QERResourceType 
                }     
            $newQERAssignFull = New-OIMObject -ObjectName QERAssign -Properties $values


        
        # Create ITShopOrgHasQERAssign          
            Try{   
                     Write-Host ([string]::Format("{0}: Create ITShopOrgHASQERAssign",$ident_org) )
                $values=@{ 
                    UID_QERAssign =$newQERAssignFull.UID_QERAssign
                    UID_ITShopOrg= $ITShopOrg }          
                $newITShopOrgHASQERAssign = New-OIMObject -ObjectName ITShopOrgHASQERAssign -Properties $values
            }Catch{
                Write-Warning ([string]::Format("{0}: Create ITShopOrgHASQERAssign ({1})",$ident_org,$newQERAssignFull.Ident_QERAssign) )

            }
} #End Function Add-O1IMServiceItem

    Function Add-OIMPersonWantsOrg($CentralAccount, $Ident_org  ) {
        # Retrieving Person 
            $OneIM_UserFull = Get-OIMObject -ObjectName Person -Where "Personnelnumber  = '$CentralAccount'"   -First -Full
            if([string]::IsNullOrEmpty($OneIM_UserFull.uri)){
                Write-Warning ([string]::Format("{0}: Could not user",$CentralAccount,$Ident_org) )
                return
            }
            $UID_Person     = $OneIM_UserFull.UID_Person
            $UID_PersonHead = $OneIM_UserFull.UID_PersonHead
            if($UID_PersonHead.Length  -lt 10 ){$UID_PersonHead    = $UID_Person   }

        # Retrieving Business role
            $OneIM_OrgFull = Get-OIMObject -ObjectName Org -Where "ident_org  = '$($Ident_org)'"   -First -Full
            if([string]::IsNullOrEmpty($OneIM_OrgFull.uri) ){
                Write-Warning ([string]::Format("{0}: Could not find org '{1}'",$CentralAccount,$Ident_org) )
                $errorOrg = $Ident_org
                return
            }
    
        # Retrieving Assignment Resource              
            $OneIM_QerASsignFull = Get-OIMObject -ObjectName QERAssign -Where "ObjectKeyAssignTarget  = '$($OneIM_OrgFull.xObjectKey)'"   -First -Full
            if([string]::IsNullOrEmpty($OneIM_QerASsignFull.uri) ){
                Write-Warning ([string]::Format("{0}: Could not find ITShop Assignment '{1}'",$CentralAccount,$Ident_org) )
                return
            }
            $OneIM_UID_AccProduct = $OneIM_QerASsignFull.UID_AccProduct

        # Retrieving IT Shop ORg 
            $OneIM_ITShopOrg = Get-OIMObject -ObjectName ITShopOrg -Where "UID_AccProduct  = '$OneIM_UID_AccProduct'"   -First -Full
            if([string]::IsNullOrEmpty($OneIM_ITShopOrg.uri) ){
                Write-Warning ([string]::Format("{0}: Could not find ITShop Org '{1}'",$CentralAccount,$Ident_org) )
                return
            }
        # Person Wants Org 
            $requestITShopOrg = $OneIM_ITShopOrg.UID_ITShopOrg
            $requestPersonInserted =   $UID_PersonHead   
            $requestPersonOrder =   $UID_Person
            $requestObjectKeyOrderd = $OneIM_QerASsignFull.Xobjectkey
            $requestObjectKeyAssignment = "<Key><T>PersonInOrg</T><P>$($OneIM_OrgFull.UID_Org)</P><P>$requestPersonOrder</P></Key>"


        # Check already exists 
            $where="UID_PersonInserted  = '$requestPersonInserted' AND UID_PersonOrdered = '$requestPersonOrder' AND  ObjectKeyOrdered  = '$requestObjectKeyOrderd ' AND OrderState = 'Assigned'"
            $OneIM_ITShopOrg2 = Get-OIMObject -ObjectName PersonWantsorg -Where $where   -First -Full

            if(-not [string]::IsNullOrEmpty($OneIM_ITShopOrg2.uri) ){
                Write-Warning ([string]::Format("{0}: Request already done '{1}'",$CentralAccount,$Ident_org) )
                return
            }else{        

                Write-Host ([string]::Format("{0}: Create Request for org '{1}'",$CentralAccount,$Ident_org) )
                $values=@{ 
                        OrderREason ="Migratie";
                        UID_ORG= "$requestITShopOrg";
                        UID_PersonOrdered="$requestPersonOrder";
                        UID_PersonInserted="$requestPersonInserted";
                        ObjectKeyOrdered=$requestObjectKeyOrderd;
                        ObjectKeyAssignment=$requestObjectKeyAssignment}
                $newPersonWantsorg = New-OIMObject -ObjectName PersonWantsOrg -Properties $values

            }
    } #End Function Add-OIMPersonWantsOrg

# Main
    ForEach($user in $MigrateUsers ){

        #0. Retrieve OneIM user 
            $UserID = $user
        
        Write-Host "Processing user: $user "
            $OneIM_UserFull   =   Get-OIMObject -ObjectName Person     -Where "Personnelnumber = '$UserID'"   -First -Full    
            $OneIM_UserADFull =   Get-OIMObject -ObjectName ADSACcount -Where "SamACcountName  = '$UserID'"   -First -Full   

            if([string]::IsNullOrEmpty($OneIM_UserFull.uri)){
                Write-Warning ([string]::Format("{0}: Could not find user",$UserID) )
                return
            }

        #1. Retrieve mail
            $mail =  $OneIM_UserADFull.Mail


        #2. Update AD User:  HomeShare, Link unmanaged user
            # HomeShare = Home$\4
            # HomeDirPath = 9321871
            # SharedAs = 9321871
            # HomeDirectory = \\verz.local\home$\2\0011041
            If($OneIM_UserADFull.HomeDirectory -match "\\\\verz.local\\(.*)\\$($UserID)"){
                $HomeShare = $Matches[1]
                $values=@{ 
                    HomeShare = $HomeShare  
                    UID_Person = $OneIM_UserFull.UID_Person
                    UID_TSBAccountDef  = "78f72c8a-ccf8-47ca-9339-c9dbc566d178" #Standaard VERZ account
                    UID_TSBBehavior = "TSB-UnManaged"    
                }      
                $OneIM_UserADUpdate = Update-OIMObject -Object $OneIM_UserADFull  -Properties $values -Session $conAdmin
            }
        #3. Update person with DefaultEmailAddress (OneIM)

            $values=@{  DefaultEmailAddress =   $mail      }     
            $OneIM_UserFullUpdate = Update-OIMObject -Object $OneIM_UserFull  -Properties $values -Session $conAdmin
            
        #4. Add Person to the migratie business role (dynamic Role)
    
            $OneIM_OrgMigrate =  Get-OIMObject -ObjectName Org        -Where "Ident_Org  = 'BAR_000001__Migratie'"   -First -Full     
            Iets
            $PersonInOrg =  Get-OIMObject -ObjectName PersonInOrg     -Where "UID_Org  = '$($OneIM_OrgMigrate.UID_Org)' AND  UID_Person  = '$($OneIM_UserFull.UID_Person)' "   -First -Full      
            
            if( [string]::IsNullOrEmpty($PersonInOrg.uri )){
            
                $values=@{ 
                    UID_Org = $OneIM_OrgMigrate.UID_Org
                    UID_Person= $OneIM_UserFull.UID_Person;
            
                    }     
                $newPersonInOrg =  New-OIMObject -ObjectName PersonInOrg -Properties $values -Session $conAdmin
            }
        
       
       #5. UNSAccount
            $OneIM_UserUNSACcounts = Get-OIMObject -ObjectName UNSACcount -Where "accountname  like '%$UserID'"  -Full   

            ForEach( $OneIM_UserUNSACcountFull in  $OneIM_UserUNSACcounts ){
                
                #VERZ ACcount 
                If($OneIM_UserUNSACcountFull.UID_TSBBehavior -eq "" -and $OneIM_UserUNSACcountFull.CanonicalName -like  "verz.local/Organisatie/Users*"){
                    #$OneIM_UserUNSACcountFull
                    # $OneIM_UserUNSACcountFull.UID_TSBBehavior -ne "TSB-FullManaged"
                
                    <#
                        $OneIM_UserUNSACcountFull.CanonicalName
                        $OneIM_UserUNSACcountFull.UID_DPRNameSpace
                        $OneIM_UserUNSACcountFull.UID_TSBBehavior
                    #>
                }else{
                    #other UNSAccount

                }
            

            }
       

        
        

        #6. Add Person to PersonWantsOrg
                ForEach($record in $csv){

                    if($record.CentralAccount -eq $UserID){
                        $centralaccount = $record.CentralAccount 
                        $org = $record.Ident_org 
                        $orgroot = $record.ident_OrgRoot
                        Write-Host "Add user $centralaccount to org: $org"
                        Add-O1IMServiceItem -ident_org $org -ident_orgRoot $orgroot
                        Add-OIMPersonWantsOrg -CentralAccount $centralaccount -Ident_org $org
                    }
                }
        }# End ForEach MigrateUsers 
