$Prod =$true 
$HostnamePROD = "https://svpv26236.verz.local"
$HostnameACC = "https://SVAV26226.verz-iam.local"

if($prod){
    $Hostname = $HostnamePROD
    $authdata = @{AuthString="Module=RoleBasedADSAccount"}
    $authdata = @{AuthString="Module=DialogUser;User=viAdmin;Password=359EBUY7d3w!!"}
    $authJson = ConvertTo-Json $authdata -Depth 2
    #Login
    if($wsessionPRod -eq $null){
        Invoke-RestMethod -Uri "$Hostname/AppServer/auth/apphost" -Body $authJson.ToString() -Method Post -UseDefaultCredentials -Headers @{Accept="application/json"} -SessionVariable wsessionPRod
    }
    $wsession = $wsessionPRod

 }else  {
    $Hostname = $HostnameACC
    $authdata = @{AuthString="Module=RoleBasedADSAccount"}
    $authdata = @{AuthString="Module=DialogUser;User=dvbAdmin;Password=WIG^EXgM5GOfbI"}
    $authJson = ConvertTo-Json $authdata -Depth 2
    #Login
    if($wsessionACC -eq $null){
        Invoke-RestMethod -Uri "$Hostname/AppServer/auth/apphost" -Body $authJson.ToString() -Method Post -UseDefaultCredentials -Headers @{Accept="application/json"} -SessionVariable wsessionACC
    }
     $wsession = $wsessionACC
}

If(-not $load){ 

# Retrieve migrate users
    $csv = Import-Csv F:\applic\OneIM\Quest\Import\Org_Oud.csv -UseCulture
    $csv = $csv | Select-Object -Unique Ident_Org,ident_OrgRoot     
#Credential
   

    $csvAC = Import-Csv C:\IAM\ScriptRepository\Scripts\OneIM\Migratie\AC.csv -UseCulture
    $dicAC = $csvAC | Group-Object -AsHashTable -Property Number
    $load = $true
}



if($Prod){# Prod 
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
}Else{
# ACC 

    $Global:UID_QERResourceType = "f3fc5b23-fafa-4e3c-831d-fbe97f1e75df"
       
    $Global:ACCProductGroupSPR = "e8fef9fd-63f1-41cd-b5be-e4e37ed92041"
    $Global:ACCProductGroupTR = "2bd3d893-0290-4915-9020-e62bd48c0d67" 
    $Global:ACCProductGroupLR = "9bd3eb2b-c091-4734-a721-517bfe28c78c"       
    $Global:ACCProductGroupOld ="5f3c1e45-f288-43cc-bbb1-569783927a1d" 
    $Global:ACCProductGroupRR = "d2014904-c684-4ba8-86b6-ebd163fabc7f" 
    $Global:ITShopOrgTR =  "78905e70-f6f5-4436-adb3-4d50156b4ce4" 
    $Global:ITShopOrgLR =  "8fae10cd-1e5d-4dfd-b6d0-b11645ea5c1c" 
    $Global:ITShopOrgOld = "9eb6ffc4-89c9-4b13-a48d-ad6e101829b5" 
    $Global:ITShopOrgRR  ="bee37442-c79d-48eb-bd10-839a2a9d8ffb" 
    $Global:ITShopOrgSPR ="a2483702-c307-4ca3-b29f-2708abce53c7"  
}

ForEach($record in $csv){
    # Retrieving org in OneIM 
        $body = @{where="ident_org  = '$($record.Ident_Org)'"} | ConvertTo-Json
        $OneIM_Org = Invoke-RestMethod -Uri "$Hostname/AppServer/api/entities/Org?DisplayColumn=xObjectKey" -WebSession $wsession -Method Post -Body $body -ContentType application/json 
        if([string]::IsNullOrEmpty($OneIM_Org.uri) ){
            Write-Warning ([string]::Format("{0}: Could not find org",$record.Ident_Org) )
            continue
        }
        $OneIM_OrgFull = Invoke-RestMethod -Uri "$Hostname/$($OneIM_Org[0].uri)" -WebSession $wsession -Method Get -ContentType application/json
    
    # Retrieving QERAssign 
    
    
        $body = @{where="ObjectKeyAssignTarget  = '$($OneIM_OrgFull.values.xObjectKey)'"} | ConvertTo-Json
        $OneIM_QerASsign = Invoke-RestMethod -Uri "$Hostname/AppServer/api/entities/QERAssign" -WebSession $wsession -Method Post -Body $body -ContentType application/json
        if(-not [string]::IsNullOrEmpty($OneIM_QerASsign.uri) ){
            Write-Warning ([string]::Format("{0}: Already Exists ",$record.Ident_Org) )
            if($record.ident_OrgRoot -eq "Projectrole"){}else{continue}
        }
        

    # ACCProduct
    $ITShopOrg = ""
    switch ($record.ident_OrgRoot)
    {
        "Specialization Role" {
                   # $ACCProductGroup =  "d0f45e28-9339-46bb-9cd4-eb79bc1b044d"
                   Write-Warning "Specialization Role"
                  $ACCProductGroup =  $ACCProductGroupSPR
                    $ITShopOrg = $ITShopOrgSPR 
                    
        }
        "TeamRole" {
                   $ACCProduct = $record.Ident_org
                   $ACCProductGroup =  $ACCProductGroupTR
                    $ITShopOrg =  $ITShopOrgTR 
        }
        "License Role" {
                    
                   $ACCProduct = $record.Ident_org
                   $ACCProduct = $ACCProduct.Substring(13,$ACCProduct.Length-13)               
                   $ACCProductGroup = $ACCProductGroupLR 
                    $ITShopOrg =  $ITShopOrgLR 
        }
        "functionalRole"{
                   $ACCProduct = $record.Ident_org
                   $ACCProductGroup = $ACCProductGroupOld 
                    $ITShopOrg = $ITShopOrgOld 
        }
        "Retail Role"{
                    $ACCProduct = $record.Ident_org
                    $ACCProduct = $ACCProduct.Substring(11,$ACCProduct.Length-11)     
                    $ACCProductGroup =  $ACCProductGroupRR 
                    $ITShopOrg =  $ITShopOrgRR  
        }
        "projectRole"{
                   $ACCProduct = $record.Ident_org
                   $ACCProductGroup =  $ACCProductGroupOld 
                    $ITShopOrg = $ITShopOrgOld
        }
        "afdelingsbasisrol"{
                   $ACCProduct = $record.Ident_org
                   $ACCProductGroup =  $ACCProductGroupOld 
                    $ITShopOrg = $ITShopOrg
        }
        "SNS Approval Workflow"{
            Write-Warning "Invalid Org Root"
            continue 
        }else{
        
            Write-Warning "Invalid Org Root ($($record.ident_OrgRoot))"
            continue 
        }
        
    }

    if($record.ident_OrgRoot -eq "Specialization Role" ){continue}
    


         $OneIM_Ident_Org = $OneIM_OrgFull.values.Ident_Org
         $OneIM_Ident_xObjectKey = $OneIM_OrgFull.values.xObjectKey
      
 
        $Description = $OneIM_OrgFull.values.FullPath

            
            
 

        # Check already exists 
                

        Write-Host ([string]::Format("{0}: Create ACCProcut",$record.Ident_org) )
              
              
        $ArticleCode = $OneIM_OrgFull.values.Ident_Org
        if( $ArticleCode.Length  -gt 64){$ArticleCode =  $ArticleCode.Substring(0,63)}


        $body = @{values=@{ 
            ArticleCode =$ArticleCode;
            Description= "$Description";
            Ident_ACCProduct=$ACCProduct;
            UID_ACCProductGroup=$ACCProductGroup}    }      | ConvertTo-Json
       # $newACCProduct = (Invoke-RestMethod -Uri "$Hostname/AppServer/api/entity/ACCProduct" -WebSession $wsession -Method Post -Body $body -ContentType application/json).uri
        # $newACCProductFull = Invoke-RestMethod -Uri "$Hostname/$newACCProduct" -WebSession $wsession -Method Get -ContentType application/json  

    

    #QERAssign
        $Ident_QERAssign = $OneIM_Ident_Org
        if( $Ident_QERAssign.Length  -gt 64){$Ident_QERAssign  = $Ident_QERAssign.Substring(0,63)}
    
        Write-Host ([string]::Format("{0}: Create QERAssign",$record.Ident_org) )
        $body = @{values=@{ 
            Ident_QERAssign = $Ident_QERAssign
            isForITSHop= "True";
            isITShopOnly= "True";
            ObjectKeyAssignTarget=$OneIM_Ident_xObjectKey;
            ACCProductGroup=$ACCProductGroup
            UID_AccPRoduct = $newACCProductFull.values.UID_AccProduct
            UID_DialogTableAssignTarget = "RMB-T-Org"
            UID_QERResourceType = $UID_QERResourceType 
            }      }    | ConvertTo-Json


        $body = @{where="Ident_QERAssign  = '$Ident_QERAssign'"} | ConvertTo-Json
        $newQERAssign = Invoke-RestMethod -Uri "$Hostname/AppServer/api/entities/QERAssign" -WebSession $wsession -Method Post -Body $body -ContentType application/json
        if([string]::IsNullOrEmpty($newQERAssign.uri)){
        $newQERAssign = Invoke-RestMethod -Uri "$Hostname/AppServer/api/entity/QERAssign" -WebSession $wsession -Method Post -Body $body -ContentType application/json
        }

        $newQERAssignFull = Invoke-RestMethod -Uri "$Hostname/$($newQERAssign.uri)" -WebSession $wsession -Method Get -ContentType application/json  
        
    
        Write-Host ([string]::Format("{0}: Create ITShopOrgHASQERAssign",$record.Ident_org) )
        $body = @{values=@{ 
            UID_QERAssign =$newQERAssignFull.values.UID_QERAssign
            UID_ITShopOrg= $ITShopOrg;
            }      }    | ConvertTo-Json
        Try{
            $newITShopOrgHASQERAssign = (Invoke-RestMethod -Uri "$Hostname/AppServer/api/entity/ITShopOrgHASQERAssign" -WebSession $wsession -Method Post -Body $body -ContentType application/json).uri
        }Catch{
            Write-Warning ([string]::Format("{0}: Create ITShopOrgHASQERAssign ({1})",$record.Ident_org,$newQERAssignFull.values.Ident_QERAssign) )

        }

          
       
   

}
<#
Invoke-RestMethod -Uri "$Hostname/AppServer/auth/logout" -WebSession $wsession -Method Post



#>

