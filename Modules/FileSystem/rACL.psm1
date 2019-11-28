
Function Get-rAclCustom {
    Param($foldername,
    [switch]$Inheritance)
    Process{
        $list = @()
        $acl = Get-ACL $foldername
    
        $properties = @{
                Folder=$foldername;
                DirectLid="";
                DirectLidType="";
                IndirectLid="";
                IndirectLidINS=""
                
            }

    #$list += New-Object -TypeName PSObject -Property  $properties 

        ForEach($access in $acl.Access){



            if(($access.IdentityReference  -like "*\*")-AND ((!$access.IsInherited)-or $Inheritance)){
            


                $splitMember = $access.IdentityReference -split "\\"            
                $ObjectName = $splitMember[1]
                $ADObject = Get-ADobject -filter "SamACcountName -eq '$($ObjectName)'" 
      
                if($ADObject.ObjectClass -eq "Group"){          
                    $SNSGroup =   Get-ADGroup -identity $ObjectName -IncludeMembership
                   
                    ForEach($username in $SNSGroup.UserMembers.Keys){

                        $IndirectLid = $SNSGroup.UserMembers[$username]
                        $properties = @{
                            Folder=$foldername;                           
                            DirectLid=$access.IdentityReference;
                            DirectLidType=$ADObject.ObjectClass;
                            IndirectLidNaam= $IndirectLid.Name;
                            IndirectLidId= $IndirectLid.Personeelsnummer;
                            IndirectLidINS=$IndirectLid.isINS
                
                        }

                        $list += New-Object -TypeName PSObject -Property  $properties 

                    }

                }

            
             
             
            }
        


        }
        $list

    }


}