cls 
$dir = "\\SVPV06181.verz.local\d$\Sys\Data\Financial Reporting & Information Management"

ForEach($folder in Get-ChildItem -Path $dir -Recurse -Directory){
    $path =   $folder.FullName
    $acl = Get-ACL -Path $path 
    if($acl.AreAccessRulesProtected){
        $isProtected = $false
        $preserveInheritance = $true
        $acl.SetAccessRuleProtection($isProtected, $preserveInheritance)

        $acl | Set-ACl -Path $path  
        Write-Host $path  
    }

 }

 $dirs = $("\\verz.local\afdeling$\Werkplek Services")
$env:PSModulePath +=";\\verz.local\AFdeling$\Werkplek Services\RAP\99. Automation\Modules" 

Function Get-SNSAclCustom {
    Param($foldername)
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



            if(($access.IdentityReference  -like "VERZ\*" -or $access.IdentityReference -like "BANK\*")-AND (! $access.IsInherited)){
            


                $splitMember = $access.IdentityReference -split "\\"            
                $ObjectName = $splitMember[1]
                $ADObject = Get-ADobject -filter "SamACcountName -eq '$($ObjectName)'" 
      
                if($ADObject.ObjectClass -eq "Group"){
                    $group = Get-SNSADGroup -Name $ObjectName
                    $SNSGroup =   $group  | Get-SNSADGroupMemberInfo 
                   
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
Import-Module SNSActiveDirectory -Force
Clear-SNSCache

ForEach($dir in $dirs){
    $dirNaam = $dir.Replace("\","")
    $list = @()
     $fldr.dir 
    $list += Get-SNSAclCustom $dir

    ForEach($fldr in Get-Childitem $dir -Recurse -Directory){
        $fldr.FullName 
         $list += Get-SNSAclCustom $fldr.FullName 
   
    }
    $list |Select-Object Folder,DirectLid,DirectLidType,IndirectLidNaam,IndirectLidID,IndirectLidINS| Export-Csv "C:\Beheer\Mark\exports\ListACL4_$($dirNaam).csv"  -UseCulture -NoTypeInformation 
}



$dirs = $("\\svpn20050.verz.local\alge02$\Data\Coda\01. CODA-XL")
$env:PSModulePath +=";\\verz.local\AFdeling$\Werkplek Services\RAP\99. Automation\Modules" 

Function Get-SNSAclCustom {
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



            if(($access.IdentityReference  -like "VERZ\*" -or $access.IdentityReference -like "BANK\*")-AND ((!$access.IsInherited)-or $Inheritance)){
            


                $splitMember = $access.IdentityReference -split "\\"            
                $ObjectName = $splitMember[1]
                $ADObject = Get-ADobject -filter "SamACcountName -eq '$($ObjectName)'" 
      
                if($ADObject.ObjectClass -eq "Group"){          
                    $SNSGroup =   Get-SNSADGroup -identity $ObjectName -IncludeMembership
                   
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
Import-Module SNSActiveDirectory -Force
Clear-SNSCache

ForEach($dir in $dirs){
    $dirNaam = $dir.Replace("\","")
    $list = @()
     $fldr.dir 
    $list += Get-SNSAclCustom $dir -Inheritance

    ForEach($fldr in Get-Childitem $dir -Recurse -Directory){
        $fldr.FullName 
         $list += Get-SNSAclCustom $fldr.FullName 
   
    }
    $list |Select-Object Folder,DirectLid,DirectLidType,IndirectLidNaam,IndirectLidID,IndirectLidINS| Export-Csv "C:\Beheer\Mark\exports\AUDITZWL_$($dirNaam).csv"  -UseCulture -NoTypeInformation 
}