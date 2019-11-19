$CSV = Import-CSV .\DFSFQDN.csv -UseCulture

ForEach($folder in $CSV){
    $OldFolderTarget = $folder.OldFolderTarget
    $NewFolderTarget = $folder.NewFolderTarget
    $Folder = $folder.Folder
    $Root = $folder.Root




    #"-------------"
    #$folder.Folder
    #$oldFolderTarget
    #$newFolderTarget
    Try{
        $old =  GEt-DfsnFolderTarget -Path $Folder -TargetPath $oldFolderTarget 
        $new =  GEt-DfsnFolderTarget -Path $Folder -TargetPath $newFolderTarget

    $old | Remove-DfsnFolderTarget -Force
        #New-DfsnFolderTarget -Path $Folder -TargetPath $newFolderTarget -State Online | out-Null
    }Catch{
         $oldFolderTarget 

    }
}
<#
$CSV |%{
    $_


}

#>

#$root = Get-DfsnRoot -ErrorAction SilentlyContinue

$result  = @()
ForEach($dfs in $root){
    $folders =  Get-DfsnFolder -Path "$($dfs.Path)\*"
    ForEach($dfsfolder in $folders){
        $foldertargets = Get-DfsnFolderTarget $dfsfolder.Path
        ForEach($foldertarget in $foldertargets){
            $server = $foldertarget.TargetPath.Split("\")[2]
            $properties = @{
                "Root"=$dfs.Path;
                "Folder"=$dfsfolder.Path;
                "FolderTarget"=$foldertarget.TargetPath;
                "Server"=$server;
                "State"=$foldertarget.State}
            $result += New-Object -TypeName PSObject -Property $properties
        }

        
    }
}

