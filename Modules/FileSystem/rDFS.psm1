Get-rDFSFolderInfo($dfsRoot){
    $result  = @()
    ForEach($dfs in $dfsRoot){
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

}