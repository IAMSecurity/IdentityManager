Function Get-rFolderSize($path){
    ForEAch($folder in Get-ChildItem $path  ){
        $subpath = Join-Path  $folder.FullName  "Recycler"
        If(Test-Path $subpath ){
        $size = $subpath  | Get-ChildItem -Recurse | Measure-Object -Sum Length | Select-Object Count, Sum
        "Size $([math]::Round($size.Sum /1mb)) MB for path $subpath"
        }
        $subpath = Join-Path  $folder.FullName  "Downloads"
        If(Test-Path $subpath ){
        $size = $subpath  | Get-ChildItem -Recurse | Measure-Object -Sum Length | Select-Object Count, Sum
        "Size $([math]::Round($size.Sum /1mb)) MB for path $subpath"
        } 
    }
}
Get-rFolderSize -Path "C:\IAM\ScriptRepository\Scripts"