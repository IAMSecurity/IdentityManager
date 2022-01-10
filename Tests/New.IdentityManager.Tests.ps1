Param(
    $ObjectNew = @(
        @{Table = "Person"
           Properties = @{Firstname = 'a';lastname=''}}

    )
)
BeforeAll{
    $ModuleManifestName = 'IdentityManager.psd1'
    $ModuleManifestPath = "$PSScriptRoot\..\IdentityManager\$ModuleManifestName"
    $ModuleName = 'IdentityManager.psm1'
    $ModulePath = "$PSScriptRoot\..\IdentityManager\$ModuleName"
    Import-MOdule $ModulePath

}

Describe 'New <Table> Where <Properties>'  -ForEach $ObjectNew{
    It 'Passes Test-ModuleManifest' {
        $obj = New-OIMObject $Table -Properties $Properties -Verbose
        $obj | Should -not -BeNullOrEmpty


    }
}