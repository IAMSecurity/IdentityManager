Param(
    $ObjectWhere = @(
        @{Table = "Eset"
           Where = "Ident_Eset = 'Test'"}

    ),
    $ObjectId = @(
        @{Table = "Eset"
           Id=""}

    )
)
BeforeAll{
    $ModuleManifestName = 'IdentityManager.psd1'
    $ModuleManifestPath = "$PSScriptRoot\..\IdentityManager\$ModuleManifestName"
    $ModuleName = 'IdentityManager.psm1'
    $ModulePath = "$PSScriptRoot\..\IdentityManager\$ModuleName"
    Import-MOdule $ModulePath

}

Describe 'Get <Table> Where <where>'  -ForEach $ObjectWhere{
    It 'Passes Test-ModuleManifest' {
        $obj = Get-OIMObject $Table -where $where
        $obj | Should -not -BeNullOrEmpty

        $obj2 = $obj | Get-OIMObject
        $obj2 | Should -not -BeNullOrEmpty
        if(-not [string]::IsNullOrEmpty($description) ){
            $obj2.description| Should -be $description

        }
    }
}
Describe 'Get <Table> Id <Id>'  -ForEach $ObjectId{
    It 'Passes Test-ModuleManifest' {
        $obj = Get-OIMObject $Table -Id $Id
        $obj | Should -not -BeNullOrEmpty
        $obj."UID_$Table" | Should -be  $Id

        $obj2 = $obj | Get-OIMObject
        $obj2 | Should -not -BeNullOrEmpty
        if(-not [string]::IsNullOrEmpty($description) ){
            $obj2.description| Should -be $description

        }
    }
}
