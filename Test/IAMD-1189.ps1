Import-Module "\\jumbo.local\hkt-usrdata\r\roblooman\Documents\GitHub\IdentityManager\Modules\IdentityManager\IdentityManager.psm1"

$cred =Get-Credential

$session = Connect-OIM -AppServer SBX-IAM-9001.sandbox.local -AppName D1IMAppServer  -Cred $cred  




$TestCaseall = Get-OIMObject -ObjectName PersonHasEset -Where "UID_Eset in (Select UID_Eset from ESET where UID_accProduct is not null and ident_eset like 'R10 - 4616 - Kader Incl Cash Office%'  ) and xOrigin <> 2 "  -Full
$TestCase1 = $TestCaseall[0]
$TestCase1Person = Get-OIMObject -ObjectName Person -Where "UID_Person = '$($TestCase1.UID_Person)'" -Full
$TestCase1Eset = Get-OIMObject -ObjectName Eset -Where "UID_Eset = '$($TestCase1.UID_Eset)'" -Full

$TestCase1PreReq =  Get-OIMObject -ObjectName PersonWantsOrg -Where "ObjectKeyOrdered  = '$($TestCase1Eset.XObjectKey)' AND UID_PersonInserted = '$($TestCase1Person.UID_Person)'" -First
if($TestCase1PreReq -ne $Null){Write-Error "Already has an it shop request"}


Start-OIMEvent -Object $TestCase1 -EventName Create_ITShop -Parameters @{}

$TestCase1Result =  Get-OIMObject -ObjectName PersonWantsOrg -Where "ObjectKeyOrdered  = '$($TestCase1Eset.XObjectKey)' AND UID_PersonInserted = '$($TestCase1Person.UID_Person)'" -First
if($TestCase1Result -eq $Null){Write-Error "Not Assign"}


. "Get-OIMObject -ObjectName PersonHasEset -Where ""UID_Eset in (Select UID_Eset from ESET where UID_accProduct is not null and ident_eset like 'R10 - 4616 - Kader Incl Cash Office%'  ) and xOrigin <> 2 """  -Full"