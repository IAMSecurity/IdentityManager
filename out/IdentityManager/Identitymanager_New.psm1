Function New-OIMPerson {
    Param(
        $ApprovalState,$AuthentifierLogins,$BadPwdAnswerAttempts,$BirthDate,$Building,$CanonicalName,$CCC_BusinessCard,$CCC_Date_Block,$CCC_DisplayNamePreference,$CCC_FunctionCode,$CCC_JumboMailAddress,$CCC_LastFunctionUpdate,$CCC_OnPremImmutableID,$CCC_Origin,$CCC_PartnerSurname,$CCC_PartnerSurnamePrx,$CCC_RetainAccount,$CCC_Surname,$CCC_SurnamePrx,
            [Alias("CCC_Ident_MainJob")]
            $CCC_UID_MainJob
            ,
            [Alias("CCC_Ident_PersonHeadFromPerson")]
            $CCC_UID_PersonHeadFromPerson
            ,
            [Alias("CCC_Ident_stg_Person")]
            $CCC_UID_stg_Person
            ,$CentralAccount,$CentralSAPAccount,$City,$CompanyMember,$ContactEmail,$CustomProperty01,$CustomProperty02,$CustomProperty03,$CustomProperty04,$CustomProperty05,$CustomProperty06,$CustomProperty07,$CustomProperty08,$CustomProperty09,$CustomProperty10,$DateLastWorked,$DeactivationEnd,$DeactivationStart,$DefaultEmailAddress,$Description,$DialogUserPassword,$DialogUserSalt,$DisplayTelephoneBook,$DistinguishedName,$EmployeeType,$EntryDate,$ExitDate,$Fax,$FaxExtension,$FirstName,$Floor,$FormerName,$Gender,$GenerationalQualifier,$IdentityType,$ImportSource,$Initials,$InternalName,$IsCar,$IsDummyPerson,$IsDuplicateName,$IsExternal,$IsInActive,$IsLockedOut,$IsLockedPwdAnswer,$IsNoInherite,$IsNoteBookUser,$IsRemoteAccessAllowed,$IsSecurityIncident,$IsTemporaryDeactivated,$IsTerminalServerAllowed,$IsVIP,$IsX500Dummy,$JPegPhoto,$LastName,$MfaUserId,$MiddleName,$NameAddOn,$Passcode,$PasscodeExpires,$PasswordLastSet,$PersonalTitle,$PersonnelNumber,$Phone,$PhoneExtension,$PhoneMobile,$PostalOfficeBox,$PreferredName,$Remarks,$RiskIndexCalculated,$Room,$Salutation,$SecurityIdent,$Sponsor,$Street,$SubCompany,$TechnicalEntryDate,$Title,
            [Alias("Ident_Department")]
            $UID_Department
            ,
            [Alias("Ident_DialogCountry")]
            $UID_DialogCountry
            ,
            [Alias("Ident_DialogCulture")]
            $UID_DialogCulture
            ,
            [Alias("Ident_DialogCultureFormat")]
            $UID_DialogCultureFormat
            ,
            [Alias("Ident_DialogState")]
            $UID_DialogState
            ,
            [Alias("Ident_DialogUser")]
            $UID_DialogUser
            ,
            [Alias("Ident_FirmPartner")]
            $UID_FirmPartner
            ,
            [Alias("Ident_Locality")]
            $UID_Locality
            ,
            [Alias("Ident_Org")]
            $UID_Org
            ,$UID_Person,
            [Alias("Ident_PersonHead")]
            $UID_PersonHead
            ,
            [Alias("Ident_PersonMasterIdentity")]
            $UID_PersonMasterIdentity
            ,
            [Alias("Ident_ProfitCenter")]
            $UID_ProfitCenter
            ,
            [Alias("Ident_RealPerson")]
            $UID_RealPerson
            ,
            [Alias("Ident_WorkDesk")]
            $UID_WorkDesk
            ,
            [Alias("Ident_X500Person")]
            $UID_X500Person
            ,$UserIDTSO,$XDateInserted,$XDateUpdated,$XMarkedForDeletion,$XObjectKey,$XTouched,$XUserInserted,$XUserUpdated,$ZIPCode,$session=$Global:OIM_Session)
        $properties =  @{}
        foreach ($key in $MyInvocation.BoundParameters.keys){
           
            $ParameterVar = Get-Variable -Name $key               
            $ParamaterValue    = $ParameterVar.Value
 
            if($key.StartsWith("UID")){  
                
                $FKType     = $ParameterVar.Name.Split("_")[1] 
                $FKIdent    = "Ident_$FKType"
                $FKIdent    = $FKIdent.Replace("Ident_UNSRootB","Ident_UNSRoot")
                $FKIdent    = $FKIdent.Replace("Ident_Person","uid_Person")
                $FKObject   = Get-OIMObject -objectname $FKType -where "$FKIdent = '$ParamaterValue'"
                $properties.Add($Key,$FKObject.$key)
            }else{               
                $properties.Add($Key,$ParamaterValue)
               
            }
 
        }
        $body = @{values = $Properties } | ConvertTo-Json    
        Invoke-RestMethod -Uri "$Global:OIM_BaseURL/api/entity/person" -Method Post -WebSession $session -ContentType application/json  -Body $body
            
    
}
 
 
Function New-OIMPersonHasESet {
    Param(
       
            [Alias("Ident_ESet")]
            $UID_ESet
            ,
            [Alias("Ident_Person")]
            $UID_Person
            ,$XDateInserted,$XDateUpdated,$XIsInEffect,$XMarkedForDeletion,$XObjectKey,$XOrigin,$XTouched,$XUserInserted,$XUserUpdated,$session=$Global:OIM_Session)
    $properties =  @{}
        foreach ($key in $MyInvocation.BoundParameters.keys){
           
             
            $ParameterVar = Get-Variable -Name $key               
            $ParamaterValue    = $ParameterVar.Value
 
            if($key.StartsWith("UID")){  
                
                $FKType     = $ParameterVar.Name.Split("_")[1] 
                $FKIdent    = "Ident_$FKType"
                $FKIdent    = $FKIdent.Replace("Ident_UNSRootB","Ident_UNSRoot")
                $FKIdent    = $FKIdent.Replace("Ident_Person","uid_Person")
                Write-Warning "$FKType $FKIdent = '$ParamaterValue'"
                $FKObject   = Get-OIMObject -objectname $FKType -where "$FKIdent = '$ParamaterValue'"
                $properties.Add($Key,$FKObject.$key)
            }else{               
                $properties.Add($Key,$ParamaterValue)
               
            }
 
        }
        $body = @{values = $Properties } | ConvertTo-Json    
        Invoke-RestMethod -Uri "$Global:OIM_BaseURL/api/entity/personhaseset" -Method Post -WebSession $session -ContentType application/json  -Body $body
            
    
}