function Import-Items {
    Param(
        $DocFieldsMap,
        $DocTemplate,
        $ImportEnabled,
        $NinjaOneItemFilter,
        $MigrationName,
        $ITGImports
    )


    $ImportsMigrated = 0

    $ImportTemplate = $null
	
    Write-Host "Processing $($DocTemplate.name)"

    # Lets try to match Asset Layouts
    $ImportTemplate = Invoke-NinjaOneRequest -Path "document-templates" -Method GET -QueryParams "templateName=$($DocTemplate.name)" | where-object { $_.name -eq $DocTemplate.name }
	
    if ($ImportTemplate) {
        Write-Host "$MigrationName layout found attempting to match existing entries"

        $OrgIDs = $CompaniesToMigrate.NinjaOneID -join ','
        $NinjaOneImports = Invoke-NinjaOneRequest -Path "organization/documents" -Method GET -QueryParams "templateIds=$($ImportTemplate.id)&organizationId=s=$($OrgIDs)"

        $MatchedImports = foreach ($itgimport in $ITGImports ) {
            $NinjaOneOrgID = ($MatchedCompanies | Where-Object { $_.ITGID -eq $itgimport.attributes."organization-id" }).NinjaOneID
            $NinjaOneImport = $NinjaOneImports | where-object -filter $NinjaOneItemFilter
			
	
            if ($NinjaOneImport) {
                [PSCustomObject]@{
                    "Name"           = $itgimport.attributes.name
                    "CompanyName"    = $itgimport.attributes."organization-name"
                    "NinjaOneOrgID"  = $NinjaOneOrgID
                    "ITGID"          = $itgimport.id
                    "NinjaOneID"     = $NinjaOneImport.id
                    "Matched"        = $true
                    "NinjaOneObject" = $NinjaOneImport
                    "ITGObject"      = $itgimport
                    "Imported"       = "Pre-Existing"
					
                }
            } else {
                [PSCustomObject]@{
                    "Name"           = $itgimport.attributes.name
                    "CompanyName"    = $itgimport.attributes."organization-name"
                    "NinjaOneOrgID"  = $NinjaOneOrgID
                    "ITGID"          = $itgimport.id
                    "NinjaOneID"     = ""
                    "Matched"        = $false
                    "NinjaOneObject" = ""
                    "ITGObject"      = $itgimport
                    "Imported"       = ""
                }
            }
        }
    } else {
        $MatchedImports = foreach ($itgimport in $ITGImports ) {
            $NinjaOneOrgID = ($MatchedCompanies | Where-Object { $_.ITGID -eq $itgimport.attributes."organization-id" }).NinjaOneID
            [PSCustomObject]@{
                "Name"           = $itgimport.attributes.name
                "CompanyName"    = $itgimport.attributes."organization-name"
                "NinjaOneOrgID"  = $NinjaOneOrgID
                "ITGID"          = $itgimport.id
                "NinjaOneID"     = ""
                "Matched"        = $false
                "NinjaOneObject" = ""
                "ITGObject"      = $itgimport
                "Imported"       = ""
            }
		
        }

    }
	
    Write-Host "Matched $MigrationName (Already exist so will not be migrated)"
    Write-Host $($MatchedImports | Sort-Object CompanyName | Where-Object { $_.Matched -eq $true } | Select-Object CompanyName, Name | Format-Table | Out-String)
	
    Write-Host "Unmatched $MigrationName"
    Write-Host $($MatchedImports | Sort-Object CompanyName | Where-Object { $_.Matched -eq $false } | Select-Object CompanyName, Name | Format-Table | Out-String)

    # Import Items
    $UnmappedImportCount = ($MatchedImports | Where-Object { $_.Matched -eq $false } | measure-object).count
    if ($ImportEnabled -eq $true -and $UnmappedImportCount -gt 0) {
		
        $ImportTemplate = Invoke-NinjaOneDocumentTemplate -Template $DocTemplate
	
        $ImportOption = Get-ImportMode -ImportName $MigrationName
	
        if (($importOption -eq "A") -or ($importOption -eq "S") ) {		
	
            foreach ($company in $CompaniesToMigrate) {
                Write-Host "Migrating $($company.CompanyName) $MigrationName"
	
                foreach ($unmatchedImport in ($MatchedImports | Where-Object { $_.Matched -eq $false -and $company.ITGCompanyObject.id -eq $_."ITGObject".attributes."organization-id" })) {
	
                    $DocumentFields = & $DocFieldsMap

                    Confirm-Import -ImportObjectName "$($unmatchedImport.Name): $($AssetFields | Out-String)" -ImportObject $unmatchedImport -ImportSetting $ImportOption
	
                    Write-Host "Starting $($unmatchedImport.Name)"

                    $CreateObject = [PSCustomObject]@{
                        documentName        = $unmatchedImport.Name
                        documentDescription = $unmatchedImport.Description ?? ""
                        documentTemplateId  = $ImportTemplate.id
                        organizationId      = $unmatchedImport.NinjaOneOrgID
                        fields              = $DocumentFields
                    }
                    try {
                        try {
                            $NinjaOneNewImport = Invoke-NinjaOneRequest -Path "organization/documents" -Method POST -InputObject $CreateObject -AsArray -ea stop

                            $unmatchedImport.matched = $true
                            $unmatchedImport.NinjaOneID = $NinjaOneNewImport.id
                            $unmatchedImport."NinjaOneObject" = $NinjaOneNewImport
                            $unmatchedImport.Imported = "Created-By-Script"

                        } catch {
                            $unmatchedImport.matched = $false
                            $unmatchedImport.NinjaOneID = $null
                            $unmatchedImport."NinjaOneObject" = $Null
                            $unmatchedImport.Imported = "Creation Failed - $_"

                            Throw "Failed to create item $($unmatchedImport.Name): $_"
                        }

                        $ImportsMigrated = $ImportsMigrated + 1
	
                        Write-host "$($unmatchedImport.Name) Has been created in NinjaOne"
					
                        if ($itgimport.attributes.archived) {
                            Write-Host "WARNING: $($unmatchedImport.name) is archived in ITGlue and is being archived in NinjaOne" -ForegroundColor Magenta
                            try {
                                $Null = Invoke-NinjaOneRequest -Path "organization/document/$($NinjaOneNewImport.id)/archive" -Method POST -ea stop
                            } catch {
                                Throw "Failed to archive item $($unmatchedImport.Name): $_"
                            }
                        }
	
                        Write-Host ""
                    } catch {
                        Write-Error $_
                    }

                }
            }
        }
			
    } else {
        if ($UnmappedImportCount -eq 0) {
            Write-Host "All $MigrationName matched, no migration required" -foregroundcolor green
        } else {
            Write-Host "Warning Import $MigrationName is set to disabled so the above unmatched $MigrationName will not have data migrated" -foregroundcolor red
            Read-Host -Prompt "Press any key to continue or CTRL+C to quit" 
        }
    }
	
    Return $MatchedImports

}
