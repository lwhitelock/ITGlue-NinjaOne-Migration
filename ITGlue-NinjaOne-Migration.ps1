# Main settings load
. $PSScriptRoot\Initialize-Module.ps1 -InitType 'Full'

# Use this to set the context of the script runs
$FirstTimeLoad = 1

############################### Functions ###############################
# Import NinjaOne Helper Functions
. $PSScriptRoot\Public\Import-NinjaOneHelpers.ps1

# Import ImageMagick for Invoke-ImageTest Function (Disabled)
. $PSScriptRoot\Private\Initialize-ImageMagik.ps1

# Used to determine if a file is an image and what type of image
. $PSScriptRoot\Private\Invoke-ImageTest.ps1

# Confirm Object Import
. $PSScriptRoot\Private\Confirm-Import.ps1

# Matches items from IT Glue to NinjaONe and creates new items in NinjaONe
. $PSScriptRoot\Private\Import-Items.ps1

# Select Item Import Mode
. $PSScriptRoot\Private\Get-ImportMode.ps1

# Get Configurations Option
. $PSScriptRoot\Private\Get-ConfigurationsImportMode.ps1

# Get Flexible Asset Layout Option
. $PSScriptRoot\Private\Get-FlexLayoutImportMode.ps1

# Fetch Items from ITGlue
. $PSScriptRoot\Private\Import-ITGlueItems.ps1

# Find migrated items
. $PSScriptRoot\Private\Find-MigratedItem.ps1

# Add Replace URL functions
. $PSScriptRoot\Private\ConvertTo-NinjaOneURL.ps1

# Add Timed (Noninteractive) Messages Helper
. $PSScriptRoot\Public\Write-TimedMessage.ps1

# Add numeral casting helper method
. $PSScriptRoot\Public\Get-CastIfNumeric.ps1

# Add migration scope helper
. $PSScriptRoot\Public\Set-MigrationScope.ps1


############################### End of Functions ###############################


###################### Initial Setup and Confirmations ###############################
Write-Host "###########################################################" -ForegroundColor Green
Write-Host "#                                                         #" -ForegroundColor Green
Write-Host "#       IT Glue to NinjaONe Migration Script              #" -ForegroundColor Green
Write-Host "#                                                         #" -ForegroundColor Green
Write-Host "#          Version: 1.0-Alpha                             #" -ForegroundColor Green
Write-Host "#          Date: 09/08/2025                               #" -ForegroundColor Green
Write-Host "#                                                         #" -ForegroundColor Green
Write-Host "#          Author: Luke Whitelock                         #" -ForegroundColor Green
Write-Host "#                  https://mspp.io                        #" -ForegroundColor Green
Write-Host "#                                                         #" -ForegroundColor Green
Write-Host "##########################################################" -ForegroundColor Green
Write-Host "# Note: This is an unofficial script, please do not       #" -ForegroundColor Green
Write-Host "# contact NinjaOne support if you run into issues.        #" -ForegroundColor Green
Write-Host "# For support please visit the NinjaOne Discord           #" -ForegroundColor Green
Write-Host "# https://discord.gg/NinjaOne                             #" -ForegroundColor Green
Write-Host "# Or log an issue in the Github Respository:              #" -ForegroundColor Green
Write-Host "# https://github.com/lwhitelock/ITGlue-NinjaOne-Migration #" -ForegroundColor Green
Write-Host "#######################################################" -ForegroundColor Green
Write-Host " Instructions:                                       " -ForegroundColor Green
Write-Host " Please view the documentation here:                       " -ForegroundColor Green
Write-Host " https://mspp.io/automated-it-glue-to-ninjaone-migration-script/     " -ForegroundColor Green
Write-Host " for detailed instructions                           " -ForegroundColor Green
Write-Host "#######################################################" -ForegroundColor Green
Write-Host "# Please keep ALL COPIES of the Migration Logs folder. This can save you." -ForegroundColor Gray
Write-Host "# Please DO NOT CHANGE ANYTHING in the Migration Logs folder. This can save you." -ForegroundColor Gray

# CMA
Write-Host "######################################################" -ForegroundColor Red
Write-Host "This Script has the potential to ruin your NinjaOne Documentation environment" -ForegroundColor Red
Write-Host "It is unofficial and you run it entirely at your own risk" -ForegroundColor Red
Write-Host "You accept full responsibility for any problems caused by running it" -ForegroundColor Red
Write-Host "######################################################" -ForegroundColor Red

$backups = $(if ($true -eq $NonInteractive) { "Y" } else { Read-Host "Y/n" })

$ScriptStartTime = $(Get-Date -Format "o")

if ($backups -ne "Y" -or $backups -ne "y") {
    Write-Host "Please take a backup and run the script again"
    exit 1
}

if ((get-host).version.major -ne 7) {
    Write-Host "Powershell 7 Required" -foregroundcolor Red
    exit 1
}

  
#Login to NinjaOne
try {
    Write-Host "Connecting to NinjaOne, please login via the web browser window that was launched and then return here"
    Connect-NinjaOne -NinjaOneInstance $NinjaOneBaseDomain -NinjaOneClientID $NinjaOneClientID -NinjaOneClientSecret $NinjaOneClientSecret -ea Stop
    Write-Host "Successfully connected to NinjaOne"
} catch {
    Write-Host "Failed to connect to NinjaOne: $_"
    exit 1
}


try {
    remove-module ITGlueAPI -ErrorAction SilentlyContinue
} catch {
}
#Grabbing ITGlue Module and installing.
If (Get-Module -ListAvailable -Name "ITGlueAPIv2") { 
    Import-module ITGlueAPIv2 
} Else { 
    Install-Module ITGlueAPIv2 -Force
    Import-Module ITGlueAPIv2
}
#Settings IT-Glue logon information
Add-ITGlueBaseURI -base_uri $ITGAPIEndpoint
Add-ITGlueAPIKey $ITGKey

# Check if we have a logs folder
if (Test-Path -Path "$MigrationLogs") {
    if ($ResumePrevious -eq $true) {
        Write-Host "A previous attempt has been found job will be resumed from the last successful section" -ForegroundColor Green
        $ResumeFound = $true
    } else {
        Write-Host "A previous attempt has been found, resume is disabled so this will be lost, if you haven't reverted to a snapshot, a resume is recommended" -ForegroundColor Red
        Write-TimedMessage -Timeout 12 -Message "Press any key to continue or ctrl + c to quit and edit the ResumePrevious setting" -DefaultResponse "proceed with new migration, do not resume"
        $ResumeFound = $false
    }
} else {
    Write-Host "No previous runs found creating log directory"
    $null = New-Item "$MigrationLogs" -ItemType "directory"
    $ResumeFound = $false
}


# Setup some variables

$ManualActions = [System.Collections.ArrayList]@()


############################### Organizations and Locations ###############################

#Grab existing Organizations in NinjaOne
$NinjaOneOrganizations = Invoke-NinjaOneRequest -Method GET -Path 'organizations'

#Check for Company Resume
if ($ResumeFound -eq $true -and (Test-Path "$MigrationLogs\Companies.json")) {
    Write-Host "Loading Previous Companies Migration"
    $MatchedCompanies = Get-Content "$MigrationLogs\Companies.json" -raw | Out-String | ConvertFrom-Json
    if ($ResumeFound -eq $true -and (Test-Path "$MigrationLogs\Locations.json")) {
        Write-Host "Loading Previous Locations Migration"
        $MatchedLocations = Get-Content "$MigrationLogs\Locations.json" -raw | Out-String | ConvertFrom-Json -depth 100
    }
} else {

    #Import Companies
    Write-Host "Fetching Companies from IT Glue" -ForegroundColor Green
    $CompanySelect = { (Get-ITGlueOrganizations -page_size 1000 -page_number $i).data }
    $ITGCompanies = Import-ITGlueItems -ItemSelect $CompanySelect

    Write-Host "$($ITGCompanies.count) IT Glue Companies Found" 
    if ($ScopedMigration) {
        $OriginalCompanyCount = $($ITGcompanies.count)
        Write-Host "Setting companies to those in scope..." -foregroundcolor Yellow
        $ITGCompanies = Set-MigrationScope -AllITGCompanies $ITGCompanies -InternalCompany $InternalCompany
        Write-Host "Companies scoped... $OriginalCompanyCount => $($Itgcompanies.count)"
    }
    $ScopedITGCompanyIds = $ITGCompanies.id

    $MatchedCompanies = foreach ($itgcompany in $ITGCompanies ) {
        $NinjaOneOrganization = $NinjaOneOrganizations | where-object -filter { $_.name -eq $itgcompany.attributes.name }
        if ($InternalCompany -eq $itgcompany.attributes.name) {
            $intCompany = $true
        } else {
            $intCompany = $false
        }
	
        if ($NinjaOneOrganization) {
            [PSCustomObject]@{
                "CompanyName"                = $itgcompany.attributes.name
                "ITGID"                      = $itgcompany.id
                "NinjaOneID"                 = $NinjaOneOrganization.id
                "Matched"                    = $true
                "InternalCompany"            = $intCompany
                "NinjaOneOrganizationObject" = $NinjaOneOrganization
                "ITGCompanyObject"           = $itgcompany
                "Imported"                   = "Pre-Existing"
            }
        } else {
            [PSCustomObject]@{
                "CompanyName"                = $itgcompany.attributes.name
                "ITGID"                      = $itgcompany.id
                "NinjaOneID"                 = ""
                "Matched"                    = $false
                "InternalCompany"            = $intCompany
                "NinjaOneOrganizationObject" = ""
                "ITGCompanyObject"           = $itgcompany
                "Imported"                   = ""
            }
        }
    }

    # Check if the internal company was found and that there was only 1 of them
    $PrimaryCompany = $MatchedCompanies | Sort-Object CompanyName | Where-Object { $_.InternalCompany -eq $true } | Select-Object CompanyName

    if (($PrimaryCompany | measure-object).count -ne 1) {
        Write-Host "A single Internal Company was not found please run the script again and check the company name entered exactly matches what is in ITGlue" -foregroundcolor red
        exit 1
    }

    # Lets confirm it is the correct one
    Write-Host ""
    Write-Host "Your Internal Company has been matched to: $(($MatchedCompanies | Sort-Object CompanyName | Where-Object {$_.InternalCompany -eq $true} | Select-Object CompanyName).companyname) in IT Glue"
    Write-Host "The documents under this customer will be migrated to the Global KB in NinjaOne"
    Write-Host ""
    Write-TimedMessage -Message "Internal Company Correct? Press Return to continue or CTRL+C to quit if this is not correct" -Timeout 12 -DefaultResponse "Assuming found match on '$(($MatchedCompanies | Sort-Object CompanyName | Where-Object {$_.InternalCompany -eq $true} | Select-Object CompanyName).companyname)' is correct."

    Write-Host "Matched Companies (Already exist so will not be migrated)"
    $MatchedCompanies | Sort-Object CompanyName | Where-Object { $_.Matched -eq $true } | Select-Object CompanyName | Format-Table

    Write-Host "Unmatched Companies"
    $MatchedCompanies | Sort-Object CompanyName | Where-Object { $_.Matched -eq $false } | Select-Object CompanyName | Format-Table

    #Import Locations
    Write-Host "Fetching Locations from IT Glue" -ForegroundColor Green
    $LocationsSelect = { (Get-ITGlueLocations -page_size 1000 -page_number $i -include related_items).data }
    $ITGLocations = Import-ITGlueItems -ItemSelect $LocationsSelect
    if ($ScopedMigration) {
        $OriginalLocationsCount = $($ITGLocations.count)
        Write-Host "Setting locations to those in scope..." -foregroundcolor Yellow
        $ITGLocations = $ITGLocations | Where-Object { $ScopedITGCompanyIds -contains $_.attributes.'organization-id' }
        Write-Host "locations scoped... $OriginalLocationsCount => $($ITGLocations.count)"
    }

    # Import Companies
    $UnmappedCompanyCount = ($MatchedCompanies | Where-Object { $_.Matched -eq $false } | measure-object).count
    if ($ImportCompanies -eq $true -and $UnmappedCompanyCount -gt 0) {
	
        $importCOption = Get-ImportMode -ImportName "Companies"
	
        if (($importCOption -eq "A") -or ($importCOption -eq "S") ) {		
            foreach ($unmatchedcompany in ($MatchedCompanies | Where-Object { $_.Matched -eq $false })) {
                Confirm-Import -ImportObjectName $unmatchedcompany.CompanyName -ImportObject $unmatchedcompany -ImportSetting $importCOption
						
                Write-Host "Starting $($unmatchedcompany.CompanyName)"
                $PrimaryLocation = $ITGLocations | Where-Object { $unmatchedcompany.ITGID -eq $_.attributes."organization-id" -and $_.attributes.primary -eq $true }

                $OrgDescription = $unmatchedcompany.ITGCompanyObject.attributes."description"
                if ($null -ne $OrgDescription) {
                    if ($OrgDescripion.length -ge 1000) {
                        $OrgDescripion = $OrgDescripion.SubString(0, 999)
                    }
                } else {
                    $OrgDescription = ''
                }

                if ($PrimaryLocation -and $PrimaryLocation.count -eq 1) {

                    $LocDescription = $PrimaryLocation.attributes."notes"
                    if ($null -ne $LocDescription) {
                        if ($LocDescription.length -ge 250) {
                            $LocDescription = $LocDescription.SubString(0, 249)
                        }
                    } else {
                        $LocDescription = ''
                    }
                    
                    $OrgCreation = @{
                        name        = $unmatchedcompany.CompanyName
                        description = $OrgDescription
                        locations   = @(
                            @{
                                name        = $PrimaryLocation.attributes."name"
                                address     = (@($PrimaryLocation.attributes."address-1", $PrimaryLocation.attributes."address-2", $PrimaryLocation.attributes.city, $PrimaryLocation.attributes."region-name", $PrimaryLocation.attributes."postal-code", $PrimaryLocation.attributes."country-name") | Where-Object { $null -ne $_ }) -join "`r`n"
                                description = $LocDescription
                            }
                        )
                    }
                } else {
                    $OrgCreation = @{
                        name        = $unmatchedcompany.CompanyName
                        description = $OrgDescription
                    }
                }


                $NinjaOneNewOrganization = Invoke-NinjaOneRequest -Path "organizations" -Method POST -InputObject $OrgCreation
			
                $unmatchedcompany.matched = $true
                $unmatchedcompany.NinjaOneID = $NinjaOneNewOrganization.id
                $unmatchedcompany.NinjaOneOrganizationObject = $NinjaOneNewOrganization
                $unmatchedcompany.Imported = "Created-By-Script"
			
                Write-host "$($unmatchedcompany.CompanyName) Has been created in NinjaOne"
                Write-Host ""
            }
		
        }
		

    } else {
        if ($UnmappedCompanyCount -eq 0) {
            Write-Host "All Companies matched, no migration required" -foregroundcolor green
        } else {
            Write-Host "Warning Import Companies is set to disabled so the above unmatched companies will not have data migrated" -foregroundcolor red
            Write-TimedMessage -Message "Press any key to continue or CTRL+C to quit" -DefaultResponse "continue and wrap-up companies, please." -Timeout 6
        }
    }

    # Save the results to resume from if needed
    $MatchedCompanies | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\Companies.json"
    Write-TimedMessage -Timeout 3 -Message "Snapshot Point: Companies Migrated Continue?"  -DefaultResponse "continue to Locations, please."

}

$CompaniesToMigrate = $MatchedCompanies | Sort-Object CompanyName | Where-Object { $_.Matched -eq $true }

$NinjaOneOrganizations = Invoke-NinjaOneRequest -Method GET -Path 'organizations'


############################### Locations ###############################
#Check for Location Resume
if ($ResumeFound -eq $true -and (Test-Path "$MigrationLogs\Locations.json")) {
    Write-Host "Loading Previous Locations Migration"
    $MatchedLocations = Get-Content "$MigrationLogs\Locations.json" -raw | Out-String | ConvertFrom-Json -depth 100
} else {

    #Grab existing Locations in NinjaOne
    $NinjaOneLocations = Invoke-NinjaOneRequest -Method GET -Path 'locations'

    $LocationsSelect = { (Get-ITGlueLocations -page_size 1000 -page_number $i -include related_items).data }
    $ITGLocations = Import-ITGlueItems -ItemSelect $LocationsSelect

    Write-Host "$($ITGLocations.count) IT Glue Locations Found" 
    if ($ScopedMigration) {
        $OriginalLocationsCount = $($ITGLocations.count)
        Write-Host "Setting locations to those in scope..." -foregroundcolor Yellow
        $ITGLocations = $ITGLocations | Where-Object { $ScopedITGCompanyIds -contains $_.attributes.'organization-id' }
        Write-Host "locations scoped... $OriginalLocationsCount => $($ITGLocations.count)"
    }

    $MatchedLocations = foreach ($itglocation in $ITGLocations ) {

        $NinjaOrgID = ($MatchedCompanies | Where-Object { $_.ITGID -eq $itglocation.attributes."organization-id" }).NinjaOneID

        If ($null -eq $NinjaOrgID) {
            Write-Error "Failed to find NinjaOne Organization ID: $itglocation"
            Continue
        }

        $NinjaOneLocation = $NinjaOneLocations | where-object -filter { $_.name -eq $itglocation.attributes.name -and $_.organizationId -eq $NinjaOrgID }

        if ($NinjaOneLocation) {
            [PSCustomObject]@{
                "Name"           = $itglocation.attributes.name
                "CompanyName"    = $itglocation.attributes."organization-name"
                "ITGID"          = $itglocation.id
                "NinjaOneID"     = $NinjaOneLocation.id
                "Matched"        = $true
                "NinjaOneObject" = $NinjaOneLocation
                "ITGObject"      = $itglocation
                "Imported"       = "Pre-Existing"
                "NinjaOneOrgID"  = $NinjaOrgID
					
            }
        } else {
            [PSCustomObject]@{
                "Name"           = $itglocation.attributes.name
                "CompanyName"    = $itglocation.attributes."organization-name"
                "ITGID"          = $itglocation.id
                "NinjaOneID"     = ""
                "Matched"        = $false
                "NinjaOneObject" = ""
                "ITGObject"      = $itglocation
                "Imported"       = ""
                "NinjaOneOrgID"  = $NinjaOrgID
            }
        }
    }

    Write-Host "Matched Locations (Already exist so will not be migrated)"
    $MatchedLocations | Sort-Object CompanyName, Name | Where-Object { $_.Matched -eq $true } | Select-Object CompanyName, Name | Format-Table

    Write-Host "Unmatched Locations"
    $MatchedLocations | Sort-Object CompanyName, Name | Where-Object { $_.Matched -eq $false } | Select-Object CompanyName, Name | Format-Table
   
    # Import Locations
    $UnmappedLocationsCount = ($MatchedLocations | Where-Object { $_.Matched -eq $false } | measure-object).count
    if ($ImportLocations -eq $true -and $UnmappedLocationsCount -gt 0) {
	
        $importLOption = Get-ImportMode -ImportName "Locations"
	
        if (($importLOption -eq "A") -or ($importLOption -eq "S") ) {		
            foreach ($unmatchedlocation in ($MatchedLocations | Where-Object { $_.Matched -eq $false })) {
                Confirm-Import -ImportObjectName $unmatchedlocation.CompanyName -ImportObject $unmatchedlocation -ImportSetting $importLOption
                
                Write-Host "Starting $($unmatchedlocation.Name)"
                
                $LocationCreation = @{
                    name        = $unmatchedlocation.ITGObject.attributes."name"
                    address     = (@($unmatchedlocation.ITGObject.attributes."address-1", $unmatchedlocation.ITGObject.attributes."address-2", $unmatchedlocation.ITGObject.attributes.city, $unmatchedlocation.ITGObject.attributes."region-name", $unmatchedlocation.ITGObject.attributes."postal-code", $unmatchedlocation.ITGObject.attributes."country-name") | Where-Object { $null -ne $_ }) -join "`r`n"
                    description = $(try { ($unmatchedlocation.ITGObject.attributes."notes").SubString(0, 249) } catch { $null })
                }
                
                $NinjaOneNewLocation = Invoke-NinjaOneRequest -Path "organization/$($unmatchedlocation.NinjaOneOrgID)/locations" -Method POST -InputObject $LocationCreation
			
                $unmatchedlocation.matched = $true
                $unmatchedlocation.NinjaOneID = $NinjaOneNewLocation.id
                $unmatchedlocation.NinjaOneObject = $NinjaOneNewLocation
                $unmatchedlocation.Imported = "Created-By-Script"
			
                Write-host "$($unmatchedlocation.Name) Has been created in NinjaOne"
                Write-Host ""
            }
		
        }
		
    } else {
        if ($UnmappedLocationsCount -eq 0) {
            Write-Host "All Locations matched, no migration required" -foregroundcolor green
        } else {
            Write-Host "Warning Import Locations is set to disabled so the above unmatched locations will not have data migrated" -foregroundcolor red
            Write-TimedMessage -Message "Press any key to continue or CTRL+C to quit" -DefaultResponse "continue and wrap-up locations, please." -Timeout 6
        }
    }

    # Save the results to resume from if needed
    $MatchedLocations | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\Locations.json"
    Write-TimedMessage -Timeout 3 -Message "Snapshot Point: Locations Migrated Continue?"  -DefaultResponse "continue to Websites, please."

}


############################### Domains ###############################

#Check for Website Resume
if ($ResumeFound -eq $true -and (Test-Path "$MigrationLogs\Domains.json")) {
    Write-Host "Loading Previous Domains Migration"
    $MatchedDomains = Get-Content "$MigrationLogs\Domains.json" -raw | Out-String | ConvertFrom-Json
} else {

    #Import Websites
    Write-Host "Fetching Domains from IT Glue" -ForegroundColor Green
    $DomainSelect = { (Get-ITGlueDomains -page_size 1000 -page_number $i).data }
    $ITGDomains = Import-ITGlueItems -ItemSelect $DomainSelect
    if ($ScopedMigration) {
        $OriginalDomainsCount = $($ITGDomains.count)
        Write-Host "Setting domains to those in scope..." -foregroundcolor Yellow
        $ITGDomains = $ITGdomains | Where-Object { $ScopedITGCompanyIds -contains $_.attributes.'organization-id' }
        Write-Host "domains scoped... $OriginalDomainsCount => $($ITGDomains.count)"
    }

    Write-Host "$($ITGDomains.count) ITG Glue Domains Found" 

    $DomainNinjaOneItemFilter = { ($_.documentName -eq $itgimport.attributes.name -and $_.organizationId -eq $NinjaOneOrgID) }

    $DomainsImportEnabled = $ImportDomains

    $DomainMigrationName = "Domains"

    $DomainTemplate = [PSCustomObject]@{
        name          = "$($FlexibleLayoutPrefix)$($DomainImportAssetLayoutName)"
        allowMultiple = $true
        fields        = @(
            [PSCustomObject]@{
                fieldLabel                = 'Notes'
                fieldName                 = 'notes'
                fieldType                 = 'WYSIWYG'
                fieldTechnicianPermission = 'EDITABLE'
                fieldScriptPermission     = 'READ_WRITE'
                fieldApiPermission        = 'READ_WRITE'
                fieldContent              = @{
                    required         = $False
                    advancedSettings = @{
                        expandLargeValueOnRender = $True
                    }
                }
            }
        )
    }


    $DomainDocFieldsMap = { @{
            'notes' = @{ 'html' = $unmatchedImport."ITGObject".attributes."notes" ?? "" } 
        } }
    


    $DomainImportSplat = @{
        DocFieldsMap       = $DomainDocFieldsMap
        DocTemplate        = $DomainTemplate
        ImportEnabled      = $DomainsImportEnabled
        NinjaOneItemFilter = $DomainNinjaOneItemFilter
        MigrationName      = $DomainMigrationName
        ITGImports         = $ITGDomains
    }

    #Import Domains
    $MatchedDomains = Import-Items @DomainImportSplat

    Write-Host "Doamins Complete"

    # Save the results to resume from if needed
    $MatchedDomains | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\Domains.json"
    Write-TimedMessage -Timeout 3 -Message "Snapshot Point: Domains Migrated Continue?"  -DefaultResponse "continue to Configurations, please."

}


		
############################### Configurations ###############################
	
	
#Check for Configuration Resume
if ($ResumeFound -eq $true -and (Test-Path "$MigrationLogs\Devices.json")) {
    Write-Host "Loading Previous Configurations Migration"
    $MatchedDevices = Get-Content "$MigrationLogs\Devices.json" -raw | Out-String | ConvertFrom-Json -depth 100
} else {

    #Get Configurations from IT Glue
    Write-Host "Fetching Configurations from IT Glue" -ForegroundColor Green
    $ConfigurationsSelect = { (Get-ITGlueConfigurations -page_size 1000 -page_number $i -include related_items).data }
    $ITGConfigurations = Import-ITGlueItems -ItemSelect $ConfigurationsSelect
    if ($ScopedMigration) {
        $OriginalConfigurationCount = $($ITGConfigurations.count)
        Write-Host "Setting configurations to those in scope..." -foregroundcolor Yellow        
        $ITGConfigurations = $ITGConfigurations | Where-Object { $ScopedITGCompanyIds -contains $_.attributes.'organization-id' }
        Write-Host "configurations scoped... $OriginalConfigurationCount => $($ITGConfigurations.count)"
    }

    $NinjaOneDevices = Invoke-NinjaOneRequest -Method GET -Path 'devices'
    $NinjaOneLocations = Invoke-NinjaOneRequest -Method GET -Path 'locations'

    $MatchedDevices = foreach ($itgconfiguration in $ITGConfigurations ) {

        $NinjaOrgID = ($MatchedCompanies | Where-Object { $_.ITGID -eq $itgconfiguration.attributes."organization-id" }).NinjaOneID

        If ($null -eq $NinjaOrgID) {
            Write-Error "Failed to find NinjaOne Organization ID: $itgconfiguration"
            Continue
        }
        
        try {
            $NinjaOneLocationID = (($MatchedLocations | where-object -filter { ($_.ITGID -eq $itgconfiguration.attributes."location-id" -or $null -eq $itgconfiguration.attributes."location-id") -and $_.NinjaOneOrgID -eq $NinjaOrgID }).NinjaOneID)[0]
        } catch {
            $NinjaOneLocationID = (($NinjaOneLocations | Where-Object { $_.organizationId -eq $NinjaOrgID }).id)[0]
        }

        If ($null -eq $NinjaOneLocationID) {
            Write-Error "Failed to find NinjaOne Location ID: $itgconfiguration"
            Continue
        }

        $NinjaOneDevice = $NinjaOneDevices | Where-Object { ($_.systemName -eq $itgconfiguration.attributes.name -or $_.displayName -eq $itgconfiguration.attributes.name) -and $_.organizationId -eq $NinjaOrgID }

        if ($NinjaOneDevice) {
            [PSCustomObject]@{
                "Name"           = $itgconfiguration.attributes.name
                "CompanyName"    = $itgconfiguration.attributes."organization-name"
                "ITGID"          = $itgconfiguration.id
                "NinjaOneID"     = $NinjaOneDevice.id
                "Matched"        = $true
                "NinjaOneObject" = $NinjaOneDevice
                "ITGObject"      = $itgconfiguration
                "Imported"       = "Pre-Existing"
                "NinjaOneOrgID"  = $NinjaOrgID
                "NinjaOneLocID"  = $NinjaOneDevice.locationId
					
            }
        } else {
            [PSCustomObject]@{
                "Name"           = $itgconfiguration.attributes.name
                "CompanyName"    = $itgconfiguration.attributes."organization-name"
                "ITGID"          = $itgconfiguration.id
                "NinjaOneID"     = ""
                "Matched"        = $false
                "NinjaOneObject" = ""
                "ITGObject"      = $itgconfiguration
                "Imported"       = ""
                "NinjaOneOrgID"  = $NinjaOrgID
                "NinjaOneLocID"  = $NinjaOneLocationID
            }
        }
    }

    Write-Host "Matched Devices (Already exist so will not be migrated)"
    $MatchedDevices | Sort-Object CompanyName, Name | Where-Object { $_.Matched -eq $true } | Select-Object CompanyName, Name, NinjaOneOrgID, NinjaOneLocID | Format-Table

    Write-Host "Unmatched Devices"
    $MatchedDevices | Sort-Object CompanyName, Name | Where-Object { $_.Matched -eq $false } | Select-Object CompanyName, Name, NinjaOneOrgID, NinjaOneLocID | Format-Table


    
   
    # Import Devices
    $UnmappedConfigurationssCount = ($MatchedDevices | Where-Object { $_.Matched -eq $false } | measure-object).count
    if ($ImportConfigurations -eq $true -and $UnmappedConfigurationssCount -gt 0) {        
        $ITGConfigTypes = ($MatchedDevices | Where-Object { $_.Matched -eq $false }).ITGObject.attributes."configuration-type-name" | Select-Object -unique

        $NinjaOneRoles = Invoke-NinjaOneRequest -Method GET -Path 'roles' | Where-Object { $_.nodeClass -eq 'UNMANAGED_DEVICE' -and $_.name -ne 'UNMANAGED_DEVICE' }

        if (($NinjaOneRoles | Measure-Object).count -lt 1) {
            Write-Error "No Unmanaged Device Roles were found. Devices cannot be imported without unmanaged roles."
            Read-Host "Press any key to continue or CTRL+C to quit"
            $importDeviceOption = 'NoRoles'
        } else {
            $importDeviceOption = Get-ImportMode -ImportName "Configurations"
        }

        $RoleMap = [PSCustomObject]@{}

        

        if (($importDeviceOption -eq "A") -or ($importDeviceOption -eq "S") ) {		

            foreach ($ConfigType in $ITGConfigTypes) {
                Write-Host ""
                Write-Host "Mapping Configuration type $ConfigType to NinjaOne Unmanaged Device Role"
                Write-Host "$($NinjaOneRoles | Select-Object id, name | Format-Table | Out-String)"
                $RoleFound = $False
                do {
                    $MapValue = Read-Host "Please enter the ID of the NinjaOne Unmanaged Device Role for the Configuration Type $ConfigType"
                    if ($MapValue -in $NinjaOneRoles.ID) {
                        $RoleFound = $True
                    } else {
                        Write-Error "Please enter a valid NinjaOne Unmanaged Device Role ID"
                    }
                } while ($RoleFound = $False)

                $RoleMap | Add-Member -Name $ConfigType -Value $MapValue -MemberType NoteProperty

            }
	
        
            foreach ($unmatcheddevice in ($MatchedDevices | Where-Object { $_.Matched -eq $false })) {
                Confirm-Import -ImportObjectName $unmatcheddevice.name -ImportObject $unmatcheddevice -ImportSetting $importDeviceOption
                
                Write-Host "Starting $($unmatcheddevice.Name)"
                
                $DeviceCreation = @{
                    name                = $unmatcheddevice.ITGObject.attributes."name"
                    orgId               = $unmatcheddevice.NinjaOneOrgID
                    locationId          = $unmatcheddevice.NinjaOneLocID
                    roleId              = $RoleMap."$($unmatcheddevice."ITGObject".attributes."configuration-type-name")"
                    'serialNumber'      = $unmatcheddevice."ITGObject".attributes."serial-number" ?? ""
                    'warrantyEndDate'   = $(try { (Get-NinjaOneTime -Seconds -Date "$($unmatcheddevice."ITGObject".attributes."warranty-expires-at")" -ea stop) }catch { "" })
                    'warrantyStartDate' = $(try { (Get-NinjaOneTime -Seconds -Date "$($unmatcheddevice."ITGObject".attributes."purchased-at")" -ea stop) }catch { "" })
                }
                
                $NewNinjaOneDevice = Invoke-NinjaOneRequest -Path "itam/unmanaged-device" -Method POST -InputObject $DeviceCreation
			
                $unmatcheddevice.matched = $true
                $unmatcheddevice.NinjaOneID = $NewNinjaOneDevice.id
                $unmatcheddevice.NinjaOneObject = $NewNinjaOneDevice
                $unmatcheddevice.Imported = "Created-By-Script"
			
                Write-host "$($unmatcheddevice.Name) Has been created in NinjaOne"
                Write-Host ""
            }
		
        }
		
    } else {
        if ($UnmappedConfigurationssCount -eq 0) {
            Write-Host "All Devices matched, no migration required" -foregroundcolor green
        } else {
            Write-Host "Warning Import Configurations is set to disabled so the above unmatched Configurations will not have data migrated" -foregroundcolor red
            Write-TimedMessage -Message "Press any key to continue or CTRL+C to quit" -DefaultResponse "continue and wrap-up configurations, please." -Timeout 6
        }
    }

    # Save the results to resume from if needed
    $MatchedDevices | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\Devices.json"
    Write-TimedMessage -Timeout 3 -Message "Snapshot Point: Configurations Migrated Continue?"  -DefaultResponse "continue to Contacts, please."

}

############################### Contacts ###############################
#Check for Location Resume
if ($ResumeFound -eq $true -and (Test-Path "$MigrationLogs\Contacts.json")) {
    Write-Host "Loading Previous Contacts Migration"
    $MatchedContacts = Get-Content "$MigrationLogs\Contacts.json" -raw | Out-String | ConvertFrom-Json -depth 100
} else {

    Write-Host "Fetching Contacts from IT Glue" -ForegroundColor Green
    $ContactsSelect = { (Get-ITGlueContacts -page_size 1000 -page_number $i -include related_items).data }
    $ITGContacts = Import-ITGlueItems -ItemSelect $ContactsSelect
    if ($ScopedMigration) {
        $OriginalContactsCount = $($ITGContacts.count)
        Write-Host "Setting contacts to those in scope..." -foregroundcolor Yellow               
        $ITGContacts = $ITGContacts | Where-Object { $ScopedITGCompanyIds -contains $_.attributes.'organization-id' }
        Write-Host "Contacts scoped... $OriginalContactsCount => $($ITGContacts.count)"
    }
    #($ITGContacts.attributes | sort-object -property name, "organization-name" -Unique)

    $NinjaOneUsers = Invoke-NinjaOneRequest -Method GET -Path 'users'

    $MatchedContacts = foreach ($itgcontact in $ITGContacts ) {

        $NinjaOrgID = ($MatchedCompanies | Where-Object { $_.ITGID -eq $itgconfiguration.attributes."organization-id" }).NinjaOneID

        If ($null -eq $NinjaOrgID) {
            Write-Error "Failed to find NinjaOne Organization ID: $itgconfiguration"
            Continue
        }

        $PrimaryEmail = ($itgcontact.attributes."contact-emails" | Where-Object { $_.primary -eq $True }).value

        if ($null -eq $PrimaryEmail) {
            Write-Error "$($itgcontact.attributes."organization-name") - $($itgcontact.attributes.name) Cannot be migrated as they have no email address in IT Glue"
            continue
        }

        $NinjaOneUser = $NinjaOneUsers | Where-Object { $_.email -eq $PrimaryEmail }

        
        if ($NinjaOneUser) {
            [PSCustomObject]@{
                "Name"           = $itgcontact.attributes.name
                "CompanyName"    = $itgcontact.attributes."organization-name"
                "ITGID"          = $itgcontact.id
                "NinjaOneID"     = $NinjaOneUser.id
                "Matched"        = $true
                "NinjaOneObject" = $NinjaOneUser
                "ITGObject"      = $itgcontact
                "Imported"       = "Pre-Existing"
                "NinjaOneOrgID"  = $NinjaOrgID
                "PrimaryEmail"   = $PrimaryEmail	
            }
        } else {
            [PSCustomObject]@{
                "Name"           = $itgcontact.attributes.name
                "CompanyName"    = $itgcontact.attributes."organization-name"
                "ITGID"          = $itgcontact.id
                "NinjaOneID"     = ""
                "Matched"        = $false
                "NinjaOneObject" = ""
                "ITGObject"      = $itgcontact
                "Imported"       = ""
                "NinjaOneOrgID"  = $NinjaOrgID
                "PrimaryEmail"   = $PrimaryEmail	
            }
        }
    }

    Write-Host "Matched Contacts (Already exist so will not be migrated)"
    $MatchedContacts | Sort-Object CompanyName, Name | Where-Object { $_.Matched -eq $true } | Select-Object CompanyName, Name, PrimaryEmail | Format-Table

    Write-Host "Unmatched Contacts"
    $MatchedContacts | Sort-Object CompanyName, Name | Where-Object { $_.Matched -eq $false } | Select-Object CompanyName, Name, PrimaryEmail | Format-Table


    # Import Contacts
    $UnmappedContactsCount = ($MatchedContacts | Where-Object { $_.Matched -eq $false } | measure-object).count
    if ($ImportContacts -eq $true -and $UnmappedContactsCount -gt 0) {        
        
        $importContactOption = Get-ImportMode -ImportName "contacts"

        if (($importContactOption -eq "A") -or ($importContactOption -eq "S") ) {		

      
            foreach ($unmatchedcontact in ($MatchedContacts | Where-Object { $_.Matched -eq $false })) {
                Confirm-Import -ImportObjectName $unmatchedcontact.name -ImportObject $unmatchedcontact -ImportSetting $importContactOption
                
                Write-Host "Starting $($unmatchedcontact.Name)"
                
                $ContactCreation = @{
                    firstName        = $unmatchedcontact."ITGObject".attributes."first-name" ?? 'N/A'
                    lastName         = $unmatchedcontact."ITGObject".attributes."last-name" ?? 'N/A'
                    email            = $unmatchedcontact.PrimaryEmail
                    phone            = ($unmatchedcontact."ITGObject".attributes."contact-phones" | Where-Object { $_.primary -eq $True }).value
                    organizationId   = $unmatchedcontact.NinjaOneOrgID
                    fullPortalAccess = $False
                    
                }
                
                try {
                    $NewNinjaOneContact = Invoke-NinjaOneRequest -Path "user/end-users" -Method POST -InputObject $ContactCreation -ea Stop

                    $unmatchedcontact.matched = $true
                    $unmatchedcontact.NinjaOneID = $NewNinjaOneContact.id
                    $unmatchedcontact.NinjaOneObject = $NewNinjaOneContact
                    $unmatchedcontact.Imported = "Created-By-Script"
			
                    Write-host "$($unmatchedcontact.Name) Has been created in NinjaOne"
                    Write-Host ""

                } catch {
                    Write-Error "Failed to create contact: $($unmatchedcontact."organization-name") - $($unmatchedcontact.PrimaryEmail)"
                    Write-Error "Email addresses must be unique across the NinjaOne platform. You can try adding a + address like test+example@test.com"
                    Write-Error "$_"
                    $unmatchedcontact.Imported = "Failed to create: $_"
                }
			
                
            }
		
        }
		
    } else {
        if ($UnmappedContactsCount -eq 0) {
            Write-Host "All Contacts matched, no migration required" -foregroundcolor green
        } else {
            Write-Host "Warning Import Contact is set to disabled so the above unmatched Contacts will not have data migrated" -foregroundcolor red
            Write-TimedMessage -Message "Press any key to continue or CTRL+C to quit" -DefaultResponse "continue and wrap-up contacts, please." -Timeout 6
        }
    }
    
    Write-Host "Contacts Complete"

    # Save the results to resume from if needed
    $MatchedContacts | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\Contacts.json"
    Write-TimedMessage -Timeout 3 -Message "Snapshot Point: Contacts Migrated Continue?"  -DefaultResponse "continue to Users, please."

}

############################### Users ###############################
#Check for Location Resume
if ($ResumeFound -eq $true -and (Test-Path "$MigrationLogs\Users.json")) {
    Write-Host "Loading Previous Users Migration"
    $MatchedContacts = Get-Content "$MigrationLogs\Users.json" -raw | Out-String | ConvertFrom-Json -depth 100
} else {

    Write-Host "Fetching users from IT Glue" -ForegroundColor Green
    $UsersSelect = { (Get-ITGlueUsers -page_size 1000 -page_number $i).data }
    $ITGUsers = Import-ITGlueItems -ItemSelect $UsersSelect

    $NinjaOneUsers = Invoke-NinjaOneRequest -Method GET -Path 'users'

    $MatchedUsers = foreach ($itguser in $ITGUsers ) {

        $NinjaOneUser = $NinjaOneUsers | Where-Object { $_.email -eq $itguser.attributes.email }

        
        if ($NinjaOneUser) {
            [PSCustomObject]@{
                "Name"           = $itguser.attributes.name
                "ITGID"          = $itguser.id
                "NinjaOneID"     = $NinjaOneUser.id
                "Matched"        = $true
                "NinjaOneObject" = $NinjaOneUser
                "ITGObject"      = $itguser
                "Imported"       = "Pre-Existing"
                "PrimaryEmail"   = $itguser.attributes.email	
            }
        } else {
            [PSCustomObject]@{
                "Name"           = $itguser.attributes.name
                "ITGID"          = $itguser.id
                "NinjaOneID"     = ""
                "Matched"        = $false
                "NinjaOneObject" = ""
                "ITGObject"      = $itguser
                "Imported"       = ""
                "PrimaryEmail"   = $itguser.attributes.email	
            }
        }
    }

    Write-Host "Matched Users"
    $MatchedUsers | Sort-Object Name | Where-Object { $_.Matched -eq $true } | Select-Object Name, PrimaryEmail | Format-Table

    Write-Host "Unmatched Users"
    $UnmatchedUsers = $MatchedUsers | Sort-Object Name | Where-Object { $_.Matched -eq $false } | Select-Object Name, PrimaryEmail | Format-Table
    $UnmatchedUsers

    if (($UnmatchedUsers | Measure-Object).count -gt 0) {
        Write-Error "Any unmatched Users will not be migrated. They should be manually created as technicians in NinjaOne if you wish them to set in relations"
        Write-TimedMessage -Message "Press any key to continue or CTRL+C to quit" -DefaultResponse "continue and wrap-up contacts, please." -Timeout 6
    }
   

    # Save the results to resume from if needed
    $MatchedUsers | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\Users.json"
    Write-TimedMessage -Timeout 3 -Message "Snapshot Point: Users Matched Continue?"  -DefaultResponse "continue to Flexible Asset Layouts, please."

}
	
############################### Flexible Asset Layouts and Assets ###############################
#Check for Layouts Resume
if ($ResumeFound -eq $true -and (Test-Path "$MigrationLogs\AssetLayouts.json")) {
    Write-Host "Loading Previous Asset Layouts Migration"
    $MatchedLayouts = Get-Content "$MigrationLogs\AssetLayouts.json" -raw | Out-String | ConvertFrom-Json -depth 100
    $AllFields = Get-Content "$MigrationLogs\AssetLayoutsFields.json" -raw | Out-String | ConvertFrom-Json -depth 100
} else {

    Write-Host "Fetching Flexible Asset Layouts from IT Glue" -ForegroundColor Green
    $FlexLayoutSelect = { (Get-ITGlueFlexibleAssetTypes -page_size 1000 -page_number $i -include related_items).data }
    $FlexLayouts = Import-ITGlueItems -ItemSelect $FlexLayoutSelect

    $NinjaOneLayouts = Invoke-NinjaOneRequest -Method GET -Path 'document-templates'

    Write-Host "The script will now migrate IT Glue Flexible Asset Layouts to NinjaOne"
    Write-Host "Please select the option you would like"
    Write-Host "1) Move all Flexible Asset Layouts and their assets to NinjaOne"
    Write-Host "2) Determine on a layout by layout basis if you want to migrate"
    $ImportOption = Get-FlexLayoutImportMode

    $AllFields = [System.Collections.ArrayList]@()

    # Match to existing layouts
    $MatchedLayouts = foreach ($ITGLayout in $FlexLayouts) {
        $NinjaOneLayout = $NinjaOneLayouts | where-object -filter { $_.name -eq "$($FlexibleLayoutPrefix)$($ITGLayout.attributes.name)" }
		
        if ($NinjaOneLayout) {
            [PSCustomObject]@{
                "Name"           = $ITGLayout.attributes.name
                "ITGID"          = $ITGLayout.id
                "NinjaOneID"     = $NinjaOneLayout.id
                "Matched"        = $true
                "NinjaOneObject" = $NinjaOneLayout
                "ITGObject"      = $ITGLayout
                "ITGAssets"      = ""
                "Imported"       = "Pre-Existing"
                "Import"         = $true
			
            }
        } else {
            [PSCustomObject]@{
                "Name"           = $ITGLayout.attributes.name
                "ITGID"          = $ITGLayout.id
                "NinjaOneID"     = ""
                "Matched"        = $false
                "NinjaOneObject" = ""
                "ITGObject"      = $ITGLayout
                "ITGAssets"      = ""
                "Imported"       = ""
                "Import"         = $true
            }
        }
    }

    if ($ImportOption -eq 2) {
        foreach ($Layout in $MatchedLayouts) {
            $Import = ''
            do {
                $Import = Read-Host "Would you like to import $($Layout.name) and it's assets Y/N"
                if ($Import -eq 'Y') {
                    $Layout.Import = $true
                } elseif ($Import -eq 'N') {
                    $Layout.Import = $false
                } else {
                    Write-Error 'Please enter Y or N'
                }
            } while ($Import -notin @('Y', 'N') )

        }
    }

    Write-Host "Existing Templates, Fields will be updated if required and assets migrated"
    $MatchedLayouts | Sort-Object Name | Where-Object { $_.Matched -eq $true -and $_.Import -eq $True } | Select-Object Name | Format-Table

    Write-Host "Unmatched, will be created and assets migrated"
    $MatchedLayouts | Sort-Object Name | Where-Object { $_.Matched -eq $false -and $_.Import -eq $True } | Select-Object Name | Format-Table

    Read-Host 'Press any key to continue or Ctrl+C to cancel'


    if ($ImportFlexibleAssetLayouts -eq $true) {

        foreach ($ImportLayout in $MatchedLayouts | Where-Object { $_.Matched -eq $false -and $_.import -eq $true }) {
            
            $Template = @{
                name                      = "$($FlexibleLayoutPrefix)$($ImportLayout.ITGObject.attributes.name)"
                description               = "$($FlexibleLayoutPrefix)$($ImportLayout.ITGObject.attributes.description)"
                allowMultiple             = $true
                mandatory                 = $false
                availableToAllTechnicians = $true
                fields                    = @(
                    @{
                        fieldLabel                = 'ITGlue Import Date'
                        fieldName                 = 'itgImportDate'
                        fieldDescription          = 'The date this asset was imported from IT Glue'
                        fieldType                 = 'DATE'
                        fieldTechnicianPermission = 'EDITABLE'
                        fieldScriptPermission     = 'NONE'
                        fieldApiPermission        = 'READ_WRITE'
                        fieldContent              = @{
                            required = $false
                        }
                    },
                    @{
                        fieldLabel                = 'ITGlue URL'
                        fieldName                 = 'itgUrl'
                        fieldDescription          = 'The URL to the original item in ITGlue'
                        fieldType                 = 'URL'
                        fieldTechnicianPermission = 'EDITABLE'
                        fieldScriptPermission     = 'NONE'
                        fieldApiPermission        = 'READ_WRITE'
                        fieldContent              = @{
                            required = $false
                        }
                    }
                )
            }
            
            $NewTemplate = Invoke-NinjaOneDocumentTemplate -Template $Template
            $ImportLayout.NinjaOneObject = $NewTemplate
            $ImportLayout.NinjaOneID = $NewTemplate.id
            $ImportLayout.Imported = "Created-By-Script"

        }

        foreach ($UpdateLayout in $MatchedLayouts | Where-Object { $_.Import -eq $true }) {
            Write-Host "Starting $($UpdateLayout.Name)" -ForegroundColor Green

            # Grab the fields for the layout
            Write-Host "Fetching Flexible Asset Fields from IT Glue"
            $FlexLayoutFieldsSelect = { (Get-ITGlueFlexibleAssetFields -page_size 1000 -page_number $i -flexible_asset_type_id $UpdateLayout.ITGID).data }
            $FlexLayoutFields = Import-ITGlueItems -ItemSelect $FlexLayoutFieldsSelect

				
            # Grab all the Assets for the layout
            Write-Host "Fetching Flexible Assets from IT Glue (This may take a while)"
            $FlexAssetsSelect = { (Get-ITGlueFlexibleAssets -page_size 1000 -page_number $i -filter_flexible_asset_type_id $UpdateLayout.ITGID -include related_items).data }
            $FlexAssets = Import-ITGlueItems -ItemSelect $FlexAssetsSelect
		
            
				
            [System.Collections.Generic.List[PSCustomObject]]$UpdateTemplateFields = @()
            foreach ($ITGField in $FlexLayoutFields | Sort-Object $_.attributes.order) {
                if ($ITGField.Attributes.kind -eq 'Header') {
                    $TemplateField = @{
                        uiElementUid   = (New-Guid).Guid
                        uiElementName  = $ITGField.Attributes.name
                        uiElementType  = 'TITLE'
                        uiElementValue = $ITGField.Attributes.name
                    }
                } else {

                    $supported = $true
		
                    switch ($ITGField.Attributes.kind) {
                        "Checkbox" {
                            $NinjaOneFieldType = 'CHECKBOX'
                            $NinjaOneFieldContent = @{
                                required    = $ITGField.Attributes.required
                                tooltipText = $ITGField.Attributes.hint ?? ''
                            }
                        }
                        "Date" {
                            $NinjaOneFieldType = 'DATE'
                            $NinjaOneFieldContent = @{
                                required    = $ITGField.Attributes.required
                                tooltipText = $ITGField.Attributes.hint ?? ''
                            }
                        }
                        "Number" {
                            if ($ITGField.Attributes.decimals -gt 0) {
                                $NinjaOneFieldType = 'DECIMAL'
                            } else {
                                $NinjaOneFieldType = 'NUMERIC'
                            }
                            $NinjaOneFieldContent = @{
                                required    = $ITGField.Attributes.required
                                tooltipText = $ITGField.Attributes.hint ?? ''
                            }
                        }
                        "Select" {
                            $NinjaOneFieldType = 'DROPDOWN'
                            $NinjaOneValues = (($ITGField.Attributes."default-value") -split "`n" | ForEach-Object {
                                    @{name = $_ }
                                }) ?? @()
                            $NinjaOneFieldContent = @{
                                values      = $NinjaOneValues
                                required    = $ITGField.Attributes.required
                                tooltipText = $ITGField.Attributes.hint ?? ''
                            }
                        }
                        "Text" {
                            $NinjaOneFieldType = 'TEXT'
                            $NinjaOneFieldContent = @{
                                required    = $ITGField.Attributes.required
                                tooltipText = $ITGField.Attributes.hint ?? ''
                            }
                        }
                        "Textbox" {
                            $NinjaOneFieldType = 'WYSIWYG'
                            $NinjaOneFieldContent = @{
                                required         = $ITGField.Attributes.required
                                tooltipText      = $ITGField.Attributes.hint ?? ''
                                advancedSettings = @{
                                    expandLargeValueOnRender = $true
                                }
                            }
                        }
                        "Upload" {
                            $NinjaOneFieldType = 'ATTACHMENT'
                            $NinjaOneFieldContent = @{
                                required    = $ITGField.Attributes.required
                                tooltipText = $ITGField.Attributes.hint ?? ''
                            }
                        }
                        "Tag" {
                            switch (($ITGField.Attributes."tag-type").split(":")[0]) {
                                "AccountsUsers" { Write-Host "Tags to Account Users are not supported $($ITGField.Attributes.name) in $($UpdateLayout.name) will need to be manually migrated, Sorry!" ; $supported = $false }
                                "Checklists" { Write-Host "Tags to Checklists are not supported $($ITGField.Attributes.name) in $($UpdateLayout.name) will need to be manually migrated, Sorry!"; $supported = $false }
                                "ChecklistTemplates" { Write-Host "Tags to Checklists Templates are not supported $($ITGField.Attributes.name) in $($UpdateLayout.name) will need to be manually migrated, Sorry!"; $supported = $false }
                                "Contacts" { Write-Host "Tags to Contacts are not supported $($ITGField.Attributes.name) in $($UpdateLayout.name) will need to be manually migrated, Sorry!"; $supported = $false }
                                "Configurations" {
                                    $NinjaOneFieldType = 'NODE_MULTI_SELECT'
                                    $NinjaOneFieldContent = @{
                                        required    = $ITGField.Attributes.required
                                        tooltipText = $ITGField.Attributes.hint ?? ''
                                    }
                                }
                                "Documents" { Write-Host "Tags to Documents are not supported $($ITGField.Attributes.name) in $($UpdateLayout.name) will need to be manually migrated, Sorry!"; $supported = $false } 
                                "Domains" { Write-Host "Tags to Websites are not supported $($ITGField.Attributes.name) in $($UpdateLayout.name) will need to be manually migrated, Sorry!"; $supported = $false }
                                "Passwords" { Write-Host "Tags to Passwords are not supported $($ITGField.Attributes.name) in $($UpdateLayout.name) will need to be manually migrated, Sorry!"; $supported = $false }
                                "Locations" {
                                    $NinjaOneFieldType = 'CLIENT_LOCATION_MULTI_SELECT'
                                    $NinjaOneFieldContent = @{
                                        required    = $ITGField.Attributes.required
                                        tooltipText = $ITGField.Attributes.hint ?? ''
                                    }
                                }
                                "Organizations" { 
                                    $NinjaOneFieldType = 'CLIENT_MULTI_SELECT'
                                    $NinjaOneFieldContent = @{
                                        required    = $ITGField.Attributes.required
                                        tooltipText = $ITGField.Attributes.hint ?? ''
                                    }
                                }
                                "SslCertificates" { Write-Host "Tags to SSL Certificates are not supported $($ITGField.Attributes.name) in $($UpdateLayout.name) will need to be manually migrated, Sorry!"; $supported = $false }
                                "Tickets" { Write-Host "Tags to Tickets are not supported $($ITGField.Attributes.name) in $($UpdateLayout.name) will need to be manually migrated, Sorry!"; $supported = $false }
                                "FlexibleAssetType" { Write-Host "Tags to Assets are not supported $($ITGField.Attributes.name) in $($UpdateLayout.name) will need to be manually migrated, Sorry!"; $supported = $false }
                            }
                        }
                        "Percent" {
                            $NinjaOneFieldType = 'NUMERIC'
                            $NinjaOneFieldContent = @{
                                required    = $ITGField.Attributes.required
                                tooltipText = $ITGField.Attributes.hint ?? ''
                            }
                        }
                        "Password" {
                            $NinjaOneFieldType = 'TEXT_ENCRYPTED'
                            $NinjaOneFieldContent = @{
                                required         = $ITGField.Attributes.required
                                tooltipText      = $ITGField.Attributes.hint ?? ''
                                advancedSettings = @{
                                    maxCharacters = 20000
                                }
                                
                            }
                        }
                    }

                    $TemplateField = @{
                        fieldLabel                = $ITGField.Attributes.name
                        fieldName                 = (ConvertTo-CamelCase -InputString $ITGField.Attributes.name)
                        fieldType                 = $NinjaOneFieldType
                        fieldTechnicianPermission = 'EDITABLE'
                        fieldScriptPermission     = 'NONE'
                        fieldApiPermission        = 'READ_WRITE'
                        fieldDefaultValue         = $ITGField.Attributes.'default-value' ?? ""
                        fieldContent              = $NinjaOneFieldContent
                    }

                }

                if ($ITGField.Attributes.kind -eq "Tag") {
                    $SubKind = ($ITGField.Attributes."tag-type").split(":")[0]
                } else {
                    $SubKind = ""
                }  

                $FieldDetails = [PSCustomObject]@{
                    LayoutName         = $UpdateLayout.Name
                    FieldName          = $ITGField.Attributes.name
                    FieldType          = $ITGField.Attributes.kind
                    FieldSubType       = $SubKind
                    NinjaOneLayoutID   = $UpdateLayout.NinjaOneID
                    IGLayoutID         = $UpdateLayout.ITGID
                    ITGParsedName      = $ITGField.Attributes."name-key"
                    NinjaOneParsedName = (ConvertTo-CamelCase -InputString $ITGField.Attributes.name)
                    Supported          = $supported
                    NinjaOneField      = $TemplateField
                }
                $null = $AllFields.add($FieldDetails)


                if ($supported -eq $true) {
                    $UpdateTemplateFields.add($TemplateField)
                }

            }

            $UpdateTemplateFields.add(@{
                    fieldLabel                = 'ITGlue Import Date'
                    fieldName                 = 'itgImportDate'
                    fieldDescription          = 'The date this asset was imported from IT Glue'
                    fieldType                 = 'DATE'
                    fieldTechnicianPermission = 'EDITABLE'
                    fieldScriptPermission     = 'READ_ONLY'
                    fieldApiPermission        = 'READ_WRITE'
                    fieldContent              = @{
                        required = $false
                    }
                })

            $UpdateTemplateFields.add(@{
                    fieldLabel                = 'ITGlue URL'
                    fieldName                 = 'itgUrl'
                    fieldDescription          = 'The URL to the original item in ITGlue'
                    fieldType                 = 'URL'
                    fieldTechnicianPermission = 'EDITABLE'
                    fieldScriptPermission     = 'READ_ONLY'
                    fieldApiPermission        = 'READ_WRITE'
                    fieldContent              = @{
                        required = $false
                    }
                })


            $UpdateTemplate = @{
                name          = $UpdateLayout.NinjaOneObject.Name
                description   = $UpdateLayout.ITGObject.attributes.description
                allowMultiple = $true
                mandatory     = $false
                fields        = $UpdateTemplateFields
            }

            $UpdatedLayout = Invoke-NinjaOneDocumentTemplate -Template $UpdateTemplate
            Write-Host "Finished $($UpdateLayout.NinjaOneObject.Name)"
            $UpdateLayout.NinjaOneObject = $UpdatedLayout
            $UpdateLayout.ITGAssets = $FlexAssets
            $UpdateLayout.Matched = $true

        }

    }


    $AllFields | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\AssetLayoutsFields.json"
    $MatchedLayouts | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\AssetLayouts.json"
    Write-TimedMessage -Timeout 3 -Message "Snapshot Point: Layouts Migrated Continue?"  -DefaultResponse "continue to Flexible Assets, please."

}



############################### Flexible Assets ###############################
#Check for Assets Resume
if ($ResumeFound -eq $true -and (Test-Path "$MigrationLogs\Assets.json")) {
    Write-Host "Loading Previous Asset Migration"
    $MatchedAssets = Get-Content "$MigrationLogs\Assets.json" -raw | Out-String | ConvertFrom-Json -depth 100
    $MatchedAssetPasswords = Get-Content "$MigrationLogs\AssetPasswords.json" -raw | Out-String | ConvertFrom-Json -depth 100
    $RelationsToCreate = [System.Collections.ArrayList](Get-Content "$MigrationLogs\RelationsToCreate.json" -raw | Out-String | ConvertFrom-Json -depth 100)
    $ManualActions = [System.Collections.ArrayList](Get-Content "$MigrationLogs\ManualActions.json" -raw | Out-String | ConvertFrom-Json -depth 100)
} else {

    

    # Load raw passwords for embedded fields and future use
    $ITGPasswordsRaw = Import-CSV -Path "$ITGLueExportPath\passwords.csv"
    if ($ImportFlexibleAssets -eq $true) {
        [System.Collections.Generic.List[PSCustomObject]]$RelationsToCreate = @()
        [System.Collections.Generic.List[PSCustomObject]]$MatchedAssets = @()
        [System.Collections.Generic.List[PSCustomObject]]$MatchedAssetPasswords = @()
        

        #We need to do a first pass creating empty assets with just the ITG migrated data. This builds an array we need to use to lookup relations when populating the entire assets
        if ($ScopedMigration) {
            $OriginalLayoutsCount = $($MatchedLayouts.count)
            Write-Host "Setting layouts to those in scope..." -foregroundcolor Yellow               
            $MatchedLayouts = Filter-ScopedAssets -Layouts $MatchedLayouts -ScopedCompanyIds $ScopedITGCompanyIds
            Write-Host "Layouts scoped... $OriginalLayoutsCount => $($MatchedLayouts.count)"
        }

        Foreach ($Layout in $MatchedLayouts) {
            [System.Collections.Generic.List[PSCustomObject]]$DocumentsToCreate = @()
            $CurrentDocuments = Invoke-NinjaOneRequest -Method GET -Path 'organization/documents' -QueryParams "templateIds=$($Layout.NinjaOneObject.id)"

            Write-Host "Creating base assets for $($layout.name)" -ForegroundColor "Green"
            foreach ($ITGAsset in $Layout.ITGAssets) {
                # Match Company
                $NinjaCompanyID = ($MatchedCompanies | where-object -filter { $_.ITGID -eq $ITGAsset.attributes.'organization-id' }).NinjaOneID
                $MatchedDocument = $CurrentDocuments | Where-object { $_.documentName -eq $ITGAsset.attributes.name -and $_.organizationId -eq $NinjaCompanyID }

                if (($MatchedDocument | Measure-Object).count -eq 0) {
                    $DocCreation = @{ 
                        documentName       = $ITGAsset.attributes.name
                        documentTemplateId = $Layout.NinjaOneObject.id
                        organizationId     = $NinjaCompanyID
                        fields             = @{
                            'itgImportDate' = Get-NinjaOneTime -Date $(Get-Date)
                            'itgUrl'        = $ITGAsset.attributes.'resource-url'
                        }
                    }

                    $DocumentsToCreate.Add($DocCreation)
                } else {
                    Write-Host "Document already found, skipping $($ITGAsset.attributes.name)"
                }
            }
            

            for ($i = 0; $i -lt $DocumentsToCreate.Count; $i += 100) {
                $start = $i
                $end = [Math]::Min($i + 99, $DocumentsToCreate.Count - 1)
                $batch = @($DocumentsToCreate[$start..$end])
                try {
                    $null = Invoke-NinjaOneRequest -InputObject $Batch -Method POST -Path 'organization/documents' -AsArray -ea Stop
                } catch {
                    Write-Error "One or more items in the batch request failed $_"
                }
            }
        

            $CurrentDocuments = Invoke-NinjaOneRequest -Method GET -Path 'organization/documents' -QueryParams "templateIds=$($Layout.NinjaOneObject.id)"

            foreach ($ITGAsset in $Layout.ITGAssets) {
                $NinjaCompanyID = ($MatchedCompanies | where-object -filter { $_.ITGID -eq $ITGAsset.attributes.'organization-id' }).NinjaOneID                
                $MatchedDocument = $CurrentDocuments | Where-object { $_.fields.value -contains $ITGAsset.attributes.'resource-url' }
                if (($MatchedDocument | Measure-Object).count -eq 1) {
                    $AssetDetails = [PSCustomObject]@{
                        "Name"           = $ITGAsset.attributes.name
                        "ITGID"          = $ITGAsset.id
                        "NinjaOneID"     = $MatchedDocument.documentId
                        "Matched"        = $true
                        "NinjaOneObject" = $MatchedDocument
                        "ITGObject"      = $ITGAsset
                        "Imported"       = "First Pass"
                    }

                    $null = $MatchedAssets.add($AssetDetails)
                } else {
                    Write-Error "Failed to match $($ITGAsset.attributes.name) in $($ITGAsset.attributes.'organization-name'), creation may have failed due to duplicate name or other issue."
                }
                    
            }
		
        }
	
        #We now need to loop through all Assets again updating the assets to their final version
        [System.Collections.Generic.List[PSCustomObject]]$DocumentsToUpdate = @()
        foreach ($UpdateAsset in $MatchedAssets) {
            Write-Host "Populating $($UpdateAsset.Name)"
		
            $AssetFields = @{ 
                'itgImportDate' = Get-NinjaOneTime -Date $(Get-Date)
                'itgUrl'        = $UpdateAsset.ITGObject.attributes.'resource-url'
            }

            $traits = $UpdateAsset.ITGObject.attributes.traits
            $traits.PSObject.Properties | ForEach-Object {
                # Find the corresponding field we are working on
                $ITGParsed = $_.name
                $ITGValues = $_.value
                $field = $AllFields | where-object -filter { $_.IGLayoutID -eq $UpdateAsset.ITGObject.attributes.'flexible-asset-type-id' -and $_.ITGParsedName -eq $ITGParsed }
                if ($field) {
                    $supported = $true

                    if ($field.FieldType -eq "Tag") {
				
                        switch ($field.FieldSubType) {
                            "AccountsUsers" { Write-Host "Tags to Account Users are not supported $($field.FieldName) in $($UpdateAsset.Name) will need to be manually migrated, Sorry!"; $supported = $false }
                            "Checklists" { Write-Host "Tags to Checklists are not supported $($field.FieldName) in $($UpdateAsset.Name) will need to be manually migrated, Sorry!"; $supported = $false }
                            "ChecklistTemplates" { Write-Host "Tags to Checklists Templates are not supported $($field.FieldName) in $($UpdateAsset.Name) will need to be manually migrated, Sorry!"; $supported = $false }
                            "Contacts" { Write-Host "Tags to Contacts are not supported $($field.FieldName) in $($UpdateAsset.Name) will need to be manually migrated, Sorry!"; $supported = $false }
                            "Configurations" {
                                [System.Collections.Generic.List[int]]$ConfigsLinked = foreach ($IDMatch in $ITGValues.values) {
                                    $(($MatchedDevices | where-object -filter { $_.ITGID -eq $IDMatch.id }).NinjaOneID)
                                }
                                $ReturnData = @{
                                    entityIds = $ConfigsLinked
                                    type      = 'NODE'
                                }
                                $null = $AssetFields.add("$($field.NinjaOneParsedName)", ($ReturnData))
											
                            }
                            "Documents" { Write-Host "Tags to Documents are not supported $($field.FieldName) in $($UpdateAsset.Name) will need to be manually migrated, Sorry!"; $supported = $false }
                            "Domains" { Write-Host "Tags to Domains are not supported $($field.FieldName) in $($UpdateAsset.Name) will need to be manually migrated, Sorry!"; $supported = $false }
                            "Passwords" { Write-Host "Tags to Passwords are not supported $($field.FieldName) in $($UpdateAsset.Name) will need to be manually migrated, Sorry!"; $supported = $false }
                            "Locations" {
                                [System.Collections.Generic.List[int]]$LocationsLinked = foreach ($IDMatch in $ITGValues.values) {
                                    $(($MatchedLocations | where-object -filter { $_.ITGID -eq $IDMatch.id }).NinjaOneID)
                                }
                                $ReturnData = @{
                                    entityIds = $LocationsLinked
                                    type      = 'CLIENT_LOCATION'
                                }
                                $null = $AssetFields.add("$($field.NinjaOneParsedName)", ($ReturnData))
											
                            }
                            "Organizations" { 
                                [System.Collections.Generic.List[int]]$OrganizationsLinked = foreach ($IDMatch in $ITGValues.values) {
                                    $(($MatchedLocations | where-object -filter { $_.ITGID -eq $IDMatch.id }).NinjaOneID)
                                }
                                $ReturnData = @{
                                    entityIds = $OrganizationsLinked
                                    type      = 'CLIENT'
                                }
                                $null = $AssetFields.add("$($field.NinjaOneParsedName)", ($ReturnData))
                            }
                            "SslCertificates" { Write-Host "Tags to SSL Certificates are not supported $($field.FieldName) in $($UpdateAsset.Name) will need to be manually migrated, Sorry!"; $supported = $false }
                            "Tickets" { Write-Host "Tags to Tickets are not supported $($field.FieldName) in $($UpdateAsset.Name) will need to be manually migrated, Sorry!"; $supported = $false }
                            "FlexibleAssetType" { Write-Host "Tags to Flexable Assets are not supported $($field.FieldName) in $($UpdateAsset.Name) will need to be manually migrated, Sorry!"; $supported = $false }
                        }

                        if ($Supported -eq $False) {
                            $ManualLog = [PSCustomObject]@{
                                Document_Name     = $UpdateAsset.Name
                                Asset_Type        = $UpdateAsset.NinjaOneObject.documentTemplateName
                                Organization_Name = ($MatchedCompanies | Where-Object { $_.NinjaOneID -eq $UpdateAsset.NinjaOneObject.organizationId }).CompanyName
                                NinjaOneID        = $UpdateAsset.NinjaOneID
                                Field_Name        = $($field.FieldName)
                                Notes             = "Unsupported Tag Type Manual Tag Required"
                                Action            = "Manually tag to Asset"
                                Data              = $ITGValues.values.name -join ","
                                NinjaOne_URL      = "https://$($NinjaOneBaseDomain)/#/customerDashboard/$($UpdateAsset.NinjaOneObject.organizationId)/documentation/appsAndServices/$($UpdateAsset.NinjaOneObject.documentTemplateId)/$($UpdateAsset.NinajOneID)"
                                ITG_URL           = $UpdateAsset.ITGObject.attributes."resource-url"
                            }
                            $null = $ManualActions.add($ManualLog)
                        }

                    } elseif ($field.FieldType -eq "Upload") {
                        $SupportedFiles = $ITGValues | Where-Object { ($_.name -split '\.')[-1] -in $NinjaOneSupportedUploadTypes }
                        $UnsupportedFiles = $ITGValues | Where-Object { ($_.name -split '\.')[-1] -notin $NinjaOneSupportedUploadTypes }

                        $SupportedFound = $False
                        foreach ($SupportedFile in $SupportedFiles) {

                            if ($SupportedFound -eq $False) {
                                $AttachedFile = Get-ChildItem -Path $ITGLueExportPath -Filter "$(($SupportedFile.url -split "/")[-1])-$($SupportedFile.name)" -Recurse
                                if (($AttachedFile | Measure-Object).count -eq 1) {
                                    try {
                                        $FileResponse = Invoke-UploadNinjaOneFile -FileName $SupportedFile.name -FilePath $AttachedFile.VersionInfo.FileName -ContentType $SupportedFile.'content-type' -EntityType 'DOCUMENT' -ea stop
                                        $null = $AssetFields.add("$($field.NinjaOneParsedName)", ($FileResponse))
                                        $SupportedFound = $True

                                    } catch {
                                        Write-Error "$_"
                                        $ManualLog = [PSCustomObject]@{
                                            Document_Name     = $UpdateAsset.Name
                                            Asset_Type        = $UpdateAsset.NinjaOneObject.documentTemplateName
                                            Organization_Name = ($MatchedCompanies | Where-Object { $_.NinjaOneID -eq $UpdateAsset.NinjaOneObject.organizationId }).CompanyName
                                            NinjaOneID        = $UpdateAsset.NinjaOneID
                                            Field_Name        = $($field.FieldName)
                                            Notes             = "Failed uploading file"
                                            Action            = "Upload of $(($SupportedFile.url -split "/")[-1])-$($SupportedFile.name) Failed"
                                            Data              = $SupportedFile.name
                                            NinjaOne_URL      = "https://$($NinjaOneBaseDomain)/#/customerDashboard/$($UpdateAsset.NinjaOneObject.organizationId)/documentation/appsAndServices/$($UpdateAsset.NinjaOneObject.documentTemplateId)/$($UpdateAsset.NinajOneID)"
                                            ITG_URL           = $UpdateAsset.ITGObject.attributes."resource-url"
                                        }
                                        $null = $ManualActions.add($ManualLog)
                                    }
                                        
                                } else {
                                    Write-Error "Could not find $(($SupportedFile.url -split "/")[-1])-$($SupportedFile.name) in export or multiple files matched"
                                    $ManualLog = [PSCustomObject]@{
                                        Document_Name     = $UpdateAsset.Name
                                        Asset_Type        = $UpdateAsset.NinjaOneObject.documentTemplateName
                                        Organization_Name = ($MatchedCompanies | Where-Object { $_.NinjaOneID -eq $UpdateAsset.NinjaOneObject.organizationId }).CompanyName
                                        NinjaOneID        = $UpdateAsset.NinjaOneID
                                        Field_Name        = $($field.FieldName)
                                        Notes             = "File not found in export"
                                        Action            = "Could not find $(($SupportedFile.url -split "/")[-1])-$($SupportedFile.name) in export or multiple files matched"
                                        Data              = $SupportedFile.name
                                        NinjaOne_URL      = "https://$($NinjaOneBaseDomain)/#/customerDashboard/$($UpdateAsset.NinjaOneObject.organizationId)/documentation/appsAndServices/$($UpdateAsset.NinjaOneObject.documentTemplateId)/$($UpdateAsset.NinajOneID)"
                                        ITG_URL           = $UpdateAsset.ITGObject.attributes."resource-url"
                                    }
                                    $null = $ManualActions.add($ManualLog)
                                }
                                
                                    
                            } else {
                                $ManualLog = [PSCustomObject]@{
                                    Document_Name     = $UpdateAsset.Name
                                    Asset_Type        = $UpdateAsset.NinjaOneObject.documentTemplateName
                                    Organization_Name = ($MatchedCompanies | Where-Object { $_.NinjaOneID -eq $UpdateAsset.NinjaOneObject.organizationId }).CompanyName
                                    NinjaOneID        = $UpdateAsset.NinjaOneID
                                    Field_Name        = $($field.FieldName)
                                    Notes             = "Multiple supported files found"
                                    Action            = "Manual upload to related items of $(($SupportedFile.url -split "/")[-1])-$($SupportedFile.name) required"
                                    Data              = $SupportedFile.name
                                    NinjaOne_URL      = "https://$($NinjaOneBaseDomain)/#/customerDashboard/$($UpdateAsset.NinjaOneObject.organizationId)/documentation/appsAndServices/$($UpdateAsset.NinjaOneObject.documentTemplateId)/$($UpdateAsset.NinajOneID)"
                                    ITG_URL           = $UpdateAsset.ITGObject.attributes."resource-url"
                                }
                                $null = $ManualActions.add($ManualLog)
                            }

                        }

                        if (($UnsupportedFiles | Measure-Object).count -ge 1) {
                            $ManualLog = [PSCustomObject]@{
                                Document_Name     = $UpdateAsset.Name
                                Asset_Type        = $UpdateAsset.NinjaOneObject.documentTemplateName
                                Organization_Name = ($MatchedCompanies | Where-Object { $_.NinjaOneID -eq $UpdateAsset.NinjaOneObject.organizationId }).CompanyName
                                NinjaOneID        = $UpdateAsset.NinjaOneID
                                Field_Name        = $($field.FieldName)
                                Notes             = "Unsupported file type"
                                Action            = "Zip file and manually upload to the field or related items."
                                Data              = $UnsupportedFiles.name -join ","
                                NinjaOne_URL      = "https://$($NinjaOneBaseDomain)/#/customerDashboard/$($UpdateAsset.NinjaOneObject.organizationId)/documentation/appsAndServices/$($UpdateAsset.NinjaOneObject.documentTemplateId)/$($UpdateAsset.NinajOneID)"
                                ITG_URL           = $UpdateAsset.ITGObject.attributes."resource-url"
                            }
                            $null = $ManualActions.add($ManualLog)
                        }


                    } elseif ($field.FieldType -eq "Password") {
                        $ITGPassword = (Get-ITGluePasswords -id $ITGValues -include related_items).data
                        $ITGPasswordValue = ($ITGPasswordsRaw | Where-Object { $_.id -eq $ITGPassword.id }).password
                        try {
                            if ($ITGPasswordValue) {
                                $null = $AssetFields.add("$($field.NinjaOneParsedName)", $ITGPasswordValue)
                                $MigratedPasswordStatus = "Into Asset"
                            }
                        } catch {
                            Write-Host "Error occured adding field, possible duplicate name" -ForegroundColor Red
                            $ManualLog = [PSCustomObject]@{
                                Document_Name     = $UpdateAsset.Name
                                Asset_Type        = $UpdateAsset.NinjaOneObject.documentTemplateName
                                Organization_Name = ($MatchedCompanies | Where-Object { $_.NinjaOneID -eq $UpdateAsset.NinjaOneObject.organizationId }).CompanyName
                                NinjaOneID        = $UpdateAsset.NinjaOneID
                                Field_Name        = $($field.FieldName)
                                Notes             = "Failed to add password to Asset"
                                Action            = "Manually add the password to the asset"
                                Data              = ($ITGPassword.attributes.'resource-url' -replace '[^\x09\x0A\x0D\x20-\xD7FF\xE000-\xFFFD\x10000\x10FFFF]')
                                NinjaOne_URL      = "https://$($NinjaOneBaseDomain)/#/customerDashboard/$($UpdateAsset.NinjaOneObject.organizationId)/documentation/appsAndServices/$($UpdateAsset.NinjaOneObject.documentTemplateId)/$($UpdateAsset.NinajOneID)"
                                ITG_URL           = $UpdateAsset.ITGObject.attributes."resource-url"
                            }
                            $null = $ManualActions.add($ManualLog)
                            $MigratedPasswordStatus = "Failed to add"
                        }

                        $MigratedPassword = [PSCustomObject]@{
                            "Name"       = $ITGPassword.attributes.name
                            "ITGID"      = $ITGPassword.id
                            "NinjaOneID" = $UpdateAsset.NinjaOneID
                            "Matched"    = $true
                            "ITGObject"  = $ITGPassword
                            "Imported"   = $MigratedPasswordStatus
                        }
                        $null = $MatchedAssetPasswords.add($MigratedPassword)

                    } elseif ($field.FieldType -eq "Date") {
                        $null = $AssetFields.add("$($field.NinjaOneParsedName)", (Get-NinjaOneTime -Date $(Get-Date($_.value))))

                    } elseif ($field.FieldType -eq "Textbox") {
                        $null = $AssetFields.add("$($field.NinjaOneParsedName)", (@{html = $_.value -replace '[^\x09\x0A\x0D\x20-\xD7FF\xE000-\xFFFD\x10000\x10FFFF]' }))

                    } elseif ($field.NinjaOneField.fieldType -eq 'NUMERIC') {
                        $null = $AssetFields.add("$($field.NinjaOneParsedName)", ([int]$_.value))

                    } else {
                        $null = $AssetFields.add("$($field.NinjaOneParsedName)", ($_.value -replace '[^\x09\x0A\x0D\x20-\xD7FF\xE000-\xFFFD\x10000\x10FFFF]'))

                    }
                              
                } else {
                    Write-Host "Warning $ITGParsed : $ITGValues Could not be added" -ForegroundColor Red
                }
            }

            $UpdateNinjaOneBody = @{
                documentId   = $UpdateAsset.NinjaOneID
                documentName = $UpdateAsset.NinjaOneObject.documentName
                fields       = $AssetFields
            }

            $DocumentsToUpdate.add($UpdateNinjaOneBody)

            $UpdateAsset.Imported = "Created-By-Script"
        }

        for ($i = 0; $i -lt $DocumentsToUpdate.Count; $i += 100) {
            $start = $i
            $end = [Math]::Min($i + 99, $DocumentsToUpdate.Count - 1)
            $batch = @($DocumentsToUpdate[$start..$end])

            try {
                $null = Invoke-NinjaOneRequest -InputObject $Batch -Method PATCH -Path 'organization/documents' -AsArray -ea Stop
            } catch {
                Write-Error "One or more items in the batch request failed $_"
            }

        }

        Read-Host 'End of apps and services pause'

        $MatchedAssets | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\Assets.json"
        $MatchedAssetPasswords | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\AssetPasswords.json"
        $ManualActions | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\ManualActions.json"
        $RelationsToCreate | ConvertTo-Json -Depth 20 | Out-File "$MigrationLogs\RelationsToCreate.json"
        Write-TimedMessage -Timeout 3 -Message "Snapshot Point: Assets Migrated Continue?" -DefaultResponse "continue to Documents/Articles, please."
    }
}

############################### Documents / Articles ###############################

#Check for Article Resume
if ($ResumeFound -eq $true -and (Test-Path "$MigrationLogs\ArticleBase.json")) {
    Write-Host "Loading Article Migration"
    $MatchedArticles = Get-Content "$MigrationLogs\ArticleBase.json" -raw | Out-String | ConvertFrom-Json -depth 100
} else {

    if ($ImportArticles -eq $true) {
        [System.Collections.Generic.List[PSCustomObject]]$AllKBArticles = @()

        [System.Collections.Generic.List[PSCustomObject]]$OrgNinjaKBArticles = Invoke-NinjaOneRequest -Method GET -Path 'knowledgebase/organization/articles'
        [System.Collections.Generic.List[PSCustomObject]]$GlobalNinjaKBArticles = Invoke-NinjaOneRequest -Method GET -Path 'knowledgebase/global/articles'

        $AllKBArticles = Invoke-AddToArray -InputArray $AllKBArticles -AppendItem $OrgNinjaKBArticles
        $AllKBArticles = Invoke-AddToArray -InputArray $AllKBArticles -AppendItem $GlobalNinjaKBArticles

        $ITGDocuments = Import-CSV -Path (Join-Path -path $ITGLueExportPath -ChildPath "documents.csv")
        [string]$ITGDocumentsPath = Join-Path -path $ITGLueExportPath -ChildPath "Documents"
        [string]$ITGAttachmentsPath = Join-Path -path $ITGLueExportPath -ChildPath "attachments"

        $files = Get-ChildItem -Path $ITGDocumentsPath -recurse
        $attachmentFiles = Get-ChildItem -Path $ITGAttachmentsPath -recurse

        [System.Collections.Generic.List[PSCustomObject]]$ArticlesToCreate = @()

        # First lets find each article in the file system and then create blank stubs for them all so we can match relations later
        $MatchedArticles = Foreach ($doc in $ITGDocuments) {
            Write-Host "Starting $($doc.name)" -ForegroundColor Green

            $dir = $files | Where-Object { $_.PSIsContainer -eq $true -and $_.Name -match $doc.locator }
            $RelativePath = ($dir.FullName).Substring($ITGDocumentsPath.Length)
            $folders = $RelativePath -split '\\'
            $FilenameFromFolder = ($folders[$folders.count - 1] -split ' ', 2)[1]
            $Filename = $FilenameFromFolder

            $company = $MatchedCompanies | where-object -filter { $_.CompanyName -eq $doc.organization }
            if (($company | Measure-Object).count -ne 1) {
                Write-Host "Company $($doc.organization) Not Found Please migrate $($doc.name) manually"
                continue
            }

            if (($folders | Measure-Object).count -gt 2) {
                $folders = ($folders[1..$($folders.count - 2)]) -join '|'
            } else {
                $folders = ''
            }

            # Handle uploading files rather than HTML based documents
            if (($Doc.name -split '\.')[-1] -in $NinjaOneSupportedKBTypes) {
                
                # Check if exists
                $NinjaArticle = $AllKBArticles | Where-Object { $_.isNinjaArticle -eq $false -and ($_.organizationId -eq $company.NinjaOneID -or ($company.InternalCompany -eq $true -and $_.organizationId -eq $null)) -and $_.path -like "*$($folders)" -and $_.name -eq [io.path]::GetFileNameWithoutExtension($doc.name) }
                if (($NinjaArticle | Measure-Object).count -ge 1) {
                    Write-Host "Existing Uploaded document found"
                    [PSCustomObject]@{
                        "Name"           = $doc.name
                        "Filename"       = $Filename
                        "Path"           = $($dir.Fullname)
                        "FullPath"       = "$($dir.Fullname)\$($filename).html"
                        "ITGID"          = $doc.id
                        "ITGLocator"     = $doc.locator
                        "NinjaOneID"     = $NinjaArticle.id
                        "NinjaOneObject" = $NinjaArticle
                        "Folders"        = $folders
                        "Imported"       = "Uploaded"
                        "Company"        = $company
                        "ArticleType"    = 'FileUpload'
                    }
                    continue
                }

                $dir = $attachmentFiles | Where-Object { $_.PSIsContainer -eq $true -and $_.Name -match $doc.id }
                $Path = "$($dir.Fullname)\$((($doc.name) -replace '\+', '_') -replace ' ','_')" 
                $pathtest = Test-Path -LiteralPath $Path
                if ($pathtest -eq $true) {
                    try {
                        if ($company.InternalCompany -eq $false) {
                            $UploadedFile = Invoke-UploadNinjaOneKBArticle -FileName $doc.name -FilePath $Path -FolderPath $Folders -OrganizationID $company.NinjaOneID -ea stop
                        } else {
                            $UploadedFile = Invoke-UploadNinjaOneKBArticle -FileName $doc.name -FilePath $Path -FolderPath $Folders -ea stop
                        }
                        [PSCustomObject]@{
                            "Name"           = $doc.name
                            "Filename"       = $Filename
                            "Path"           = $($dir.Fullname)
                            "FullPath"       = $Path
                            "ITGID"          = $doc.id
                            "ITGLocator"     = $doc.locator
                            "NinjaOneID"     = $UploadedFile.id
                            "NinjaOneObject" = $UploadedFile
                            "Folders"        = $folders
                            "Imported"       = "Uploaded"
                            "Company"        = $company
                            "ArticleType"    = 'FileUpload'
                        }                        
                    } catch {
                        Write-Error "File upload failed $($dir.Fullname)\$($filename): $_"
                        [PSCustomObject]@{
                            "Name"           = $doc.name
                            "Filename"       = $Filename
                            "Path"           = $($dir.Fullname)
                            "FullPath"       = "$($dir.Fullname)\$($doc.name)" -replace '\+', '_'
                            "ITGID"          = $doc.id
                            "ITGLocator"     = $doc.locator
                            "NinjaOneID"     = ''
                            "NinjaOneObject" = ''
                            "Folders"        = $folders
                            "Imported"       = "Failed"
                            "Company"        = $company
                            "ArticleType"    = 'FileUpload'
                        }                        
                    }
                
                } else {
                    Write-Error "File was not found to upload $($dir.Fullname)\$($filename)"
                    [PSCustomObject]@{
                        "Name"           = $doc.name
                        "Filename"       = $Filename
                        "Path"           = $($dir.Fullname)
                        "FullPath"       = "$($dir.Fullname)\$($filename)"
                        "ITGID"          = $doc.id
                        "ITGLocator"     = $doc.locator
                        "NinjaOneID"     = ''
                        "NinjaOneObject" = ''
                        "Folders"        = $folders
                        "Imported"       = "NotFound"
                        "Company"        = $company
                        "ArticleType"    = 'FileUpload'
                    }
                }

                
            } else {
                # Process HTML Based Articles
                # Check if exists
                $NinjaArticle = $AllKBArticles | Where-Object { $_.content.html -eq $doc.id }
                if (($NinjaArticle | Measure-Object).count -ge 1) {
                    Write-Host "Existing Stub Document Found skipping creation"
                    [PSCustomObject]@{
                        "Name"           = $doc.name
                        "Filename"       = $Filename
                        "Path"           = $($dir.Fullname)
                        "FullPath"       = "$($dir.Fullname)\$($filename).html"
                        "ITGID"          = $doc.id
                        "ITGLocator"     = $doc.locator
                        "NinjaOneID"     = $NinjaArticle.id
                        "NinjaOneObject" = $NinjaArticle
                        "Folders"        = $folders
                        "Imported"       = "Stub-Created"
                        "Company"        = $company
                        "ArticleType"    = 'HTML'
                    }
                    continue
                }

                $pathtest = Test-Path -LiteralPath "$($dir.Fullname)\$($filename).html"

                if ($pathtest -eq $false) {
                    $filename = $doc.name
                    $pathtest = Test-Path -LiteralPath "$($dir.Fullname)\$($filename).html"
                    if ($pathtest -eq $false) {
                        $filename = $FilenameFromFolder -replace '_', '$1,$2'
                        $pathtest = Test-Path -LiteralPath "$($dir.Fullname)\$($filename).html"
                        if ($pathtest -eq $false) {
                            Write-Host "Not Found $($dir.Fullname)\$($filename).html this article will need to be migrated manually" -foregroundcolor red
                            continue
                        }
                    }
	
                }
                
                if ($company.InternalCompany -eq $false) {
                    $ArticleCreate = @{
                        name                  = $doc.name
                        organizationId        = $company.NinjaOneID
                        destinationFolderPath = $folders
                        content               = @{
                            html = "$($Doc.id)"
                        }
                    }

                } else {
                    # Set sub folder if set for global KB
                    if ($GlobalKBFolder) {
                        if ($folders -ne '') {
                            $folders = "$($GlobalKBFolder)|" + $folders
                        } else {
                            $folders = $GlobalKBFolder
                        }
                    }

                    $ArticleCreate = @{
                        name                  = $doc.name
                        destinationFolderPath = $folders
                        content               = @{
                            html = "$($Doc.id)"
                        }
                    }
                }
		
                $null = $ArticlesToCreate.add($ArticleCreate)

                [PSCustomObject]@{
                    "Name"           = $doc.name
                    "Filename"       = $Filename
                    "Path"           = $($dir.Fullname)
                    "FullPath"       = "$($dir.Fullname)\$($filename).html"
                    "ITGID"          = $doc.id
                    "ITGLocator"     = $doc.locator
                    "NinjaOneID"     = ''
                    "NinjaOneObject" = ''
                    "Folders"        = $folders
                    "Imported"       = "Batched"
                    "Company"        = $company
                    "ArticleType"    = 'HTML'
                }
            }

        }
        
        # Create stub articles in batches of 100
        for ($i = 0; $i -lt $ArticlesToCreate.Count; $i += 100) {
            $start = $i
            $end = [Math]::Min($i + 99, $ArticlesToCreate.Count - 1)
            $batch = @($ArticlesToCreate[$start..$end])

            try {
                $null = Invoke-NinjaOneRequest -InputObject $Batch -Method POST -Path 'knowledgebase/articles' -AsArray -ea Stop
            } catch {
                Write-Error "One or more items in the batch request failed $_"
            }

        }

        # Retrieve all articles and match to what we expect
        [System.Collections.Generic.List[PSCustomObject]]$AllKBArticles = @()
        [System.Collections.Generic.List[PSCustomObject]]$OrgNinjaKBArticles = Invoke-NinjaOneRequest -Method GET -Path 'knowledgebase/organization/articles'
        [System.Collections.Generic.List[PSCustomObject]]$GlobalNinjaKBArticles = Invoke-NinjaOneRequest -Method GET -Path 'knowledgebase/global/articles'

        $AllKBArticles = Invoke-AddToArray -InputArray $AllKBArticles -AppendItem $OrgNinjaKBArticles
        $AllKBArticles = Invoke-AddToArray -InputArray $AllKBArticles -AppendItem $GlobalNinjaKBArticles


        foreach ($Article in $MatchedArticles | Where-Object { $_.ArticleType -eq 'HTML' }) {
            $NinjaArticle = $AllKBArticles | Where-Object { $_.content.html -eq $Article.ITGID -and $_.isNinjaArticle -eq $True }

            if (($NinjaArticle | Measure-Object).count -eq 1) {
                $Article.NinjaOneID = $NinjaArticle.id
                $Article.NinjaOneObject = $NinjaArticle
                $Article.Imported = 'Stub-Created'
            } else {
                Write-Error "Failed to create $($Article.ITGID) - $($Article.name)"
                $Article.Imported = 'Creation-Failed'
            }
        }

        $MatchedArticles | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\ArticleBase.json"
        $ManualActions | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\ManualActions.json"
        Write-TimedMessage -Timeout 3 -Message "Snapshot Point: Stub Articles Created Continue?"  -DefaultResponse "continue to Document/Article Bodies, please."
    }

}


############################### Documents / Articles Bodies ###############################

#Check for Articles Resume
if ($ResumeFound -eq $true -and (Test-Path "$MigrationLogs\Articles.json")) {
    Write-Host "Loading Article Content Migration"
    $MatchedArticles = Get-Content "$MigrationLogs\Articles.json" -raw | Out-String | ConvertFrom-Json -depth 100
} else {
	
    if ($ImportArticles -eq $true) {
        $Attachfiles = Get-ChildItem (Join-Path -Path $ITGLueExportPath -ChildPath "attachments\documents") -recurse

        [System.Collections.Generic.List[PSCustomObject]]$ArticlesToUpdate = @()

        # Now do the actual work of populating the content of articles
        $ArticleErrors = foreach ($Article in $MatchedArticles | Where-Object { $_.ArticleType -eq 'HTML' -and $_.Imported -eq 'Stub-Created' }) {

            $page_out = ''
            $imagePath = $null
	    
            # Check for attachments
            $attachdir = $Attachfiles | Where-Object { $_.PSIsContainer -eq $true -and $_.Name -match $Article.ITGID }
            if ($Attachdir) {
                # TODO - Loop folder and upload to related items.

                $InFile = ''
                $html = ''
                $rawsource = ''

                $ManualLog = [PSCustomObject]@{
                    Document_Name = $Article.Name
                    Asset_Type    = "Article"
                    Company_Name  = $Article.Company.company_name
                    NinjaOneID    = $Article.NinjaOneID
                    Field_Name    = "N/A"
                    Notes         = "Attached Files not Supported"
                    Action        = "Manually Upload files to Related Files"
                    Data          = $attachdir.fullname
                    NinjaOne_URL  = "https://$($NinjaOneBaseDomain)/#/systemDashboard/knowledgeBase/$($Article.NinjaOneID)/file"
                    ITG_URL       = "$ITGURL/$($Article.ITGLocator)"
                }
                $null = $ManualActions.add($ManualLog)

            }


            Write-Host "Starting $($Article.Name) in $($Article.Company.CompanyName)" -ForegroundColor Green
				
            $InFile = $Article.FullPath
				
            $html = New-Object -ComObject "HTMLFile"
            $rawsource = Get-Content -encoding UTF8 -LiteralPath $InFile -Raw
            if ($rawsource.Length -gt 0) {
                $source = [regex]::replace($rawsource , '\xa0+', ' ')
                $src = [System.Text.Encoding]::Unicode.GetBytes($source)
                $html.write($src)
                $images = @($html.Images)

                $images | ForEach-Object {
                    $fullImgPath = $null
                    
                    if (($_.src -notmatch '^http[s]?://') -or ($_.src -match [regex]::Escape($ITGURL))) {
                        $script:HasImages = $true
                        $imgHTML = $_.outerHTML
                        Write-Host "Processing HTML: $imgHTML"
                        if ($_.src -match [regex]::Escape($ITGURL)) {
                            $matchedImage = Update-StringWithCaptureGroups -inputString $imgHTML -type 'img' -pattern $ImgRegexPatternToMatch
                            if ($matchedImage) {
                                Write-Host 'Matched by regex'
                                $tnImgUrl = $matchedImage.url
                                $tnImgPath = $matchedImage.path
                            } else {
                                $tnImgPath = $_.src
                            }
                        } else {
                            $basepath = Split-Path $InFile
                            if ($fullImgUrl = $imgHTML.split('data-src-original="')[1]) { $fullImgUrl = $fullImgUrl.split('"')[0] }
                            $tnImgUrl = $imgHTML.split('src="')[1].split('"')[0]
                            if ($fullImgUrl) { $fullImgPath = Join-Path -Path $basepath -ChildPath $fullImgUrl.replace('/', '\') }
                            $tnImgPath = Join-Path -Path $basepath -ChildPath $tnImgUrl.replace('/', '\')
                        }
                        
                        Write-Host "Processing IMG: $($fullImgPath ?? $tnImgPath)"
                        
                        # Some logic to test for the original data source being specified vs the thumbnail. Grab the Thumbnail or final source.
                        if ($fullImgUrl -and ($foundFile = Get-Item -Path "$fullImgPath*" -ErrorAction SilentlyContinue)) {
                            $imagePath = $foundFile.FullName
                        } elseif ($tnImgUrl -and ($foundFile = Get-Item -Path "$tnImgPath*" -ErrorAction SilentlyContinue)) {
                            $imagePath = $foundFile.FullName
                        } else { 
                            Remove-Variable -Name imagePath -ErrorAction SilentlyContinue
                            Remove-Variable -Name foundFile -ErrorAction SilentlyContinue
                            Write-Warning "Unable to validate image file."
                            $ManualLog = [PSCustomObject]@{
                                Document_Name = $Article.Name
                                Asset_Type    = "Article"
                                Company_Name  = $Article.Company.CompanyName
                                NinjaOneID    = $Article.NinjaOneID
                                Notes         = 'Missing image, file not found'
                                Actions       = "Neither $fullImgPath or $tnImgPath were found, validate the images exist in the export, or retrieve them from ITGlue directly"
                                Data          = "$InFile"
                                NinjaOne_URL  = "https://$($NinjaOneBaseDomain)/#/systemDashboard/knowledgeBase/$($Article.NinjaOneID)/file"
                                ITG_URL       = "$ITGURL/$($Article.ITGLocator)"
                            }

                            $null = $ManualActions.add($ManualLog)

                        }

                        # Test the path to ensure that a file extension exists, if no file extension we get problems later on. We rename it if there's no ext.
                        if ($imagePath -and (Test-Path $imagePath -ErrorAction SilentlyContinue)) {
                            if ((Get-Item -path $imagePath).extension -eq '') {
                                Write-Warning "$imagePath is undetermined image. Testing..."
                                if ($Magick = New-Object ImageMagick.MagickImage($imagePath)) {
                                    $OriginalFullImagePath = $imagePath
                                    $imagePath = "$($imagePath).$(($Magick.format).ToLower())"
                                    $MovedItem = Move-Item -Path $OriginalFullImagePath -Destination $imagePath
                                }
                            }                        
                            $imageType = Invoke-ImageTest($imagePath)
                            if ($imageType) {
                                Write-Host "Uploading new image $imagePath"
                                try {
                                    $FileName = ([io.path]::GetFileName("$imagePath")).toLower()
                                    $MimeType = Get-MimeType -Path "$imagePath"
                                    $UploadImage = Invoke-UploadNinjaOneFile -FileName $FileName -FilePath "$imagePath" -ContentType $MimeType
                                    $NewImageURL = "cid:$($UploadImage.contentId)"
                                    $ImgLink = $html.Links | Where-Object { $_.innerHTML -eq $imgHTML }
                                    Write-Host "Setting image to: $NewImageURL"
                                    $_.src = [string]$NewImageURL
                                    
                                    # Update Links for this image
                                    $ImgLink.href = [string]$NewImageUrl

                                } catch {
                                    $ManualLog = [PSCustomObject]@{
                                        Document_Name = $Article.Name
                                        Asset_Type    = "Article"
                                        Company_Name  = $Article.Company.CompanyName
                                        NinjaOneID    = $Article.NinjaOneID
                                        Notes         = 'Failed to Upload to NinjaOne'
                                        Action        = "$imagePath failed to upload to NinjaOne with error $_"
                                        Data          = "$InFile"
                                        NinjaOne_URL  = "https://$($NinjaOneBaseDomain)/#/systemDashboard/knowledgeBase/$($Article.NinjaOneID)/file"
                                        ITG_URL       = "$ITGURL/$($Article.ITGLocator)"
                                    }

                                    $null = $ManualActions.add($ManualLog)
                                }

                                if ($Magick -and $MovedItem) {
                                    Move-Item -Path $imagePath -Destination $OriginalFullImagePath
                                }
        
                            } else {

                                $ManualLog = [PSCustomObject]@{
                                    Document_Name = $Article.Name
                                    Asset_Type    = "Article"
                                    Company_Name  = $Article.Company.CompanyName
                                    NinjaOneID    = $Article.NinjaOneID
                                    Notes         = 'Image Not Detected'
                                    Action        = "$imagePath not detected as image, validate the identified file is an image, or imagemagick modules are loaded"        
                                    Data          = "$InFile"
                                    NinjaOne_URL  = "https://$($NinjaOneBaseDomain)/#/systemDashboard/knowledgeBase/$($Article.NinjaOneID)/file"
                                    ITG_URL       = "$ITGURL/$($Article.ITGLocator)"
                                }

                                $null = $ManualActions.add($ManualLog)

                            }
                        } else {
                            Write-Warning "Image $tnImgUrl file is missing"
                            $ManualLog = [PSCustomObject]@{
                                Document_Name = $Article.Name
                                Asset_Type    = "Article"
                                Company_Name  = $Article.Company.CompanyName
                                Field_Name    = 'N/A'
                                NinjaOneID    = $Article.NinjaOneID
                                Notes         = 'Image File Missing'
                                Action        = "$tnImgUrl is not present in export,validate the image exists in ITGlue and manually replace in NinjaOne"   
                                Data          = "$InFile"
                                NinjaOne_URL  = "https://$($NinjaOneBaseDomain)/#/systemDashboard/knowledgeBase/$($Article.NinjaOneID)/file"
                                ITG_URL       = "$ITGURL/$($Article.ITGLocator)"
                            }

                            $null = $ManualActions.add($ManualLog)
                        }
                    }
                }
            
                $page_Source = $html.documentelement.outerhtml
                $page_out = [regex]::replace($page_Source , '\xa0+', ' ')
                        
            }
        
            if ($page_out -eq '') {
                $page_out = 'Empty Document in IT Glue Export - Please Check IT Glue'
                $ManualLog = [PSCustomObject]@{
                    Document_Name = $Article.name
                    Asset_Type    = 'Article'
                    Company_Name  = $Article.Company.CompanyName
                    Field_Name    = 'N/A'
                    NinjaOneID    = $Article.NinjaOneID                    
                    Notes         = 'Empty Document'
                    Action        = 'Validate the document is blank in ITGlue, or manually copy the content across. Note that embedded documents in ITGlue will be migrated in blank with an attachment of the original doc'
                    Data          = "$InFile"
                    NinjaOne_URL  = "https://$($NinjaOneBaseDomain)/#/systemDashboard/knowledgeBase/$($Article.NinjaOneID)/file"
                    ITG_URL       = "$ITGURL/$($Article.ITGLocator)"
                }

                $null = $ManualActions.add($ManualLog)
            }
			
				
            $ArticleUpdate = @{
                id      = $Article.NinjaOneID
                name    = $Article.Name
                content = @{
                    html = ($page_out -replace '<HTML><HEAD></HEAD>\r\n<BODY>\r\n', '') -replace '</BODY></HTML>', ''
                }
            }

            try {
                $UpdatedDoc = Invoke-NinjaOneRequest -InputObject $ArticleUpdate -Method PATCH -Path 'knowledgebase/articles' -AsArray -ea Stop
                $Article.Imported = "Created-By-Script"
                $Article.NinjaOneObject = $UpdatedDoc
            } catch {
                Write-Error "Creation Failed $_"
                $Article.Imported = "Failed"
            }

            $ArticlesToUpdate.add($ArticleUpdate)            
			
        } 

        $ArticlesToUpdate | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\ArticlesToUpdate.json"


        Read-Host 'Articles Complete'


        $MatchedArticles | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\Articles.json"
        $ArticleErrors | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\ArticleErrors.json"
        $ManualActions | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\ManualActions.json"
        Write-TimedMessage -Timeout 3 -Message "Snapshot Point: Articles Created Continue?" -DefaultResponse "continue to Passwords, please."

    }

}



############################### Passwords ###############################


#Check for Passwords Resume
if ($ResumeFound -eq $true -and (Test-Path "$MigrationLogs\Passwords.json")) {
    Write-Host "Loading Previous Paswords Migration"
    $MatchedPasswords = Get-Content "$MigrationLogs\Passwords.json" -raw | Out-String | ConvertFrom-Json
} else {

    #Import Passwords
    Write-Host "Fetching Passwords from IT Glue" -ForegroundColor Green
    $PasswordSelect = { (Get-ITGluePasswords -page_size 1000 -page_number $i).data }
    $ITGPasswords = Import-ITGlueItems -ItemSelect $PasswordSelect -MigrationName 'Passwords'
    if ($ScopedMigration) {
        $OriginalPasswordsCount = $($ITGPasswords.count)
        Write-Host "Setting passwords to those in scope..." -foregroundcolor Yellow        
        $ITGPasswords = $ITGPasswords | Where-Object { $ScopedITGCompanyIds -contains $_.attributes.'organization-id' }
        Write-Host "Passwords scoped... $OriginalPasswordsCount => $($ITGPasswords.count)"
    }
    
    try {
        Write-Host "Loading Passwords from CSV for faster import" -foregroundcolor Cyan
        $ITGPasswordsRaw = Import-CSV -Path "$ITGLueExportPath\passwords.csv"
    } catch {
        $ITGPasswordsSingle = foreach ($ITGRawPass in $ITGPasswords) {
            $ITGPassword = (Get-ITGluePasswords -id $ITGRawPass.id -include related_items).data
            $ITGPassword
        }
        $ITGPasswords = $ITGPasswordsSingle
    }
    
    Write-Host "$($ITGPasswords.count) IT Glue Passwords Found"

    $PasswordsInCSV = [System.Collections.ArrayList]::new()
    $PasswordsNotInCSV = [System.Collections.ArrayList]::new()

    $IdOrganizationMap = @{}
    foreach ($row in $ITGPasswordsRaw) {
        $IdOrganizationMap[[string]$row.id] = @{
            'password'   = $row.password
            'otp_secret' = $row.otp_secret
        }
    }

    foreach ($row in $ITGPasswords) {
        if ($IdOrganizationMap.ContainsKey([string]$row.id) -eq $true) {
            $row.attributes | Add-Member -MemberType 'NoteProperty' -Name 'password' -Value $IdOrganizationMap[[string]$row.id].password
            $row.attributes | Add-Member -MemberType 'NoteProperty' -Name 'otp_secret' -Value $IdOrganizationMap[[string]$row.id].otp_secret
            [void]$PasswordsInCSV.Add($row)
        } else {
            [void]$PasswordsNotInCSV.Add($row)
        }
    }

    $MatchedPasswords = New-Object 'System.Collections.ArrayList'
    foreach ($itgpassword in $PasswordsInCSV) {
        [void]$MatchedPasswords.Add(
            [PSCustomObject]@{
                "Name"           = $itgpassword.attributes.name
                "ITGID"          = $itgpassword.id
                "NinjaOneID"     = ""
                "Matched"        = $false
                "NinjaOneObject" = ""
                "ITGObject"      = $itgpassword
                "Imported"       = ""
            }
        )
    }
    foreach ($itgpassword in $PasswordsNotInCSV) {
        $FullPassword = (Get-ITGluePasswords -id $itgpassword.id -include related_items).data
        [void]$MatchedPasswords.Add(
            [PSCustomObject]@{
                "Name"           = $itgpassword.attributes.name
                "ITGID"          = $itgpassword.id
                "NinjaOneID"     = ""
                "Matched"        = $false
                "NinjaOneObject" = ""
                "ITGObject"      = $FullPassword
                "Imported"       = ""
            }
        )
    }

    Write-Host "Passwords to Migrate"
    $MatchedPasswords | Sort-Object Name |  Select-Object Name | Format-Table


    $UnmappedPasswordCount = ($MatchedPasswords | Where-Object { $_.Matched -eq $false } | measure-object).count

    if ($ImportPasswords -eq $true -and $UnmappedPasswordCount -gt 0) {

        $importOption = Get-ImportMode -ImportName "Passwords"

        if (($importOption -eq "A") -or ($importOption -eq "S") ) {		

            $PasswordCategories = ($MatchedPasswords.ITGObject.attributes.'password-category-name' | Select-Object -Unique | ForEach-Object {
                    @{name = $_ }
                }) ?? @()

            $PasswordTemplateBody = [PSCustomObject]@{
                name          = "$($FlexibleLayoutPrefix)$($PasswordImportAssetLayoutName)"
                allowMultiple = $true
                mandatory     = $false
                fields        = @(
                    [PSCustomObject]@{
                        fieldLabel                = 'Category'
                        fieldName                 = 'category'
                        fieldType                 = 'DROPDOWN'
                        fieldTechnicianPermission = 'EDITABLE'
                        fieldScriptPermission     = 'NONE'
                        fieldApiPermission        = 'READ_WRITE'
                        fieldContent              = @{
                            values   = $PasswordCategories
                            required = $false
                        }
                    },
                    [PSCustomObject]@{
                        fieldLabel                = 'Username'
                        fieldName                 = 'username'
                        fieldType                 = 'TEXT'
                        fieldTechnicianPermission = 'EDITABLE'
                        fieldScriptPermission     = 'NONE'
                        fieldApiPermission        = 'READ_WRITE'
                        fieldContent              = @{
                            required = $false
                        }
                    },
                    [PSCustomObject]@{
                        fieldLabel                = 'Password'
                        fieldName                 = 'password'
                        fieldType                 = 'TEXT_ENCRYPTED'
                        fieldTechnicianPermission = 'EDITABLE'
                        fieldScriptPermission     = 'NONE'
                        fieldApiPermission        = 'READ_WRITE'
                        fieldContent              = @{
                            required         = $false
                            advancedSettings = @{
                                maxCharacters   = 10000
                                complexityRules = @{
                                    mustContainOneInteger           = $false
                                    mustContainOneLowercaseLetter   = $false
                                    mustContainOneUppercaseLetter   = $false
                                    greaterOrEqualThanSixCharacters = $false
                                }
                            }
                        }
                    },
                    [PSCustomObject]@{
                        fieldLabel                = 'OTP Secret'
                        fieldName                 = 'otpSecret'
                        fieldType                 = 'TEXT_ENCRYPTED'
                        fieldTechnicianPermission = 'EDITABLE'
                        fieldScriptPermission     = 'NONE'
                        fieldApiPermission        = 'READ_WRITE'
                        fieldContent              = @{
                            required         = $false
                            advancedSettings = @{
                                maxCharacters   = 10000
                                complexityRules = @{
                                    mustContainOneInteger           = $false
                                    mustContainOneLowercaseLetter   = $false
                                    mustContainOneUppercaseLetter   = $false
                                    greaterOrEqualThanSixCharacters = $false
                                }
                            }
                        }
                    },
                    [PSCustomObject]@{
                        fieldLabel                = 'URL'
                        fieldName                 = 'url'
                        fieldType                 = 'URL'
                        fieldTechnicianPermission = 'EDITABLE'
                        fieldScriptPermission     = 'NONE'
                        fieldApiPermission        = 'READ_WRITE'
                        fieldContent              = @{
                            required = $false
                        }
                    }
                )
            }

            $PasswordTemplate = Invoke-NinjaOneDocumentTemplate -Template $PasswordTemplateBody

            foreach ($company in $CompaniesToMigrate) {
                Write-Host "Migrating $($company.CompanyName)" -ForegroundColor Green

                foreach ($unmatchedPassword in ($MatchedPasswords | Where-Object { $_.Matched -eq $false -and $company.ITGCompanyObject.id -eq $_."ITGObject".attributes."organization-id" })) {

                    Confirm-Import -ImportObjectName "$($unmatchedPassword.Name)" -ImportObject $unmatchedPassword -ImportSetting $ImportOption

                    Write-Host "Starting $($unmatchedPassword.Name)"

                    $ParentItemID = $null
		    
                    if ($($unmatchedPassword.ITGObject.attributes."resource-id")) {
						
                        if ($unmatchedPassword.ITGObject.attributes."resource-type" -eq "flexible-asset-traits") {
                            # Check if it has already migrated with Assets
                            $FoundItem = $MatchedAssetPasswords | Where-Object { $_.ITGID -eq $($unmatchedPassword.ITGID) }
                            if (!$FoundItem) {
                                Write-Host "Could not find password field on asset. ParentID: $($unmatchedPassword.ITGObject.attributes.`"resource-id`")"
                                $FoundItem = $MatchedAssets | Where-Object { $_.ITGID -eq $unmatchedPassword.ITGObject.attributes."resource-id" }
                                $ManualLog = [PSCustomObject]@{
                                    Document_Name = $FoundItem.name
                                    Field_Name    = $unmatchedPassword.ITGObject.attributes.name
                                    Asset_Type    = "Asset password field"
                                    Company_Name  = $unmatchedPassword.ITGObject."organization-name"
                                    NinjaOneID    = $unmatchedPassword.NinjaOneID
                                    Notes         = "Password from FA Field not found."
                                    Action        = "Manually create password"
                                    Data          = "Type: $($unmatchedPassword.ITGObject.attributes.`"resource-type`")"
                                    NinjaOne_URL  = ""
                                    ITG_URL       = $unmatchedPassword.ITGObject.attributes."parent-url"
                                }
                                $null = $ManualActions.add($ManualLog)
                            } else {
                                Write-Host "Migrated with Asset: $($FoundItem.NinjaOneID)"
                            }
                        } else {
                            # Check if it needs to link to websites
                            if ($($unmatchedPassword.ITGObject.attributes."resource-type") -eq "domains") {
                                $ParentItemID = ($MatchedDomains | Where-Object { $_.ITGID -eq $($unmatchedPassword.ITGObject.attributes."resource-id") }).NinjaOneID
                                if ($ParentItemID) {
                                    Write-Host "Matched to $ParentItemID" -ForegroundColor Green
                                } else {
                                    Write-Host "Could not find asset to Match. ParentID: $($unmatchedPassword.ITGObject.attributes.`"resource-id`")"
                                    $ManualLog = [PSCustomObject]@{
                                        Document_Name = $unmatchedPassword.ITGObject.attributes.name
                                        Field_Name    = "N/A"
                                        Asset_Type    = 'Domains'
                                        Company_Name  = $unmatchedPassword.ITGObject.attributes.'organization-name'
                                        NinjaOneID    = $unmatchedPassword.NinjaOneID
                                        Notes         = "Password could not be related."
                                        Action        = "Manually relate password"
                                        Data          = "Type: $($unmatchedPassword.ITGObject.attributes.`"resource-type`")"
                                        NinjaOne_URL  = ""
                                        ITG_URL       = $unmatchedPassword.ITGObject.attributes."parent-url"
                                    }
                                    $null = $ManualActions.add($ManualLog)
                                }

                            } else {
                                Write-Host "Could not find asset to Match. ParentID: $($unmatchedPassword.ITGObject.attributes.`"resource-id`")"
                                $ManualLog = [PSCustomObject]@{
                                    Document_Name = $unmatchedPassword.ITGObject.attributes.name
                                    Field_Name    = "N/A"
                                    Asset_Type    = "$($unmatchedPassword.ITGObject.attributes.'resource-type')"
                                    Company_Name  = "$($unmatchedPassword.ITGObject.attributes.'organization-name')"
                                    NinjaOneID    = $unmatchedPassword.NinjaOneID
                                    Notes         = "Password could not be related."
                                    Action        = "Manually relate password"
                                    Data          = "Type: $($unmatchedPassword.ITGObject.attributes.`"resource-type`")"
                                    NinjaOne_URL  = ""
                                    ITG_URL       = $unmatchedPassword.ITGObject.attributes."parent-url"
                                }
                                $null = $ManualActions.add($ManualLog)                              
                            }
                        }
                    }
					
                    if (!($($unmatchedPassword.ITGObject.attributes."resource-type") -eq "flexible-asset-traits")) {
						
                        $validated_otp = "$($unmatchedPassword.ITGObject.attributes.otp_secret)".Trim().ToUpper()
                        if ($validated_otp) {
                            $isValidBase32 = $validated_otp -match '^[A-Z2-7]+$'
                            $lengthOK = $validated_otp.Length -ge 16 -and $validated_otp.Length -le 80

                            $validated_otp = if ($isValidBase32 -and $lengthOK) { $validated_otp } else { $null }

                            if (-not ($isValidBase32 -and $lengthOK)) {
                                Write-Warning "Invalid OTP secret for $($unmatchedPassword.ITGObject.attributes.name): $($unmatchedPassword.ITGObject.attributes.otp_secret)... valid base32? $isValidBase32 length ok? $lengthOK (min / max is 16 / 80 chars)"
                            }                            
                        }


                        $PasswordCreate = @{
                            documentName        = "$($unmatchedPassword.ITGObject.attributes.name)"
                            organizationId      = $company.NinjaOneID
                            documentDescription = $unmatchedPassword.ITGObject.attributes.notes
                            documentTemplateId  = $PasswordTemplate.id
                            fields              = @{
                                category  = $unmatchedPassword.ITGObject.attributes.'password-category-name'
                                password  = $unmatchedPassword.ITGObject.attributes.password
                                url       = Convert-ToHttpsUrl($(if ($url = $unmatchedPassword.ITGObject.attributes.url) { $url } Else { $unmatchedPassword.ITGObject.attributes.'resource-url' }))
                                username  = $unmatchedPassword.ITGObject.attributes.username
                                otpSecret = $validated_otp
                            }
                        }

                        try {
                            $NewPassword = Invoke-NinjaOneRequest -InputObject $PasswordCreate -Method POST -Path 'organization/documents' -AsArray -ea Stop 
                            $unmatchedPassword.matched = $true
                            $unmatchedPassword.NinjaOneID = $NewPassword.documentId
                            $unmatchedPassword.NinjaOneObject = $NewPassword
                            $unmatchedPassword.Imported = "Created-By-Script"
                            $ImportsMigrated = $ImportsMigrated + 1
                            Write-host "$($NewPassword.documentName) Has been created in NinjaOne"
                        } catch {
                            Write-Error "Error creating password: $($unmatchedPassword.ITGObject.attributes.name) $_"
                            Write-Host "$($PasswordCreate | ConvertTo-Json -Depth 100)"
                            $unmatchedPassword.Imported = "Failed"
                        }

                        
                    }
                }
            
            }

        }
    } else {
        if ($UnmappedPasswordCount -eq 0) {
            Write-Host "All Passwords matched, no migration required" -foregroundcolor green
        } else {
            Write-Host "Warning Import passwords is set to disabled so the above unmatched passwords will not have data migrated" -foregroundcolor red
            Write-TimedMessage -Timeout 3 -Message "Press any key to continue or CTRL+C to quit"  -DefaultResponse "continue wrap-up of passwords, please."
        }
    }

    # Save the results to resume from if needed
    $MatchedPasswords | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\Passwords.json"
    $ManualActions | ConvertTo-Json -depth 100 | Out-File "$MigrationLogs\ManualActions.json"
    Write-TimedMessage -Timeout 3 -Message "Snapshot Point: Passwords Finished. Continue?"  -DefaultResponse "continue to Document/Article Updates, please."
}

############################## Update ITGlue URLs on All Areas to NinjaOne #######################
# Fetch KB Articles
Write-Host "Updating KB Articles and Apps and Services Documents with IT Glue URLs"
[System.Collections.Generic.List[PSCustomObject]]$AllKBArticles = @()
[System.Collections.Generic.List[PSCustomObject]]$OrgNinjaKBArticles = Invoke-NinjaOneRequest -Method GET -Path 'knowledgebase/organization/articles'
[System.Collections.Generic.List[PSCustomObject]]$GlobalNinjaKBArticles = Invoke-NinjaOneRequest -Method GET -Path 'knowledgebase/global/articles'

$AllKBArticles = Invoke-AddToArray -InputArray $AllKBArticles -AppendItem $OrgNinjaKBArticles
$AllKBArticles = Invoke-AddToArray -InputArray $AllKBArticles -AppendItem $GlobalNinjaKBArticles

$AllKBArticles = $AllKBArticles | Where-Object { $_.content.html -like "*$ITGURL*" }

# Fetch Assets
$UpdateDocuments = Invoke-NinjaOneRequest -Method GET -Path 'organization/documents' | Where-Object { $_.fields.value -like "*$ITGURL*" }

# Articles
[System.Collections.Generic.List[PSCustomObject]]$articlesUpdated = @()
foreach ($articleFound in $AllKBArticles) {
    if ($NewContent = Update-StringWithCaptureGroups -inputString $articleFound.content.html -pattern $RichRegexPatternToMatchSansAssets -type "rich") {
        $NewContent = Update-StringWithCaptureGroups -inputString $NewContent -pattern $RichRegexPatternToMatchWithAssets -type "rich"
        $NewContent = Update-StringWithCaptureGroups -inputString $NewContent -pattern $RichDocLocatorUrlPatternToMatch -type "rich"
        $NewContent = Update-StringWithCaptureGroups -inputString $NewContent -pattern $RichDocLocatorRelativeURLPatternToMatch -type "rich"
        Write-Host "Updating Article $($articleFound.name) with replaced Content" -ForegroundColor 'Green'

        $ArticleUpdate = @{
            id      = $articleFound.id
            name    = $articleFound.Name
            content = @{
                html = $NewContent
            }
        }

        try {
            $UpdatedDoc = Invoke-NinjaOneRequest -InputObject $ArticleUpdate -Method PATCH -Path 'knowledgebase/articles' -AsArray -ea Stop
            $articlesUpdated.add(@{"status" = "replaced"; "original_article" = $articleFound; "updated_article" = $UpdatedDoc })
        } catch {
            Write-Error "Creation Failed $_"
            $articlesUpdated.add(@{"status" = "failed"; "original_article" = $articleFound; "attempted_changes" = $newContent })
        }

    } else {
        Write-Warning "Article $articleFound.id found ITGlue URL but didn't match"
        $articlesUpdated = $articlesUpdated.add(@{"status" = "clean"; "original_article" = $articleFound })
    }
}

$articlesUpdated | ConvertTo-Json -depth 100 | Out-file "$MigrationLogs\ReplacedArticlesURL.json"
Write-TimedMessage -Timeout 3 -Message "Snapshot Point: Article URLs Replaced. Continue?"  -DefaultResponse "continue to Assets, please."

# Apps and Services Documents
[System.Collections.Generic.List[PSCustomObject]]$documentsUpdated = @()
foreach ($documentFound in $UpdateDocuments) {
    $replacedStatus = 'clean'
    [PSCustomObject]$UpdateFields = @{}

    foreach ($field in $documentFound.fields) {
        $label = $field.name

        if ($label -ne 'itgUrl' -and $field.value -like "*$ITGURL*") {
            if ($field.value.html) {
                $NewContent = Update-StringWithCaptureGroups -inputString $field.value.html -pattern $RichRegexPatternToMatchSansAssets -type "rich"
                $NewContent = Update-StringWithCaptureGroups -inputString $NewContent -pattern $RichRegexPatternToMatchWithAssets -type "rich"
            } else {
                $NewContent = Update-StringWithCaptureGroups -inputString $field.value -pattern $RichRegexPatternToMatchSansAssets -type "rich"
                $NewContent = Update-StringWithCaptureGroups -inputString $NewContent -pattern $RichRegexPatternToMatchWithAssets -type "rich"  
            }

            if ($NewContent -and $NewContent -ne $field.value) {
                Write-Host "Replacing Document $($documentFound.documentName) field $($field.name) with updated content"
                if ($field.value.html) {
                    $UpdateFields | Add-Member -NotePropertyName $label -NotePropertyValue @{html = $NewContent }
                } else {
                    $UpdateFields | Add-Member -NotePropertyName $label -NotePropertyValue $NewContent
                }
                $replacedStatus = 'replaced'
            }
        }
    }

    if ($replacedStatus -eq 'replaced') {
        Write-Host "Updating Document $($documentFound.documentName) with new custom_fields array" -ForegroundColor 'Green'
        try {
            $DocUpdate = @{
                documentId   = $documentFound.documentId
                documentName = $documentFound.documentName
                fields       = $UpdateFields
            }
            $null = Invoke-NinjaOneRequest -InputObject $DocUpdate -Method PATCH -Path 'organization/documents' -AsArray -ea Stop
        } catch {
            Write-Error "Failed to update Document $_"
            Write-Host "$($DocUpdate | ConvertTo-Json)"
        }
    }

    $documentsUpdated.add($documentFound)
}
$documentsUpdated | ConvertTo-Json -depth 100 | Out-file "$MigrationLogs\ReplacedDocumentsURL.json"
Write-TimedMessage -Timeout 3 -Message  "Snapshot Point: Document URLs Replaced. Continue?" -DefaultResponse "continue to Passwords Matching, please."



############################### Generate Manual Actions Report ###############################

$Head = @"
<html>
<head>
<Title>Manual Actions Required Report</Title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.1/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-+0n0xVW2eSR5OomGNYDnhzAbDsOXxcvSN1TPprVMTNDbiYZCxYbOOl7+AMvyTG2x" crossorigin="anonymous">
<style type="text/css">
<!-
body {
    font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
}
h2{ clear: both; font-size: 100%;color:#354B5E; }
h3{
    clear: both;
    font-size: 75%;
    margin-left: 20px;
    margin-top: 30px;
    color:#475F77;
}
table{
	border-collapse: collapse;
	margin: 5px 0;
	font-size: 0.8em;
	font-family: sans-serif;
	min-width: 400px;
	box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);
}

th, td {
	padding: 5px 5px;
	max-width: 400px;
	width:auto;
}
thead tr {
	background-color: #009879;
	color: #ffffff;
	text-align: left;
}
tr {
	border-bottom: 1px solid #dddddd;
}
tr:nth-of-type(even) {
	background-color: #f3f3f3;
}
->
</style>
</head>
<body>
<div style="padding:40px">


"@


$MigrationReport = @"
<h1> Migration Report </h1>
Started At: $ScriptStartTime <br />
Completed At: $(Get-Date -Format "o") <br />
$(($MatchedCompanies | Measure-Object).count) : Companies Migrated <br />
$(($MatchedLocations | Measure-Object).count) : Locations Migrated <br />
$(($MatchedDomains | Measure-Object).count) : Websites Migrated <br />
$(($MatchedDevices | Measure-Object).count) : Configurations Migrated <br />
$(($MatchedContacts | Measure-Object).count) : Contacts Migrated <br />
$(($MatchedLayouts | Measure-Object).count) : Layouts Migrated <br />
$(($MatchedAssets | Measure-Object).count) : Assets Migrated <br />
$(($MatchedArticles | Measure-Object).count) : Articles Migrated <br />
$(($MatchedPasswords | Measure-Object).count) : Passwords Migrated <br />
<hr>
<h1>Manual Actions Required Report</h1>
"@

$footer = "</div></body></html>"

$UniqueItems = $ManualActions | Select-Object ninjaoneid, ninjaone_url -unique

$ManualActionsReport = foreach ($item in $UniqueItems) {
    $items = $ManualActions | where-object { $_.NinjaOneid -eq $item.ninjaoneid -and $_.NinjaOne_url -eq $item.NinjaOne_url }
    $core_item = $items | Select-Object -First 1
    $Header = "<h2><strong>Name: $($core_item.Document_Name)</strong></h2>
				<h2>Type: $($core_item.Asset_Type)</h2>
				<h2>Company: $($core_item.Company_name)</h2>
				<h2>NinjaOne URL: <a href=$($core_item.NinjaOne_URL)>$($core_item.NinjaOne_URL)</a></h2>
				<h2>IT Glue URL: <a href=$($core_item.ITG_URL)>$($core_item.ITG_URL)</a></h2>
				"
    $Actions = $items | Select-Object Field_Name, Notes, Action, Data | ConvertTo-Html -fragment | Out-String

    $OutHTML = "$Header $Actions <hr>"

    $OutHTML

}

$FinalHtml = "$Head $MigrationReport $ManualActionsReport $footer"
$FinalHtml | Out-File ManualActions.html



############################### End ###############################


Write-Host "#######################################################" -ForegroundColor Green
Write-Host "#                                                     #" -ForegroundColor Green
Write-Host "#        IT Glue to NinjaOne Migration Complete           #" -ForegroundColor Green
Write-Host "#                                                     #" -ForegroundColor Green
Write-Host "#######################################################" -ForegroundColor Green
Write-Host "Started At: $ScriptStartTime"
Write-Host "Completed At: $(Get-Date -Format "o")"
Write-Host "$(($MatchedCompanies | Measure-Object).count) : Companies Migrated" -ForegroundColor Green
Write-Host "$(($MatchedLocations | Measure-Object).count) : Locations Migrated" -ForegroundColor Green
Write-Host "$(($MatchedDomains | Measure-Object).count) : Websites Migrated" -ForegroundColor Green
Write-Host "$(($MatchedDevices | Measure-Object).count) : Configurations Migrated" -ForegroundColor Green
Write-Host "$(($MatchedContacts | Measure-Object).count) : Contacts Migrated" -ForegroundColor Green
Write-Host "$(($MatchedLayouts | Measure-Object).count) : Layouts Migrated" -ForegroundColor Green
Write-Host "$(($MatchedAssets | Measure-Object).count) : Assets Migrated" -ForegroundColor Green
Write-Host "$(($MatchedArticles | Measure-Object).count) : Articles Migrated" -ForegroundColor Green
Write-Host "$(($MatchedPasswords | Measure-Object).count) : Passwords Migrated" -ForegroundColor Green
Write-Host "#######################################################" -ForegroundColor Green
Write-Host "Manual Actions report can be found in ManualActions.html in the folder the script was run from"
Write-Host "Logs of what was migrated can be found in the MigrationLogs folder"
Write-TimedMessage -Message "Press any key to view the manual actions report or Ctrl+C to end" -Timeout 120  -DefaultResponse "continue, view generative Manual Actions webpage, please."

Start-Process ManualActions.html
